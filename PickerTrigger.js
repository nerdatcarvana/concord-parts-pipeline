/**
 * PickerTrigger.gs
 * Installable onEdit trigger. Fires on every edit to the Processing board.
 *
 * Two entry points:
 *   1. Stock Number edited → walk-in intake via Intake.handleWalkIn()
 *   2. Bin Location edited → bin parse, PICKED×n + COMPLETED (or DISCARDED)
 *
 * INLINE PROMOTION (bin path): if the bin is entered before the queued
 * walk-in trigger has stamped a Request ID, and the row already has both
 * Requester and Build, the bin execution promotes the row itself via the
 * same idempotent Intake.handleWalkIn() and proceeds straight into the
 * pick — no more wait-and-retype loop against trigger-dispatch latency.
 * Rows genuinely missing Requester/Build still reject-and-clear.
 * NOTE: walk-ins promoted at bin time have requested_ts ≈ complete_ts, so
 * their cycle_time_min reads near zero — which is the honest number for a
 * request fulfilled on the spot (the old nonzero value was dispatch jitter).
 *
 * Runs as the script OWNER (installable trigger), so it has write access to
 * all protected sheets even though the picker who triggered it does not.
 *
 * Top-level entry: onEditTrigger(e)
 */

// --------------------------------------------------------------------------
// onEditTrigger(e)
// --------------------------------------------------------------------------
function onEditTrigger(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();

  // Gate on the stable sheet gid, not the tab name. A picker renaming the
  // Processing tab (a real production incident) must NOT silently break the pick
  // flow; resolving by gid also self-heals the tab name back to canonical.
  const proc = Sheets.get('PROCESSING', ss);
  if (!proc || range.getSheet().getSheetId() !== proc.getSheetId()) return;
  const sheet = proc;

  const rowNum = range.getRow();
  if (rowNum < 2) return;

  const H = CONFIG.HEADERS;
  const E = CONFIG.ENUMS;

  let cols;
  try {
    cols = Utils.getColMap(
      sheet,
      H.PROC_STOCK,
      H.PROC_BIN,
      H.PROC_REQ_ID,
      H.PROC_REQUESTER,
      H.PROC_BUILD,
      H.PROC_LOT,
      H.PROC_REQUESTED,
      H.PROC_NOTES,
    );
  } catch (err) {
    ss.toast('⚠️ Processing sheet headers not found. Run Setup.', '⚠️ Error', 5);
    return;
  }

  const editedCol = range.getColumn();

  // --------------------------------------------------------------------------
  // PATH 1 — Stock Number or Requester edited: walk-in intake
  //
  // Both columns funnel through handleWalkIn so the REQUESTED event is only
  // emitted once the row has BOTH a stock number AND a requester.
  // Pickers often paste the stock number first and fill the requester after —
  // firing on stock-only would lock in an empty requester permanently.
  // --------------------------------------------------------------------------
  if (editedCol === cols[H.PROC_STOCK] ||
      editedCol === cols[H.PROC_REQUESTER] ||
      editedCol === cols[H.PROC_BUILD]) {
    // Only proceed if there's at least a stock number on the row
    const stockVal = String(sheet.getRange(rowNum, cols[H.PROC_STOCK]).getValue()).trim();
    if (!stockVal) return;
    try {
      // handleWalkIn manages its own ScriptLock around the read-decide-emit
      // critical section; do NOT wrap it here. The earlier lot-autofill and the
      // incremental-field guards run lock-free by design, so a call-site lock
      // would needlessly serialize (and toast 'busy' on) those cheap paths.
      Intake.handleWalkIn(sheet, rowNum);
    } catch (err) {
      if (err.message.includes('could not acquire lock')) {
        ss.toast('System busy. Please try again.', '⏱️ Busy', 5);
      } else {
        ss.toast('Walk-in error: ' + err.message, '❌ Error', 6);
      }
    }
    return;
  }

  // --------------------------------------------------------------------------
  // PATH 2 — Bin Location edited: pick gate
  // --------------------------------------------------------------------------
  if (editedCol === cols[H.PROC_BIN]) {

    // --- Bug 1 fix: detect Sheets formula errors before reading the value ---
    // When a picker types something Sheets interprets as a bad formula (e.g. a
    // bin code starting with '='), getValue() returns a Sheets ErrorValue object.
    // Catch it here, toast, clear the cell, and bail — produce no event.
    const binCell = sheet.getRange(rowNum, cols[H.PROC_BIN]);
    if (binCell.isBlank()) return;

    let rawBin;
    try {
      const cellVal = binCell.getValue();
      // SpreadsheetApp ErrorValue objects have a getErrorCode() method.
      // Any non-string, non-number value coming back here is an error.
      if (cellVal === null || cellVal === undefined) return;
      if (typeof cellVal === 'object') {
        // Sheets returned an error object (formula error in cell)
        ss.toast(
          'Bin entry caused a formula error. Make sure it doesn\'t start with = or contain special characters.',
          '🚫 Invalid Entry',
          7
        );
        binCell.clearContent();
        return;
      }
      rawBin = String(cellVal).trim();
    } catch (err) {
      binCell.clearContent();
      return;
    }

    if (!rawBin) return;

    // Batch-read the entire row once. Individual getValue() calls per field were
    // each a separate round-trip; reading the whole row in one getValues() call
    // costs the same network latency as one of those individual reads.
    let rowData  = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    let reqId    = String(rowData[cols[H.PROC_REQ_ID]    - 1] || '').trim();
    const stockNum = String(rowData[cols[H.PROC_STOCK]     - 1] || '').trim();

    if (!reqId && !stockNum) {
      ss.toast('Missing Stock Number. Cannot process an empty row.', '⚠️ Warning', 5);
      binCell.clearContent();
      return;
    }

    // ------------------------------------------------------------------------
    // INLINE PROMOTION (was: hard reject). Stock present but no Request ID
    // means the walk-in promotion hasn't landed yet. Under the old guard that
    // was always a reject-and-clear — but installable-trigger dispatch runs
    // seconds behind the keyboard, so a picker who typed the row straight
    // through hit "Not Ready" and lost their bin entry to a race they couldn't
    // see. If the row already carries Requester AND Build, THIS execution can
    // promote it right now via the exact same idempotent path the build-edit
    // trigger uses: Intake.handleWalkIn takes its own ScriptLock and its
    // (stock, requester, OPEN) probe against _Requests adopts any request a
    // racing execution already minted — no new concurrency surface. Called
    // BEFORE the pick's withLock below: GAS locks are not reentrant, so the
    // two critical sections must be sequential, never nested. After promotion
    // the row is re-read so reqId and Requested-ts are fresh for the pick.
    // Only a genuinely incomplete row (Requester/Build still blank) rejects.
    // ------------------------------------------------------------------------
    if (!reqId) {
      const hasRequester = String(rowData[cols[H.PROC_REQUESTER] - 1] || '').trim() !== '';
      const hasBuild     = String(rowData[cols[H.PROC_BUILD]     - 1] || '').trim() !== '';

      if (hasRequester && hasBuild) {
        try {
          Intake.handleWalkIn(sheet, rowNum);
        } catch (err) {
          if (err.message.includes('could not acquire lock')) {
            // Transient contention — leave the bin value in place and tell the
            // picker to re-commit it, same contract as the pick path's busy case.
            ss.toast('System busy. Click Bin Location and press Enter to retry.', '⏱️ Busy', 6);
          } else {
            ss.toast('Walk-in error: ' + err.message, '❌ Error', 6);
            binCell.clearContent();
          }
          return;
        }
        // Promotion committed durably (handleWalkIn flushes before returning) —
        // re-read the row so the pick sees the fresh Request ID and Requested ts.
        rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
        reqId   = String(rowData[cols[H.PROC_REQ_ID] - 1] || '').trim();
      }

      // Still no id: Requester/Build genuinely missing (or promotion bailed).
      // Reject and CLEAR — leaving a bin value on an un-promoted row would be a
      // silent "looks picked, never logged" state, which no toast can excuse.
      if (!reqId) {
        ss.toast(
          'Fill in Requester (and Build Load Level) first — this row has no Request ID yet.',
          '⚠️ Not Ready',
          6
        );
        binCell.clearContent();
        return;
      }
    }

    // N/A discard — resolved before the lock; no Ledger reads needed until inside.
    // Runs AFTER inline promotion, so N/A on a complete-but-unpromoted row now
    // produces the correct REQUESTED → DISCARDED pair instead of a reject.
    const normalizedBin = rawBin.toLowerCase().replace(/[\s\/\-\.]/g, '');
    if (normalizedBin === 'na') {
      _handleDiscard(sheet, rowNum, cols, rowData, reqId, stockNum, rawBin, ss, e);
      return;
    }

    // Normal pick path
    try {
      Utils.withLock(() => {
        _handlePick(sheet, rowNum, cols, rowData, reqId, stockNum, rawBin, ss, e);
      });
    } catch (err) {
      if (err.message.includes('could not acquire lock')) {
        ss.toast('System busy. Click Bin Location and press Enter to retry.', '⏱️ Busy', 6);
      } else {
        ss.toast('Pick error: ' + err.message, '❌ Error', 6);
      }
    }
  }
}

// --------------------------------------------------------------------------
// _handlePick — the core pick + complete flow
// --------------------------------------------------------------------------
function _handlePick(sheet, rowNum, cols, rowData, reqId, stockNum, rawBin, ss, e) {
  const H   = CONFIG.HEADERS;
  const E   = CONFIG.ENUMS;
  const now = Utils.isoNow();

  // Capture picker identity
  const pickerRaw   = (e && e.user && e.user.email) ? e.user.email
                      : Session.getActiveUser().getEmail();
  const pickerIdent = Identity.resolve(pickerRaw);
  const pickerEmail = Identity.isUnmapped(pickerIdent) ? pickerRaw : pickerIdent.email;
  const pickerDisp  = Identity.isUnmapped(pickerIdent) ? pickerRaw : pickerIdent.display;

  // Row context already batch-read by the caller — no additional getValue() calls.
  const rawBuild    = String(rowData[cols[H.PROC_BUILD]     - 1] || '').trim();
  const rawLot      = String(rowData[cols[H.PROC_LOT]       - 1] || '').trim();
  const requestedTs = String(rowData[cols[H.PROC_REQUESTED] - 1] || '').trim();

  const buildLevel = Utils.normalizeBuildLevel(rawBuild) || rawBuild;
  const lotStatus  = Utils.normalizeLotStatus(rawLot);

  // --- Parse the bin string ---
  const parseResult = _parseBins(rawBin);

  if (parseResult.rejected) {
    ss.toast(
      `Bin "${parseResult.rejectedToken}" looks like a stock number. Clear the cell and re-enter.`,
      '🚫 Invalid Bin',
      6
    );
    sheet.getRange(rowNum, cols[H.PROC_BIN]).clearContent();
    return;
  }

  const { bins } = parseResult;

  if (bins.length === 0) {
    ss.toast('No valid bin codes found. Check your entry.', '⚠️ Warning', 5);
    sheet.getRange(rowNum, cols[H.PROC_BIN]).clearContent();
    return;
  }

  // --- Reject entries where every token is UNKNOWN ---
  // A keyboard smash or completely unrecognized string produces bins, but all
  // with locationType UNKNOWN. That's not a real pick — block it, clear the
  // cell, and make the picker re-enter something recognizable.
  // Partial UNKNOWN (some known bins + some unknown) is still allowed through
  // with a warning toast — the picker may have a legitimate mixed entry.
  const allUnknown = bins.every(b => b.locationType === CONFIG.BIN.LOCATION_TYPE_UNKNOWN);
  if (allUnknown) {
    ss.toast(
      `"${rawBin}" doesn't match any known bin format. ` +
      `Expected a rack code (e.g. 5D5A, URGA4D), BULK (shorthand B1.2 works), HUB, DASH, or FL location. Re-enter or ask your lead.`,
      '🚫 Unrecognized Bin',
      8
    );
    sheet.getRange(rowNum, cols[H.PROC_BIN]).clearContent();
    return;
  }

  const allBinCodes = bins.map(b => b.code).join(', ');
  const allBinTypes = [...new Set(bins.map(b => b.locationType))].join(', ');

  // --- Build event array: PICKED×N + COMPLETED ---
  const events = [];

  for (const bin of bins) {
    events.push({
      event_id         : Utils.newEventId(),
      request_id       : reqId,
      event_type       : E.EVENT_PICKED,
      event_ts         : now,
      actor_email      : pickerEmail,
      actor_display    : pickerDisp,
      stock_number     : stockNum,
      build_load_level : buildLevel,
      lot_status       : lotStatus,
      bin_code         : bin.code,
      location_type    : bin.locationType,
      raw_input        : bin.rawToken,
    });
  }

  events.push({
    event_id         : Utils.newEventId(),
    request_id       : reqId,
    event_type       : E.EVENT_COMPLETED,
    event_ts         : now,
    actor_email      : pickerEmail,
    actor_display    : pickerDisp,
    stock_number     : stockNum,
    build_load_level : buildLevel,
    lot_status       : lotStatus,
    bin_code         : allBinCodes,
    location_type    : allBinTypes,
  });

  // --- Cycle time ---
  let cycleMin = '';
  if (requestedTs) {
    try {
      cycleMin = Math.round((new Date(now) - new Date(requestedTs)) / 60000 * 10) / 10;
    } catch(_) {}
  }

  // --- Batched ledger write: events + upsert in one commitPick call ---
  // commitPick resolves _EventLog and _Requests sheet refs once (cached),
  // reads _Requests once (shared between the completed-status guard and the
  // upsert write), and returns the pre-write status so the double-processing
  // guard below requires no additional Ledger.getRequest() round-trip.
  const priorStatus = Ledger.commitPick(reqId, events, {
    request_id     : reqId,
    bin_count      : bins.length,
    bins           : allBinCodes,
    first_pick_ts  : now,
    complete_ts    : now,
    picker_email   : pickerEmail,
    picker_display : pickerDisp,
    cycle_time_min : cycleMin,
    status         : E.STATUS_COMPLETED,
  });

  // Double-processing guard. commitPick now performs this check BEFORE writing,
  // so reaching here means nothing was appended to _EventLog and _Requests was
  // left untouched.
  //
  // A COMPLETED request still on the board means THIS row is a duplicate of one
  // that was already picked and cleared. Clear the whole row, not just the bin
  // cell — clearing one cell leaves a phantom that no longer has a path to
  // being cleared through the normal flow, so it sits on the board forever.
  if (priorStatus === E.STATUS_COMPLETED) {
    ss.toast('This request was already completed — removing duplicate row.', '⚠️ Duplicate Row', 5);
    sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).clearContent();
    SpreadsheetApp.flush();
    return;
  }

  // --- Toast for unknown location types ---
  const unknowns = bins.filter(b => b.locationType === CONFIG.BIN.LOCATION_TYPE_UNKNOWN);
  if (unknowns.length > 0) {
    ss.toast(
      `Unknown bin type(s): ${unknowns.map(b => b.code).join(', ')}. Logged — please review.`,
      '⚠️ Unknown Bin Type',
      6
    );
  }

  // --- Clear the row from the board ---
  sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).clearContent();
  SpreadsheetApp.flush();

  if (unknowns.length === 0) {
    ss.toast(`✅ Picked: ${allBinCodes}`, '✅ Complete', 4);
  }
}

// --------------------------------------------------------------------------
// _handleDiscard — N/A void path
// --------------------------------------------------------------------------
function _handleDiscard(sheet, rowNum, cols, rowData, reqId, stockNum, rawBin, ss, e) {
  const H   = CONFIG.HEADERS;
  const E   = CONFIG.ENUMS;
  const now = Utils.isoNow();

  // Row context already batch-read; pull fields from rowData directly.
  const notes      = String(rowData[cols[H.PROC_NOTES]  - 1] || '').trim();
  const rawBuild   = String(rowData[cols[H.PROC_BUILD]  - 1] || '').trim();
  const rawLot     = String(rowData[cols[H.PROC_LOT]    - 1] || '').trim();
  const buildLevel = Utils.normalizeBuildLevel(rawBuild) || rawBuild;
  const lotStatus  = Utils.normalizeLotStatus(rawLot);

  // Picker identity — who physically discarded the row
  const pickerRaw   = (e && e.user && e.user.email) ? e.user.email
                      : Session.getActiveUser().getEmail();
  const pickerIdent = Identity.resolve(pickerRaw);
  const pickerEmail = Identity.isUnmapped(pickerIdent) ? pickerRaw : pickerIdent.email;
  const pickerDisp  = Identity.isUnmapped(pickerIdent) ? pickerRaw : pickerIdent.display;

  if (reqId) {
    try {
      Utils.withLock(() => {
        // Terminal-status guard. `status` is NOT in _mergeRequestRow's IMMUTABLE
        // set, so discarding a request that is already COMPLETED would flip it
        // to DISCARDED and overwrite complete_ts, picker and notes — turning
        // fulfilled work into discarded work in every report. This is reachable
        // whenever a duplicate board row is marked N/A after its twin was
        // picked, so refuse the write and let the row clear below.
        const prior = Ledger.statusOf(reqId);
        if (prior === E.STATUS_COMPLETED || prior === E.STATUS_DISCARDED ||
            prior === E.STATUS_SUPERSEDED) {
          ss.toast(`This request is already ${prior.toLowerCase()} — removing duplicate row.`,
                   '⚠️ Duplicate Row', 5);
          return;
        }

        Ledger.appendEvent({
          event_id         : Utils.newEventId(),
          request_id       : reqId,
          event_type       : E.EVENT_DISCARDED,
          event_ts         : now,
          actor_email      : pickerEmail,
          actor_display    : pickerDisp,
          stock_number     : stockNum,
          build_load_level : buildLevel,
          lot_status       : lotStatus,
          raw_input        : rawBin,
          notes            : notes,
        });
        Ledger.upsertRequest({
          request_id     : reqId,
          complete_ts    : now,
          picker_email   : pickerEmail,
          picker_display : pickerDisp,
          status         : E.STATUS_DISCARDED,
          notes          : notes,
        });
      });
    } catch (err) {
      console.error('_handleDiscard ledger write failed:', err.message);
    }
  }

  sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).clearContent();
  SpreadsheetApp.flush();

  const toastMsg = notes
    ? `Row discarded (N/A). Note: "${notes}"`
    : 'Row discarded (N/A). Logged as DISCARDED.';
  ss.toast(toastMsg, '🗑️ Discarded', 5);
}

// --------------------------------------------------------------------------
// _parseBins(rawBin)
// Parses a picker's bin entry into structured bin objects.
// Priority has been removed — URG is a physical rack name, not an urgency flag.
// --------------------------------------------------------------------------
function _parseBins(rawBin) {
  const BIN = CONFIG.BIN;

  const tokens = rawBin.split(BIN.DELIMITER_PATTERN).map(t => t.trim().toUpperCase()).filter(t => t !== '');
  const bins   = [];

  for (const rawToken of tokens) {
    if (BIN.GARBAGE_PATTERN.test(rawToken)) {
      return { rejected: true, rejectedToken: rawToken };
    }

    // Bulk shorthand (B1.2 → BULK1.2): expand BEFORE classification so the
    // stored bin code is the full BULK form everywhere downstream — ledger,
    // reports, archives. rawToken (what was typed, uppercased) still flows
    // to _EventLog.raw_input untouched.
    const code = BIN.BULK_SHORTHAND_PATTERN.test(rawToken)
        ? rawToken.replace(BIN.BULK_SHORTHAND_PATTERN, BIN.BULK_SHORTHAND_REPLACEMENT)
        : rawToken;

    const locationType = _classifyLocationType(code);
    bins.push({ code: code, locationType, rawToken });
  }

  return { bins };
}

// --------------------------------------------------------------------------
// _classifyLocationType(code)
// --------------------------------------------------------------------------
function _classifyLocationType(code) {
  const BIN = CONFIG.BIN;

  for (const entry of BIN.LOCATION_PREFIXES) {
    if (code.startsWith(entry.prefix)) return entry.type;
  }

  if (BIN.RACK_PATTERN.test(code)) return 'RACK';

  return BIN.LOCATION_TYPE_UNKNOWN;
}
