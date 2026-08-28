/**
 * Intake.gs
 * Handles both intake paths: Google Form sync and walk-in manual entry.
 * Both paths resolve identity, assign a request_id, emit a REQUESTED event,
 * upsert _Requests, and place the row on the Processing board.
 *
 * Public API:
 *   Intake.syncFormResponses()          → processes new Form Responses 1 rows
 *   Intake.handleWalkIn(sheet, rowNum)  → called by PickerTrigger on manual stock entry
 */

const Intake = (() => {

  // --------------------------------------------------------------------------
  // syncFormResponses()
  // --------------------------------------------------------------------------
  function syncFormResponses() {
    Utils.withLock(() => {
      const ss     = SpreadsheetApp.getActiveSpreadsheet();
      const respS  = Sheets.getByName(CONFIG.SHEETS.FORM_RESPONSES, ss);
      const procS  = Sheets.getByName(CONFIG.SHEETS.PROCESSING, ss);

      if (!respS) throw new Error('Intake.syncFormResponses: Form Responses 1 sheet not found.');
      if (!procS) throw new Error('Intake.syncFormResponses: Processing sheet not found.');

      const lastRow = respS.getLastRow();
      if (lastRow < 2) {
        Utils.toast('No submissions to sync.', '📥 Sync', 3);
        return;
      }

      const allData    = respS.getRange(2, 1, lastRow - 1, respS.getLastColumn()).getValues();
      const respHeaders = respS.getRange(1, 1, 1, respS.getLastColumn()).getValues()[0].map(Utils.cleanHeader);

      const frTs    = respHeaders.indexOf(Utils.cleanHeader(CONFIG.HEADERS.FR_TIMESTAMP));
      const frEmail = respHeaders.indexOf(Utils.cleanHeader(CONFIG.HEADERS.FR_EMAIL));
      const frStock = respHeaders.indexOf(Utils.cleanHeader(CONFIG.HEADERS.FR_STOCK));
      const frBuild = respHeaders.indexOf(Utils.cleanHeader(CONFIG.HEADERS.FR_BUILD));
      const frLot   = respHeaders.indexOf(Utils.cleanHeader(CONFIG.HEADERS.FR_LOT));

      let frSync = respHeaders.indexOf(Utils.cleanHeader(CONFIG.HEADERS.FR_SYNC_STATUS));
      if (frSync === -1) {
        const newCol = respS.getLastColumn() + 1;
        respS.getRange(1, newCol).setValue(CONFIG.HEADERS.FR_SYNC_STATUS);
        frSync = newCol - 1;
      }

      let placed = 0;

      for (let i = 0; i < allData.length; i++) {
        const row = allData[i];
        const syncStatus = String(row[frSync] || '').trim();

        if (syncStatus === CONFIG.ENUMS.SYNC_PROCESSED) continue;

        const rawStock = frStock !== -1 ? String(row[frStock] || '').trim() : '';
        if (!rawStock) {
          // A response with no stock number can never be synced. Stamp it so it
          // stops counting as unprocessed: left blank it pins
          // Health._checkStuckResponses above zero permanently, which makes
          // heartbeatCheck re-run this function after EVERY form submission and
          // monitor() alert every hour. SKIPPED is re-evaluated on later runs —
          // if someone fills the stock number in, the row syncs normally.
          if (syncStatus !== CONFIG.ENUMS.SYNC_SKIPPED) {
            respS.getRange(i + 2, frSync + 1).setValue(CONFIG.ENUMS.SYNC_SKIPPED);
          }
          continue;
        }

        const rawEmail = frEmail !== -1 ? String(row[frEmail] || '').trim() : '';
        const rawBuild = frBuild !== -1 ? String(row[frBuild] || '').trim() : '';
        const rawLot   = frLot   !== -1 ? String(row[frLot]   || '').trim() : '';
        const rawTs    = frTs    !== -1 ? row[frTs] : new Date();

        const identity        = Identity.resolve(rawEmail);
        const resolvedEmail   = Identity.isUnmapped(identity) ? rawEmail : identity.email;
        const resolvedDisplay = Identity.isUnmapped(identity) ? rawEmail : identity.display;
        const resolvedDept    = Identity.isUnmapped(identity) ? ''       : identity.department;

        const requestTs  = Utils.toIso(rawTs instanceof Date ? rawTs : new Date(rawTs));
        const buildLevel = Utils.normalizeBuildLevel(rawBuild) || rawBuild;
        const lotStatus  = Utils.normalizeLotStatus(rawLot);
        const now        = Utils.isoNow();

        // --- Durable dedup: guard against concurrent execution double-emit ---
        // SYNC_PROCESSED in FR1 is the fast-path skip for already-committed rows,
        // but it is not reliable as the sole dedup gate across concurrent GAS
        // executions. GAS's server-side sheet read cache can serve a stale snapshot
        // to a second execution (e.g. heartbeatCheck retry, or a queued
        // onFormSubmitTrigger) even after a prior execution flushed the flag —
        // both executions pass the syncStatus check above, both mint a new
        // requestId, and both emit REQUESTED for the same FR1 row.
        //
        // _Requests rows are written via appendRow, which is immediately durable
        // and not subject to the same cache lag. A second execution under the lock
        // will always observe appendRow writes from the first. findOpenRequest()
        // probes (stock, requester, OPEN) — same dedup key as handleWalkIn — and
        // adopts the existing request_id rather than minting a new one.
        //
        // The SYNC_PROCESSED flag is still written unconditionally below; it
        // remains the correct fast-path skip and prevents redundant probes on
        // subsequent calls once the cache has caught up.
        const existingOpenId = Ledger.findOpenRequest(rawStock, resolvedEmail);
        if (existingOpenId) {
          // Prior execution — or an earlier duplicate submission of the same
          // (stock, requester) pair — already committed this request.
          //
          // Re-place the board row ONLY if the board is not already carrying one
          // for this request_id. _placeBoardRows is NOT idempotent: it writes
          // into the first all-empty row it finds and knows nothing about which
          // requests are already on the board. Calling it unconditionally here
          // produces a second Processing line for one request_id with no ledger
          // trail on either side of it — this branch deliberately writes no
          // REQUESTED event and no _Requests row, so nothing in the manifest
          // ever records the extra line.
          //
          // The re-place is still correct when the prior row HAS been cleared
          // (picked/discarded while the request stayed OPEN), which is the case
          // the original unconditional call was reaching for.
          if (!_boardHasRequestId(procS, existingOpenId)) {
            _placeBoardRows(procS, [
              { requestTs, resolvedDisplay, rawStock, buildLevel, lotStatus, requestId: existingOpenId },
            ]);
            placed++;
          }
          respS.getRange(i + 2, frSync + 1).setValue(CONFIG.ENUMS.SYNC_PROCESSED);
          continue;
        }

        const requestId = Utils.newRequestId();

        Ledger.appendEvent({
          event_id             : Utils.newEventId(),
          request_id           : requestId,
          event_type           : CONFIG.ENUMS.EVENT_REQUESTED,
          event_ts             : now,
          actor_email          : resolvedEmail,
          actor_display        : resolvedDisplay,
          intake_source        : CONFIG.ENUMS.INTAKE_FORM,
          stock_number         : rawStock,
          build_load_level     : buildLevel,
          lot_status           : lotStatus,
          requester_department : resolvedDept,
          raw_input            : Identity.isUnmapped(identity) ? rawEmail : '',
        });

        Ledger.upsertRequest({
          request_id           : requestId,
          intake_source        : CONFIG.ENUMS.INTAKE_FORM,
          requested_ts         : requestTs,
          requester_email      : resolvedEmail,
          requester_display    : resolvedDisplay,
          requester_department : resolvedDept,
          stock_number         : rawStock,
          build_load_level     : buildLevel,
          lot_status           : lotStatus,
          status               : CONFIG.ENUMS.STATUS_OPEN,
        });

        // Place this row's board entry, THEN flag the source row PROCESSED — in
        // this order so a mid-loop failure can neither (a) re-emit an already-
        // committed row (the flag is written before we advance to the next) nor
        // (b) flag a row whose board entry is missing. _Requests is the source of
        // truth; the board is a derived view. Per-row placement costs N board reads
        // on a bulk backfill, but the common path (onFormSubmit) is one new row.
        _placeBoardRows(procS, [
          { requestTs, resolvedDisplay, rawStock, buildLevel, lotStatus, requestId },
        ]);
        respS.getRange(i + 2, frSync + 1).setValue(CONFIG.ENUMS.SYNC_PROCESSED);
        placed++;
      }

      SpreadsheetApp.flush();
      Utils.toast(`${placed} submission(s) synced.`, '📥 Sync Complete', 4);
    });
  }

  // --------------------------------------------------------------------------
  // handleWalkIn(procSheet, rowNum)
  //
  // Called by PickerTrigger Path 1 on Stock/Requester/Build edits. Pickers fill
  // these one cell at a time, so this fires repeatedly for one logical request;
  // it must be idempotent. The dedup key is (stock_number, requester_email, OPEN)
  // looked up in _Requests — the source of truth — NOT a flag stamped into the
  // board's Request ID cell. An in-cell claim cannot close the race: two onEdit
  // executions for one row do not share a read cache, so both can read the cell
  // empty before either flush propagates, mint two ids, and emit twice (the exact
  // 8-seconds-apart duplicate seen in _EventLog). Only the read-decide-emit
  // critical section is serialized under the lock.
  // --------------------------------------------------------------------------
  function handleWalkIn(procSheet, rowNum) {
    const H = CONFIG.HEADERS;

    const cols = Utils.getColMap(
      procSheet,
      H.PROC_STOCK,
      H.PROC_LOT,
      H.PROC_REQ_ID,
      H.PROC_REQUESTER,
      H.PROC_BUILD,
      H.PROC_REQUESTED,
    );

    const rawStock = String(procSheet.getRange(rowNum, cols[H.PROC_STOCK]).getValue()).trim();
    if (!rawStock) return;

    // Hold REQUESTED until stock, requester, AND build are all present — pickers
    // fill them incrementally and the event must never carry missing fields.
    const rawRequester = String(procSheet.getRange(rowNum, cols[H.PROC_REQUESTER]).getValue()).trim();
    if (!rawRequester) return;

    const rawBuildCheck = String(procSheet.getRange(rowNum, cols[H.PROC_BUILD]).getValue()).trim();
    if (!rawBuildCheck) return;

    // Cheap pre-lock bailout: if THIS execution already sees a Request ID on the
    // row, the request is committed — skip the lock entirely. This catches the
    // common serialized re-fire; the locked re-check below catches the racy one.
    if (String(procSheet.getRange(rowNum, cols[H.PROC_REQ_ID]).getValue()).trim()) return;

    const identity        = Identity.resolve(rawRequester);
    const resolvedEmail   = Identity.isUnmapped(identity) ? rawRequester : identity.email;
    const resolvedDisplay = Identity.isUnmapped(identity) ? rawRequester : identity.display;
    const resolvedDept    = Identity.isUnmapped(identity) ? ''           : identity.department;
    const buildLevel      = Utils.normalizeBuildLevel(rawBuildCheck) || rawBuildCheck;

    // --- Critical section: serialize read-decide-write across edit executions ---
    // The whole decision (does an OPEN request already exist?) and the writes that
    // act on it must be atomic. Locking only the writes — as the prior version did
    // by wrapping the call site — leaves the check-then-act window open, which is
    // where the double-emit happened. Acquire ONE lock, re-read _Requests under it,
    // and bail if a matching OPEN row already exists. _Requests is durable and
    // flushed by appendRow, so the second execution sees the first's row even when
    // its board snapshot is stale.
    Utils.withLock(() => {
      // Re-read the id cell under the lock — the winning execution may have stamped
      // it after our pre-lock read above.
      const idCell = procSheet.getRange(rowNum, cols[H.PROC_REQ_ID]);
      if (String(idCell.getValue()).trim()) return;

      // Durable dedup: a prior execution for this same logical request will have
      // left an OPEN _Requests row keyed by (stock, requester). If found, adopt it
      // onto the board row (self-heals a row whose id cell was lost) and bail.
      const existingOpenId = Ledger.findOpenRequest(rawStock, resolvedEmail);
      if (existingOpenId) {
        idCell.setValue(existingOpenId);
        SpreadsheetApp.flush();
        return;
      }

      const requestId = Utils.newRequestId();
      const now       = Utils.isoNow();
      // Read lot directly from the cell — no autofill fallback. Picker supplies
      // it or it records as blank. normalizeLotStatus passes unknown values through
      // as-is, so blank is valid and downstream analytics handle it gracefully.
      const lotStatus = Utils.normalizeLotStatus(
        String(procSheet.getRange(rowNum, cols[H.PROC_LOT]).getValue()).trim()
      );

      // Order: append the durable event/request rows FIRST, THEN stamp the board.
      // If we die mid-way, the worst case is a committed _Requests row whose board
      // id cell is blank — which the findOpenRequest probe above repairs on the
      // next fire, rather than an orphaned board claim with no backing request.
      Ledger.appendEvent({
        event_id             : Utils.newEventId(),
        request_id           : requestId,
        event_type           : CONFIG.ENUMS.EVENT_REQUESTED,
        event_ts             : now,
        actor_email          : resolvedEmail,
        actor_display        : resolvedDisplay,
        intake_source        : CONFIG.ENUMS.INTAKE_WALKIN,
        stock_number         : rawStock,
        build_load_level     : buildLevel,
        lot_status           : lotStatus,
        requester_department : resolvedDept,
        raw_input            : Identity.isUnmapped(identity) ? rawRequester : '',
      });

      Ledger.upsertRequest({
        request_id           : requestId,
        intake_source        : CONFIG.ENUMS.INTAKE_WALKIN,
        requested_ts         : now,
        requester_email      : resolvedEmail,
        requester_display    : resolvedDisplay,
        requester_department : resolvedDept,
        stock_number         : rawStock,
        build_load_level     : buildLevel,
        lot_status           : lotStatus,
        status               : CONFIG.ENUMS.STATUS_OPEN,
      });

      idCell.setValue(requestId);
      procSheet.getRange(rowNum, cols[H.PROC_REQUESTED]).setValue(now);
      SpreadsheetApp.flush();
    });
  }

  // --------------------------------------------------------------------------
  // Internal: is this request_id already sitting on the Processing board?
  //
  // Reads the Request ID column by VALUE (not TextFinder, which matches on
  // displayed text) across the full grid budget, so a row placed anywhere —
  // including a reclaimed slot above the newest entry — is seen. The board is
  // bounded by PROCESSING_MIN_ROWS, so this is one cheap single-column read.
  //
  // This is the dedup gate the board never had: _Requests answers "does this
  // request exist?", but only the board can answer "is it already displayed?".
  // --------------------------------------------------------------------------
  function _boardHasRequestId(procSheet, requestId) {
    const id = String(requestId || '').trim();
    if (!id) return false;

    const idCol   = Utils.getColIndex(procSheet, CONFIG.HEADERS.PROC_REQ_ID);
    const maxRows = procSheet.getMaxRows();
    if (maxRows < 2) return false;

    const ids = procSheet.getRange(2, idCol, maxRows - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === id) return true;
    }
    return false;
  }

  // --------------------------------------------------------------------------
  // Internal: place new request rows on the Processing board
  // --------------------------------------------------------------------------
  function _placeBoardRows(procSheet, rows) {
    const H = CONFIG.HEADERS;
    const cols = Utils.getColMap(
      procSheet,
      H.PROC_REQUESTED,
      H.PROC_REQUESTER,
      H.PROC_STOCK,
      H.PROC_BUILD,
      H.PROC_LOT,
      H.PROC_BIN,
      H.PROC_REQ_ID,
    );

    const maxRows   = procSheet.getMaxRows();
    const sheetData = maxRows > 1
      ? procSheet.getRange(2, 1, maxRows - 1, procSheet.getLastColumn()).getValues()
      : [];

    const numCols = procSheet.getLastColumn();
    let itemIdx   = 0;

    for (let i = 0; i < sheetData.length && itemIdx < rows.length; i++) {
      if (sheetData[i].every(cell => cell === '' || cell === null)) {
        _writeBoardRow(procSheet, i + 2, cols, rows[itemIdx]);
        itemIdx++;
      }
    }

    for (; itemIdx < rows.length; itemIdx++) {
      const targetRowNum = procSheet.getLastRow() + 1;
      procSheet.getRange(targetRowNum, 1, 1, numCols).setValues([new Array(numCols).fill('')]);
      _writeBoardRow(procSheet, targetRowNum, cols, rows[itemIdx]);
    }
  }

  // --------------------------------------------------------------------------
  // Internal: write a single request's fields to a specific board row
  // --------------------------------------------------------------------------
  function _writeBoardRow(sheet, rowNum, cols, r) {
    const H = CONFIG.HEADERS;
    sheet.getRange(rowNum, cols[H.PROC_REQUESTED]).setValue(r.requestTs);
    sheet.getRange(rowNum, cols[H.PROC_REQUESTER]).setValue(r.resolvedDisplay);
    sheet.getRange(rowNum, cols[H.PROC_STOCK]).setValue(r.rawStock);
    sheet.getRange(rowNum, cols[H.PROC_BUILD]).setValue(r.buildLevel);
    sheet.getRange(rowNum, cols[H.PROC_LOT]).setValue(r.lotStatus);
    sheet.getRange(rowNum, cols[H.PROC_REQ_ID]).setValue(r.requestId);
  }

  // --------------------------------------------------------------------------
  // Public surface
  // --------------------------------------------------------------------------
  return {
    syncFormResponses,
    handleWalkIn,
  };

})();

// Top-level stub for menu item
function Intake_syncFormResponses() { Intake.syncFormResponses(); }
