/**
 * Turnover.gs   (rev 3.3)
 * Shift turnover: report + SHIFT-SCOPED WORKBOOK SNAPSHOT + full clean slate at
 * BOTH shift boundaries.
 *
 * REV 3.3 — THE ARCHIVE IS SHIFT-AWARE NOW. Everything that read the ledger was
 * already shift-scoped (ShiftReport.computeShiftStats filters on op day AND
 * shift; _requestedInOrBefore exists specifically because the trigger fires
 * after the boundary). The ARCHIVE never got that treatment: _archiveWorkbookCopy
 * was a bare DriveApp.makeCopy() with no filter of any kind, and the guard
 * deciding whether to archive at all (_hasAnyData) was a raw row-count scan.
 * The two halves therefore disagreed. Observed on 2026-08-26: the DAY report was
 * correctly suppressed as "no activity" while the DAY archive was created
 * holding 17 requests, every one stamped 17:16–17:52 — all nightshift, because
 * dayshift had not adopted the sheet yet and the trigger lands at 17:54.
 *
 * The trigger is NOT jittery, which is what made this a daily event rather than
 * an edge case: 30 of 32 DAY runs fired at exactly 17:54 and 28 of 34 NIGHT runs
 * at exactly 06:13. Google picked a minute offset and has held it. That is 54
 * minutes of the incoming shift's work swept into every DAY snapshot.
 *
 *   THE FIX — THE BOUNDARY IS A COMPUTABLE INSTANT, NOT THE TRIGGER'S FIRE TIME.
 *   _boundaryInstant() derives it from the op-day key: DAY ends at
 *   opDay@SECOND_HR, NIGHT ends at opDay+1@START_HR. Every inclusion decision
 *   compares against that, so all of this is immune to when the trigger lands.
 *
 *   MEMBERSHIP RULE: a request belongs to the archive of the shift it was
 *   REQUESTED in; its STATE is as-of the snapshot. A day request at 16:50 that a
 *   night picker completes at 17:20 stays in the DAY archive showing COMPLETED.
 *   _EventLog, Processing and Form Responses follow their REQUEST, not their own
 *   timestamps — filtering the log on event_ts would leave that request in the
 *   DAY archive with its PICKED/COMPLETED rows orphaned.
 *
 *   Consequence: archives now partition instead of overlapping wholesale. The
 *   remaining overlap is genuine carryover — an OPEN row predating two
 *   consecutive boundaries appears in both, which is exactly where you want it
 *   twice. A day request completed after 17:00 also appears in both, and that
 *   matches how the reports have always counted it: a DAY request and a NIGHT
 *   completion.
 *
 *   Three changes, all keyed off the boundary instant:
 *     1. _hasAnyData → _hasDataBefore. The nothing-to-archive guard now asks
 *        "is there anything from the ended shift?" instead of "are there rows?".
 *        Post-boundary rows no longer defeat it.
 *     2. _cleanSlate step 4 gains survivor rules. Form Responses 1 was the only
 *        one of the four data sheets wiped unconditionally — see below.
 *     3. _pruneArchiveCopy(). makeCopy() cannot filter, so ARCHIVE-FIRST is
 *        preserved by copying everything and pruning the COPY afterward.
 *        Best-effort: a failure degrades to rev-3.2 behavior (over-inclusive
 *        archive, nothing lost) and emails ADMIN_EMAIL. Kill-switchable via
 *        CONFIG.TURNOVER_PRUNE_ARCHIVE without a redeploy.
 *
 *   FORM RESPONSES DATA LOSS (fixed in 2, and the only genuine loss path here):
 *   rev 3.2 deleted every FR1 data row with no survivor rule while _Requests,
 *   _EventLog and Processing all had one. A response that landed at 17:53 and
 *   had not synced yet was destroyed — Intake and Health._checkStuckResponses
 *   both read FR1, and nothing ever reads an archive back. Two rules now: keep
 *   rows stamped at/after the boundary (the incoming shift's), and keep rows
 *   that are neither PROCESSED nor SKIPPED_NO_STOCK (never reached _Requests).
 *   NOTE the trade: an unsyncable row now PERSISTS across turnovers instead of
 *   being swept silently, so it will keep Health alerting until a human clears
 *   it. That is the intended direction — loud beats lossy — but watch for it.
 *
 *   TRIGGER SCHEDULE DELIBERATELY UNCHANGED. .nearMinute() is ±15 min, and
 *   anything that can fire BEFORE the boundary is worse than firing late:
 *   _endedShift() would resolve to the PREVIOUS boundary, _alreadyRan() would
 *   match that morning's stamp, and the run would return silently — no archive,
 *   no wipe, console log only, with the next boundary quietly absorbing two
 *   shifts. The current 17:54 slot has 54 minutes of margin and has never fired
 *   early. Keep it. The durable fix is self-identifying handlers
 *   (Turnover_runDayBoundary / Turnover_runNightBoundary passing an expected
 *   shift) so run() validates against intent rather than inferring from the
 *   clock; that touches the Setup trigger inventory and is left for later.
 *
 * REV 3.2 — DUPLICATED-FORM CLEANUP ACTUALLY WORKS NOW. The rev-3.1 cleanup
 * threw on every run because the manifest's explicit oauthScopes lacked
 * https://www.googleapis.com/auth/forms, so FormApp.openByUrl() was denied,
 * the best-effort catch swallowed it into a trigger-log warning nobody saw,
 * and one stray live-linked "Copy of <form>" accreted per boundary. Fixes:
 *   • appsscript.json now declares the forms scope. RE-AUTH REQUIRED: the
 *     trigger owner must run any function once from the editor to grant it,
 *     or trigger runs fail on auth instead of warning.
 *   • Form-link discovery hardened: sheet-level getFormUrl() (more reliable
 *     on a fresh copy than the spreadsheet-level call, which can return null
 *     before the duplicated link propagates) with a short retry loop.
 *   • Cleanup failure is no longer silent: a live-linked form pointed at an
 *     archive violates the immutability invariant, so failure now emails
 *     ADMIN_EMAIL (still never aborts the turnover — the archive stands).
 *
 * Boundaries (constants in CONFIG.SHIFT):
 *   ~17:00 — DAY shift just ended  → report + snapshot + wipe day's data.
 *   ~06:00 — NIGHT shift just ended → report + snapshot + wipe night's data.
 *
 * REV 3 — SNAPSHOT ARCHIVE + CLEAN SLATE. Storage is unlimited, so the rev-2
 * per-shift CSV partitioning (grouping by complete_ts, UNBUCKETED files,
 * contiguous-block deletion) is retired. Instead:
 *
 *   ARCHIVE = one full Drive copy of the workbook, made BEFORE anything is
 *   wiped, then pruned to the ended shift (rev 3.3). Full fidelity: every tab
 *   (_EventLog included), formatting, notes.
 *   Named Turnover_{opDay}_{SHIFT}_{HHmm} and filed in the dated backup tree.
 *   Copies are inert — installable triggers do not survive makeCopy(), and the
 *   duplicated Form that makeCopy() creates for a form-linked workbook is
 *   unlinked and trashed immediately (rev 3.1, repaired in 3.2), so an
 *   archived workbook never syncs, emails, purges, or RECEIVES anything.
 *
 *   EMPTY BOUNDARIES (rev 3.1, corrected in 3.3): an automated run whose ended
 *   shift produced zero activity, zero open carryover, and no data sheet rows
 *   PREDATING THE BOUNDARY makes NO snapshot and wipes nothing — it only stamps
 *   the once-per guard. A missing file in the backup tree therefore means
 *   "empty shift" OR "failed run"; the trigger-failure email from Apps Script
 *   disambiguates. Manual runs still archive unconditionally.
 *
 *   CLEAN SLATE = every data sheet is reset so the incoming shift's data
 *   views start empty: Form Responses 1, Processing, _Requests, _EventLog.
 *   _Identity and _Roles are configuration, not data — never touched.
 *
 * WHAT SURVIVES THE WIPE (and why — do not "simplify" these away):
 *   • OPEN _Requests rows (+ their _EventLog rows + their board rows).
 *     Dept rule says no open work crosses a boundary, so these should be zero;
 *     when they aren't, destroying live requests is not an acceptable failure
 *     mode. Carryover is counted, flagged ⚠ in the completion message, and
 *     still disables skip-empty on the report (openTotal > 0).
 *   • Terminal rows whose complete_ts lands in the CURRENT (live) op day and
 *     shift — the LIVE-SHIFT RULE, unchanged from rev 2. Trigger jitter means
 *     the 17:00 run really fires at 17:5x; picks completed in those minutes
 *     belong to the shift now on, and ShiftReport.computeShiftStats() reads
 *     them from _Requests at the NEXT boundary. Wipe them and the next
 *     report undercounts. Evaluated at execution time; jitter/DST-proof.
 *   • In-flight board rows (content but no Request ID yet) — a walk-in the
 *     picker is mid-typing. Never yank a row out from under someone.
 *   • Form Responses rows stamped at/after the boundary, and any response that
 *     has not reached _Requests yet (rev 3.3 — see the header note above).
 *
 * _EventLog SEMANTICS: the log is no longer immutable in the LIVE workbook —
 * immutability lives in the archive chain. Each snapshot contains the ended
 * shift's log; rows carried over (OPEN / live-shift) appear again in the next
 * snapshot, so the union of snapshots is complete and gap-free.
 *
 * ARCHIVE-FIRST invariant (unchanged): the snapshot is created and verified
 * before a single row is touched. If the copy throws, nothing is wiped.
 *
 * Missed-boundary healing is intact: leftover rows from an earlier unarchived
 * shift are by definition pre-boundary, so they still force a full snapshot and
 * survive pruning; the wipe then clears them.
 *
 * What one run does, in order:
 *   0. RECONCILE — Reconcile.reconcile() flips re-fulfilled discards to
 *      SUPERSEDED so the discard rate below is accurate (see Reconcile.gs).
 *   1. SNAPSHOT STATS — ShiftReport.computeShiftStats() for the ended shift.
 *      MUST precede the wipe: it reads the rows the wipe removes.
 *   2. ARCHIVE — full workbook copy into the dated backup folder.
 *   3. CLEAN SLATE — wipe the four data sheets, honoring the survivor rules.
 *   4. Re-apply formatting; stamp the once-per-(op day, shift) guard.
 *      [steps 0–4 inside one lock]
 *   5. PRUNE — trim the archive copy to the ended shift, OUTSIDE the lock. The
 *      copy is already committed and verified and the live workbook is done
 *      being modified, so holding the lock through four more sheet rewrites
 *      during shift change — the busiest minutes of the day for onEdit and
 *      onFormSubmit — buys nothing.
 *   6. EMAIL — ended shift's report from the step-1 snapshot, OUTSIDE the
 *      lock. Skip-empty (CONFIG.TURNOVER_SKIP_EMPTY_REPORT) unchanged:
 *      automated + zero activity + zero open carryover → suppressed;
 *      manual runs always send. Send failure alerts ADMIN_EMAIL; the
 *      turnover itself never throws over mail.
 *
 * Automated vs manual:
 *   - Automated: guarded once per (op day, shift); skip-empty may suppress mail.
 *   - Manual (menu): confirms via dialog, always proceeds, always mails.
 *     A manual re-run after a boundary already wiped produces a mostly-empty
 *     snapshot and a zeros report — harmless; the boundary run's email is
 *     the report of record.
 *
 * Public API:
 *   Turnover.run()   → interactive (confirms via dialog); safe under time trigger too
 */

const Turnover = (() => {

  // -------------------------------------------------------------------------------
  // run()
  // -------------------------------------------------------------------------------
  function run() {
    // SpreadsheetApp.getUi() throws under a time-driven trigger — its absence is
    // exactly how we distinguish automated from manual.
    let ui = null;
    try { ui = SpreadsheetApp.getUi(); } catch (_) { /* time-driven — no UI */ }

    const isAutomated = !ui;
    const now         = new Date();
    const ended       = _endedShift(now);        // { shift, target:{ key, date } }
    const boundary    = _boundaryInstant(ended); // the instant that shift finished

    if (ui) {
      const resp = ui.alert(
        '🗄️ CONFIRM SHIFT TURNOVER',
        `Runs turnover for the ${ended.shift} shift of op day ${ended.target.key} ` +
        '(the shift that most recently ended):\n\n' +
        '• Emails that shift\'s report to all managers\n' +
        '• Archives a COPY of this workbook, trimmed to that shift, to the backup folder\n' +
        '• Wipes Form Responses, Processing, _Requests, and _EventLog — a clean\n' +
        '  slate for the shift now on\n\n' +
        'OPEN requests, the CURRENT shift\'s completions, unsynced form responses,\n' +
        'and in-progress board entries are always carried over. The archive copy\n' +
        'preserves everything else, including the full event history.\n\n' +
        'Note: if this boundary already ran, the re-sent report will show zeros — ' +
        'the boundary email is the report of record.\n\nProceed?',
        ui.ButtonSet.YES_NO
      );
      if (resp !== ui.Button.YES) return;
    }

    // Once-per-(op day, shift) guard — automated only. Manual runs always
    // proceed (deliberate re-runs; the survivor rules keep them safe).
    if (isAutomated && _alreadyRan(ended)) {
      console.log(`Turnover.run: skipping — already ran for ${ended.target.key}|${ended.shift}.`);
      return;
    }

    let stats, archive, wiped, skipped = false;

    Utils.withLock(() => {
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      // 0. Reconcile re-fulfilled discards BEFORE the stats snapshot, so the
      //    report's discard rate is accurate and the archive copy carries the
      //    corrected SUPERSEDED statuses. Lock-free by contract — we hold it.
      Reconcile.reconcile();

      // 1. Stats snapshot FIRST — the wipe below removes these rows.
      stats = ShiftReport.computeShiftStats(ended.target, ended.shift);

      // 1.5 NOTHING-TO-ARCHIVE GUARD (automated runs only). If the ended shift
      //     had zero activity, zero open carryover, AND no data row predates
      //     the boundary, the archive would snapshot the incoming shift's work
      //     under the ended shift's name — which is the rev-3.3 bug. Skip both;
      //     stamp the boundary as done. This mirrors TURNOVER_SKIP_EMPTY_REPORT,
      //     applied to the artifacts instead of just the email, and it now uses
      //     the SAME shift-scoped notion of "empty" that the report side uses.
      //     Missed-boundary healing is intact: leftover rows from an earlier
      //     unarchived shift are pre-boundary and still force a full snapshot.
      //     Manual runs still proceed unconditionally (deliberate re-runs).
      if (isAutomated &&
          stats.requests === 0 && stats.completed === 0 && stats.discarded === 0 &&
          stats.openTotal === 0 && !_hasDataBefore(ss, boundary)) {
        skipped = true;
        _stampRun(ended);
        return;
      }

      // 2. ARCHIVE-FIRST: full workbook copy. Throws on failure → nothing wiped.
      archive = _archiveWorkbookCopy(ss, ended);

      // 3. Clean slate, honoring the survivor rules.
      wiped = _cleanSlate(ss, boundary);

      // 4. Housekeeping.
      Formatting.applyFormatting();
      SpreadsheetApp.flush();
      _stampRun(ended);
    });

    if (skipped) {
      console.log(`Turnover.run: ${ended.target.key}|${ended.shift} — nothing from that shift to ` +
                  'archive; boundary stamped, no snapshot made, nothing wiped, report suppressed.');
      return;
    }

    // --- 5. Prune the archive to the ended shift — OUTSIDE the lock ---
    // The copy is committed and verified; the live workbook is untouched by
    // this. Best-effort by design: a failure leaves the rev-3.2 archive (whole
    // workbook, over-inclusive but complete), so nothing here may throw.
    let pruned = null;
    if (archive) {
      if (CONFIG.TURNOVER_PRUNE_ARCHIVE === true) {
        try {
          pruned = _pruneArchiveCopy(archive.id, boundary);
          console.log(`Turnover: pruned "${archive.name}" to the ${ended.shift} shift — removed ` +
                      `${pruned.requests} request row(s), ${pruned.events} event row(s), ` +
                      `${pruned.responses} form response(s), ${pruned.board} board row(s).`);
        } catch (err) {
          console.error('Turnover.run: archive prune failed:', err.message);
          _alertAdminPruneFailure(ended, archive.name, err);
        }
      } else {
        console.warn('Turnover: CONFIG.TURNOVER_PRUNE_ARCHIVE is off — archive left ' +
                     'over-inclusive (rev-3.2 behavior).');
      }
    }

    // --- 6. Shift report — OUTSIDE the lock, from the pre-wipe snapshot ---
    const isEmpty  = stats.requests === 0 && stats.completed === 0 && stats.discarded === 0;
    const skipSend = isAutomated && CONFIG.TURNOVER_SKIP_EMPTY_REPORT === true &&
                     isEmpty && stats.openTotal === 0;

    let mailed = 0;
    if (skipSend) {
      console.log(`Turnover.run: ${ended.target.key}|${ended.shift} had no activity — report suppressed.`);
    } else {
      try {
        mailed = ShiftReport.sendShiftReport(ended.target, ended.shift, ShiftReport.defaultRecipients(),
                                             { partial: false, stats: stats });
      } catch (err) {
        console.error('Turnover.run: shift report failed:', err.message);
        _alertAdminMailFailure(ended, err);
      }
    }

    const msg = `${ended.shift} shift, op day ${ended.target.key}: workbook archived as "${archive.name}". ` +
                `Wiped ${wiped.requestsPurged} request row(s), ${wiped.eventsPurged} event row(s), ` +
                `${wiped.responsesCleared} form response(s), board reset. ` +
                `Carried over: ${wiped.liveKept} live-shift completion(s), ${wiped.boardKept} board row(s), ` +
                `${wiped.responsesKept} form response(s). ` +
                (skipSend ? 'Report suppressed (no activity).'
                          : `Report emailed to ${mailed} recipient(s).`) +
                (stats.openTotal > 0 ? ` ⚠️ ${stats.openTotal} OPEN request(s) crossed the boundary!` : '') +
                (wiped.responsesUnsynced > 0
                   ? ` ⚠️ ${wiped.responsesUnsynced} UNSYNCED form response(s) kept — run ` +
                     'Parts Management → "Sync Form Submissions".' : '');
    if (ui) ui.alert('✅ Turnover Complete', msg, ui.ButtonSet.OK);
    else    Utils.toast(msg, '🗄️ Shift Turnover', 5);
  }

  // -------------------------------------------------------------------------------
  // _boundaryInstant(ended) → Date
  //
  // The exact instant the ended shift finished — the anchor for every inclusion
  // decision in the archive, replacing "whenever the trigger happened to fire".
  //
  //   DAY   of op day K → K   @ SECOND_HR:00:00
  //   NIGHT of op day K → K+1 @ START_HR:00:00   (night wraps past midnight)
  //
  // Derived from the op-day KEY (which IS the calendar date the op day started)
  // by component construction, so it is DST-proof the same way _dateFromKey is —
  // and neither 06:00 nor 17:00 can land in a transition gap. Date() normalizes
  // day overflow, so K+1 is safe at month and year ends.
  // -------------------------------------------------------------------------------
  function _boundaryInstant(ended) {
    const p = String(ended.target.key).split('-');
    const y = +p[0], mo = +p[1] - 1, d = +p[2];
    return ended.shift === ShiftReport.SHIFTS.DAY
      ? new Date(y, mo, d,     CONFIG.SHIFT.SECOND_HR, 0, 0)
      : new Date(y, mo, d + 1, CONFIG.SHIFT.START_HR,  0, 0);
  }

  // -------------------------------------------------------------------------------
  // _archiveWorkbookCopy(ss, ended) → { name, id, url }
  //
  // Full-fidelity Drive copy of the entire workbook into the ended shift's
  // dated backup folder. Bound script code copies with the file; installable
  // triggers do NOT — archives are inert. Throws if the copy cannot be made
  // or verified, which aborts the lock block before any wipe (ARCHIVE-FIRST).
  // Trimming to the ended shift happens afterward in _pruneArchiveCopy().
  // -------------------------------------------------------------------------------
  function _archiveWorkbookCopy(ss, ended) {
    const stamp  = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'HHmm');
    const name   = `Turnover_${ended.target.key}_${ended.shift}_${stamp}`;
    const folder = _getOrCreateDatedFolder(ended.target.date);

    const copy = DriveApp.getFileById(ss.getId()).makeCopy(name, folder);
    if (!copy || !copy.getId()) {
      throw new Error('Turnover: workbook archive copy failed — nothing was wiped.');
    }

    // makeCopy() on a FORM-LINKED workbook also duplicates the linked Form
    // ("Copy of <form name>", filed next to the original) and live-links it to
    // the archive copy. That breaks two invariants at once: the workspace
    // accretes one stray form per boundary, and a submission to the stray form
    // would WRITE INTO the archived snapshot — the archive chain is only
    // immutable if the copy is truly inert. Unlink and trash the duplicate.
    //
    // The cleanup never aborts the turnover (the archive is already committed),
    // but as of rev 3.2 a failure is no longer silent: it emails ADMIN_EMAIL,
    // because a live-linked form on an archive is an invariant violation, not
    // a cosmetic blemish. Requires the .../auth/forms scope (see manifest).
    try {
      const copySS = SpreadsheetApp.openById(copy.getId());

      // Sheet-level getFormUrl() is more reliable on a fresh copy than the
      // spreadsheet-level call, which can return null before the duplicated
      // form-link has propagated on the backend. Retry briefly to cover the
      // propagation gap; fall back to the spreadsheet-level call each pass.
      let formUrl = null;
      for (let attempt = 0; attempt < 3 && !formUrl; attempt++) {
        if (attempt > 0) Utilities.sleep(2000);
        formUrl = copySS.getSheets()
                        .map(sh => { try { return sh.getFormUrl(); } catch (_) { return null; } })
                        .find(u => u)
                  || copySS.getFormUrl();
      }

      if (formUrl) {
        const dupForm = FormApp.openByUrl(formUrl);
        const dupName = dupForm.getTitle();
        const dupId   = dupForm.getId();
        dupForm.removeDestination();
        DriveApp.getFileById(dupId).setTrashed(true);
        console.log(`Turnover: unlinked and trashed duplicated form "${dupName}" (${dupId}).`);
      } else {
        // No link visible after retries. Either the source form-link didn't
        // duplicate (fine) or propagation outran three attempts (a stray
        // "Copy of ..." form may appear later). Log it either way.
        console.warn('Turnover: archive copy reports no linked form after retries — ' +
                     'if a stray "Copy of ..." form appears, link propagation outran the check.');
      }
    } catch (err) {
      console.warn('Turnover: could not clean up duplicated form copy:', err.message);
      _alertAdminFormCleanupFailure(ended, name, err);
    }

    console.log(`Turnover: archived workbook copy "${name}" (${copy.getId()}).`);
    return { name: name, id: copy.getId(), url: copy.getUrl() };
  }

  // -------------------------------------------------------------------------------
  // _pruneArchiveCopy(copyId, boundary) → { requests, events, responses, board }
  //
  // Trims the committed archive copy to the ended shift. makeCopy() cannot
  // filter, so ARCHIVE-FIRST is preserved by copying everything first and
  // removing the incoming shift's rows here.
  //
  // _Requests is the SPINE and is judged on requested_ts. Everything else
  // follows request membership, NOT its own timestamp — a day request picked at
  // 17:20 must keep its events in the DAY archive or the row is orphaned. Rows
  // with no request_id (in-flight walk-ins, stray log rows) fall back to their
  // own timestamp. Unparseable timestamps are KEPT (conservative — an archive
  // that holds too much is recoverable, one that holds too little is not).
  //
  // deleteRows rather than clearContent + rewrite so notes, banding and
  // conditional formatting shift with their rows instead of being stranded
  // against the wrong data — full fidelity is the point of a snapshot archive.
  // Safe because the copy is inert by now: triggers do not survive makeCopy()
  // and the duplicated form was unlinked and trashed above.
  // -------------------------------------------------------------------------------
  function _pruneArchiveCopy(copyId, boundary) {
    const H    = CONFIG.HEADERS;
    const live = SpreadsheetApp.getActiveSpreadsheet();
    const arch = SpreadsheetApp.openById(copyId);
    const out  = { requests: 0, events: 0, responses: 0, board: 0 };

    const gidOf = (name) => {
      try { const sh = Sheets.getByName(name, live); return sh ? sh.getSheetId() : null; }
      catch (_) { return null; }
    };

    // ---- 1. _Requests: the spine. Membership by INTAKE time. ----
    const keptIds = new Set();
    let   spineOk = false;

    const reqSheet = _archSheet(arch, gidOf(CONFIG.SHEETS.REQUESTS), CONFIG.SHEETS.REQUESTS);
    if (reqSheet && reqSheet.getLastRow() > 1) {
      const data  = reqSheet.getDataRange().getValues();
      const clean = data[0].map(Utils.cleanHeader);
      const tsIdx = clean.indexOf(Utils.cleanHeader(H.RQ_REQUESTED_TS));
      const idIdx = clean.indexOf(Utils.cleanHeader(H.RQ_REQUEST_ID));

      if (tsIdx !== -1) {
        const drop = [];
        for (let i = 1; i < data.length; i++) {
          const id = idIdx !== -1 ? String(data[i][idIdx]).trim() : '';
          const ts = _parseTs(data[i][tsIdx]);
          if (ts && ts >= boundary) { drop.push(i + 1); continue; }
          if (id) keptIds.add(id);                    // unparseable ts → kept
        }
        out.requests = _deleteRowsBatched(reqSheet, drop);
        spineOk = true;
      } else {
        console.warn('Turnover._pruneArchiveCopy: _Requests has no requested_ts column — ' +
                     'falling back to per-sheet timestamps.');
      }
    }

    // ---- 2. _EventLog: follows the spine, NOT event_ts. ----
    const logSheet = _archSheet(arch, gidOf(CONFIG.SHEETS.EVENT_LOG), CONFIG.SHEETS.EVENT_LOG);
    if (logSheet && logSheet.getLastRow() > 1) {
      const data  = logSheet.getDataRange().getValues();
      const clean = data[0].map(Utils.cleanHeader);
      const idIdx = clean.indexOf(Utils.cleanHeader(H.EL_REQUEST_ID));
      const tsIdx = clean.indexOf(Utils.cleanHeader(H.EL_EVENT_TS));

      const drop = [];
      for (let i = 1; i < data.length; i++) {
        const id = idIdx !== -1 ? String(data[i][idIdx]).trim() : '';
        if (spineOk && id) {
          if (!keptIds.has(id)) drop.push(i + 1);
          continue;
        }
        const ts = tsIdx !== -1 ? _parseTs(data[i][tsIdx]) : null;
        if (ts && ts >= boundary) drop.push(i + 1);
      }
      out.events = _deleteRowsBatched(logSheet, drop);
    }

    // ---- 3. Form Responses 1: its own timestamp (no request_id column). ----
    const respSheet = _archSheet(arch, gidOf(CONFIG.SHEETS.FORM_RESPONSES),
                                 CONFIG.SHEETS.FORM_RESPONSES);
    if (respSheet && respSheet.getLastRow() > 1) {
      const data  = respSheet.getRange(1, 1, respSheet.getLastRow(),
                                       respSheet.getLastColumn()).getValues();
      const tsIdx = data[0].map(Utils.cleanHeader).indexOf(Utils.cleanHeader(H.FR_TIMESTAMP));

      const drop = [];
      if (tsIdx !== -1) {
        for (let i = 1; i < data.length; i++) {
          const ts = _parseTs(data[i][tsIdx]);
          if (ts && ts >= boundary) drop.push(i + 1);
        }
      }
      out.responses = _deleteRowsBatched(respSheet, drop);
    }

    // ---- 4. Processing board: follows the spine; blank budget rows ignored. ----
    const procSheet = _archSheet(arch, gidOf(CONFIG.SHEETS.PROCESSING), CONFIG.SHEETS.PROCESSING);
    if (procSheet && procSheet.getLastRow() > 1) {
      const data  = procSheet.getRange(1, 1, procSheet.getLastRow(),
                                       procSheet.getLastColumn()).getValues();
      const clean = data[0].map(Utils.cleanHeader);
      const idIdx = clean.indexOf(Utils.cleanHeader(H.PROC_REQ_ID));
      const tsIdx = clean.indexOf(Utils.cleanHeader(H.PROC_REQUESTED));

      const drop = [];
      for (let i = 1; i < data.length; i++) {
        if (!data[i].some(v => String(v).trim() !== '')) continue;   // blank budget row
        const id = idIdx !== -1 ? String(data[i][idIdx]).trim() : '';
        if (spineOk && id) {
          if (!keptIds.has(id)) drop.push(i + 1);
          continue;
        }
        const ts = tsIdx !== -1 ? _parseTs(data[i][tsIdx]) : null;
        if (ts && ts >= boundary) drop.push(i + 1);
      }
      out.board = _deleteRowsBatched(procSheet, drop);
    }

    SpreadsheetApp.flush();
    return out;
  }

  // -------------------------------------------------------------------------------
  // _archSheet(arch, gid, name) → Sheet | null
  //
  // Resolve a managed sheet inside an ARCHIVE COPY. Sheets.getByName() must
  // NEVER be pointed at a copy: on a gid miss it adopts the sheet by name and
  // writes that sheet's id back into the shared Script Properties registry —
  // i.e. it would overwrite the LIVE workbook's gid registry with the archive's
  // ids, breaking rename-proof resolution until the next self-heal. It would
  // also rename tabs inside the archive as a side effect.
  //
  // Copies preserve sheet ids, so the gid pass is near-always a hit; the name
  // lookup is the fallback. Read-only with respect to the registry either way.
  // -------------------------------------------------------------------------------
  function _archSheet(arch, gid, name) {
    if (gid !== null && gid !== undefined) {
      const all = arch.getSheets();
      for (let i = 0; i < all.length; i++) {
        if (all[i].getSheetId() === gid) return all[i];
      }
    }
    return arch.getSheetByName(name);
  }

  // -------------------------------------------------------------------------------
  // _deleteRowsBatched(sheet, rowNumbers) → count deleted
  //
  // Deletes 1-based sheet rows, collapsing them into contiguous runs and working
  // bottom-up so earlier row numbers stay valid as the sheet shrinks. A full
  // shift costs a handful of calls instead of one per row. Input need not be
  // sorted or deduplicated in order; it is sorted here.
  // -------------------------------------------------------------------------------
  function _deleteRowsBatched(sheet, rowNumbers) {
    if (!rowNumbers || rowNumbers.length === 0) return 0;
    const rows = rowNumbers.slice().sort((a, b) => a - b);

    let removed = 0;
    for (let end = rows.length - 1; end >= 0; ) {
      let start = end;
      while (start > 0 && rows[start - 1] === rows[start] - 1) start--;
      const count = end - start + 1;
      sheet.deleteRows(rows[start], count);
      removed += count;
      end = start - 1;
    }
    return removed;
  }

  // -------------------------------------------------------------------------------
  // _cleanSlate(ss, boundary) → counts
  //
  // Resets all four data sheets. Survivor rules (see file header):
  //   _Requests   : keep OPEN/blank-status rows; keep terminal rows whose
  //                 complete_ts is in the CURRENT (live) op day + shift.
  //                 Terminal rows with an unparsable complete_ts are wiped
  //                 (they live on in the archive copy) and console.warn'd.
  //   _EventLog   : keep rows whose request_id belongs to a kept request.
  //   Processing  : keep rows with content whose Request ID is blank
  //                 (in-flight walk-in) or belongs to an OPEN request.
  //                 Live-shift TERMINAL requests keep their ledger rows but
  //                 NOT their board rows — the board shows open work only.
  //   Form Resp.  : keep rows stamped at/after the boundary, and rows that have
  //                 not reached _Requests yet. Everything else deleted (rev 3.3
  //                 — this sheet previously had NO survivor rule at all).
  //
  // Managed sheets are cleared with clearContent + rewrite (grid rows and the
  // Processing row budget stay intact; Formatting.applyFormatting() re-covers
  // validation afterward). The live shift is evaluated at execution time so
  // the rule is correct even if the lock was held across a boundary.
  // -------------------------------------------------------------------------------
  function _cleanSlate(ss, boundary) {
    const H = CONFIG.HEADERS;
    const counts = { requestsPurged: 0, liveKept: 0, openCarried: 0,
                     eventsPurged: 0, eventsKept: 0,
                     boardKept: 0, responsesCleared: 0,
                     responsesKept: 0, responsesUnsynced: 0 };

    const nowEval   = new Date();
    const liveKey   = ShiftReport.opDayKey(nowEval);
    const liveShift = ShiftReport.shiftOf(nowEval);

    // SUPERSEDED is terminal: without it here, superseded rows would read as
    // "open", survive every wipe, and re-match forever.
    const TERMINAL = new Set([CONFIG.ENUMS.STATUS_COMPLETED,
                              CONFIG.ENUMS.STATUS_DISCARDED,
                              CONFIG.ENUMS.STATUS_SUPERSEDED]);

    // ---- 1. _Requests: classify, wipe, rewrite survivors ----
    const openIds = new Set();   // OPEN request ids → board rows also survive
    const keptIds = new Set();   // OPEN + live-shift terminal → event rows survive

    const reqSheet = Sheets.getByName(CONFIG.SHEETS.REQUESTS, ss);
    if (reqSheet && reqSheet.getLastRow() > 1) {
      const data    = reqSheet.getDataRange().getValues();
      const clean   = data[0].map(Utils.cleanHeader);
      const idIdx   = clean.indexOf(Utils.cleanHeader(H.RQ_REQUEST_ID));
      const stIdx   = clean.indexOf(Utils.cleanHeader(H.RQ_STATUS));
      const compIdx = clean.indexOf(Utils.cleanHeader(H.RQ_COMPLETE_TS));

      const survivors = [];
      let unparseable = 0;

      for (let i = 1; i < data.length; i++) {
        const status = stIdx !== -1 ? String(data[i][stIdx]).trim() : '';
        const id     = idIdx !== -1 ? String(data[i][idIdx]).trim() : '';

        if (!TERMINAL.has(status)) {                       // OPEN / blank: carry over
          survivors.push(data[i]);
          if (id) { openIds.add(id); keptIds.add(id); }
          counts.openCarried++;
          continue;
        }
        const d = compIdx !== -1 ? _parseTs(data[i][compIdx]) : null;
        if (d && ShiftReport.opDayKey(d) === liveKey && ShiftReport.shiftOf(d) === liveShift) {
          survivors.push(data[i]);                          // live shift's work: keep
          if (id) keptIds.add(id);
          counts.liveKept++;
          continue;
        }
        if (!d) unparseable++;                              // archived in the copy; wiped here
        counts.requestsPurged++;
      }
      if (unparseable > 0) {
        console.warn(`Turnover._cleanSlate: ${unparseable} terminal row(s) had no parseable complete_ts — wiped (preserved in archive copy).`);
      }
      _rewriteDataRows(reqSheet, survivors);
    }

    // ---- 2. _EventLog: keep only events belonging to kept requests ----
    const logSheet = Sheets.getByName(CONFIG.SHEETS.EVENT_LOG, ss);
    if (logSheet && logSheet.getLastRow() > 1) {
      const data  = logSheet.getDataRange().getValues();
      const clean = data[0].map(Utils.cleanHeader);
      const idIdx = clean.indexOf(Utils.cleanHeader(H.EL_REQUEST_ID));

      const survivors = [];
      for (let i = 1; i < data.length; i++) {
        const id = idIdx !== -1 ? String(data[i][idIdx]).trim() : '';
        if (id && keptIds.has(id)) { survivors.push(data[i]); counts.eventsKept++; }
        else                       { counts.eventsPurged++; }
      }
      _rewriteDataRows(logSheet, survivors);
    }

    // ---- 3. Processing board: open + in-flight rows only ----
    const procSheet = Sheets.getByName(CONFIG.SHEETS.PROCESSING, ss);
    if (procSheet && procSheet.getLastRow() > 1) {
      const lastCol = procSheet.getLastColumn();
      const data    = procSheet.getRange(1, 1, procSheet.getLastRow(), lastCol).getValues();

      let idIdx = -1;
      try { idIdx = Utils.getColIndex(procSheet, H.PROC_REQ_ID) - 1; } catch (_) {}

      const survivors = [];
      for (let i = 1; i < data.length; i++) {
        const hasContent = data[i].some(c => String(c).trim() !== '');
        if (!hasContent) continue;
        const id = idIdx !== -1 ? String(data[i][idIdx]).trim() : '';
        if (id === '' || openIds.has(id)) {                 // in-flight or OPEN: keep
          survivors.push(data[i]);
          counts.boardKept++;
        }
        // terminal / stale / unknown-id rows: dropped (board shows open work)
      }
      procSheet.getRange(2, 1, data.length - 1, lastCol).clearContent();
      if (survivors.length > 0) {
        procSheet.getRange(2, 1, survivors.length, lastCol).setValues(survivors);
      }
    }

    // ---- 4. Form Responses 1: delete only PRE-BOUNDARY, TERMINALLY-SYNCED rows ----
    //      Rev 3.2 deleted every data row unconditionally — the only one of the
    //      four sheets with no survivor rule. Two rules now:
    //
    //        • Timestamp >= boundary — the INCOMING shift's submissions. The
    //          trigger lands ~54 min after the boundary, so this is not an edge
    //          case; it is every single night.
    //
    //        • Not PROCESSED and not SKIPPED_NO_STOCK — the row never reached
    //          _Requests. Deleting it destroys the only copy the live system can
    //          still act on: Intake and Health._checkStuckResponses both read
    //          FR1, and nothing ever reads an archive back. This was a silent,
    //          unrecoverable loss path.
    //
    //      Unparseable Timestamp → kept (conservative, matching the tripwire
    //      philosophy in ShiftReport). deleteRows rather than clearContent so
    //      new submissions keep appending at the top instead of below a block of
    //      blanks; batched into contiguous runs so a full shift costs a handful
    //      of API calls rather than one per row.
    const respSheet = Sheets.getByName(CONFIG.SHEETS.FORM_RESPONSES, ss);
    if (respSheet && respSheet.getLastRow() > 1) {
      const data    = respSheet.getRange(1, 1, respSheet.getLastRow(),
                                         respSheet.getLastColumn()).getValues();
      const clean   = data[0].map(Utils.cleanHeader);
      const tsIdx   = clean.indexOf(Utils.cleanHeader(H.FR_TIMESTAMP));
      const syncIdx = clean.indexOf(Utils.cleanHeader(H.FR_SYNC_STATUS));
      const SYNCED  = new Set([CONFIG.ENUMS.SYNC_PROCESSED, CONFIG.ENUMS.SYNC_SKIPPED]);

      const drop = [];
      for (let i = 1; i < data.length; i++) {
        const ts   = tsIdx   !== -1 ? _parseTs(data[i][tsIdx]) : null;
        const sync = syncIdx !== -1 ? String(data[i][syncIdx]).trim() : '';

        if (!ts || ts >= boundary) { counts.responsesKept++; continue; }
        if (!SYNCED.has(sync))     { counts.responsesKept++; counts.responsesUnsynced++; continue; }
        drop.push(i + 1);
      }
      counts.responsesCleared = _deleteRowsBatched(respSheet, drop);

      if (counts.responsesUnsynced > 0) {
        console.warn(`Turnover._cleanSlate: kept ${counts.responsesUnsynced} unsynced Form Response ` +
                     'row(s) that predate the boundary — these never reached _Requests. They now ' +
                     'persist across turnovers (rev 3.2 deleted them silently). Sync via Parts ' +
                     'Management → "Sync Form Submissions", or they will keep ' +
                     'Health._checkStuckResponses alerting.');
      }
    }

    // Ledger caches sheet handles + headers per execution; after a wipe force
    // any later same-execution reads to come from the live grid.
    Ledger.clearCache();

    return counts;
  }

  // -------------------------------------------------------------------------------
  // _hasDataBefore(ss, boundary) → boolean
  //
  // True if any row PREDATES the boundary — i.e. if there is anything belonging
  // to the ended shift (or an earlier unarchived one) worth snapshotting. Rows
  // stamped at or after the boundary are the incoming shift's and are not this
  // archive's business.
  //
  // Replaces rev 3.2's _hasAnyData(), whose raw row-count scan saw nightshift's
  // post-17:00 rows, returned true, and defeated the nothing-to-archive guard
  // even though computeShiftStats had correctly returned all zeros.
  //
  // Missed-boundary healing is intact: leftover rows from an earlier unarchived
  // shift are by definition pre-boundary, so they still force a full snapshot.
  //
  // Processing is deliberately NOT consulted. Every real board row mirrors a
  // _Requests row that IS checked here, and the one kind that does not — an
  // in-flight walk-in with no Request ID — is a half-typed row with no reliable
  // timestamp, which would force an archive on exactly the empty boundaries this
  // guard exists to skip.
  // -------------------------------------------------------------------------------
  function _hasDataBefore(ss, boundary) {
    const H = CONFIG.HEADERS;
    const checks = [
      [CONFIG.SHEETS.REQUESTS,       H.RQ_REQUESTED_TS],
      [CONFIG.SHEETS.EVENT_LOG,      H.EL_EVENT_TS],
      [CONFIG.SHEETS.FORM_RESPONSES, H.FR_TIMESTAMP],
    ];

    for (let c = 0; c < checks.length; c++) {
      const sh = Sheets.getByName(checks[c][0], ss);
      if (!sh || sh.getLastRow() < 2) continue;

      const data = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
      const idx  = data[0].map(Utils.cleanHeader).indexOf(Utils.cleanHeader(checks[c][1]));
      if (idx === -1) return true;                    // can't judge → snapshot

      for (let i = 1; i < data.length; i++) {
        if (!data[i].some(v => String(v).trim() !== '')) continue;
        const d = _parseTs(data[i][idx]);
        if (!d || d < boundary) return true;          // unparseable → snapshot
      }
    }
    return false;
  }

  // -------------------------------------------------------------------------------
  // _rewriteDataRows(sheet, rows)
  // Clears all data rows (content only — grid stays), then writes `rows`
  // starting at row 2. Content-based scans (getLastRow / getDataRange) see
  // exactly the survivors afterward.
  // -------------------------------------------------------------------------------
  function _rewriteDataRows(sheet, rows) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
  }

  // -------------------------------------------------------------------------------
  // _alertAdminPruneFailure(ended, archiveName, err)
  // Rev 3.3: a failed prune leaves an archive that still exists and still looks
  // authoritative but holds the next shift's work under this shift's name — the
  // exact bug rev 3.3 exists to fix. Degrades to rev-3.2 behavior; nothing is
  // lost. Best-effort like its siblings: nothing here may throw.
  // -------------------------------------------------------------------------------
  function _alertAdminPruneFailure(ended, archiveName, err) {
    try {
      MailApp.sendEmail({
        to      : CONFIG.ADMIN_EMAIL,
        subject : `⚠️ Parts Dept — archive NOT pruned to shift (${ended.shift} ${ended.target.key})`,
        body    : `The turnover archive "${archiveName}" was created and the clean slate ` +
                  `completed normally, but it could not be pruned to the ${ended.shift} shift.\n\n` +
                  `Error: ${err.message}\n\n` +
                  'The archive is COMPLETE but OVER-INCLUSIVE: it also holds work belonging to ' +
                  'the shift that came on after the boundary. NOTHING WAS LOST — that work also ' +
                  'appears in its own shift\'s archive. Treat counts taken from this file as ' +
                  'unreliable until it is pruned by hand or regenerated.',
        name    : 'Parts Request Workflow',
      });
    } catch (e2) {
      console.error('Turnover._alertAdminPruneFailure: fallback alert also failed:', e2.message);
    }
  }

  // -------------------------------------------------------------------------------
  // _alertAdminFormCleanupFailure(ended, archiveName, err)
  // Rev 3.2: a failed duplicated-form cleanup leaves a live form writing into
  // an archived snapshot — an invariant violation that must reach a human.
  // Best-effort like its sibling: nothing here may throw or abort the turnover.
  // -------------------------------------------------------------------------------
  function _alertAdminFormCleanupFailure(ended, archiveName, err) {
    try {
      MailApp.sendEmail({
        to      : CONFIG.ADMIN_EMAIL,
        subject : `⚠️ Parts Dept — duplicated form NOT cleaned up (${ended.shift} ${ended.target.key})`,
        body    : `The turnover archive "${archiveName}" was created, but the duplicate Google Form ` +
                  `that makeCopy() generated could not be unlinked/trashed.\n\nError: ${err.message}\n\n` +
                  'ACTION NEEDED: find the stray "Copy of ..." form next to the original form, open ' +
                  'Responses → unlink it from the archived sheet, then trash it. Until then, a ' +
                  'submission to that form would write into the archived snapshot.\n\n' +
                  'If the error mentions permissions/scopes, the trigger owner must re-authorize the ' +
                  'script from the editor (the .../auth/forms scope was added in rev 3.2).',
        name    : 'Parts Request Workflow',
      });
    } catch (e2) {
      console.error('Turnover._alertAdminFormCleanupFailure: fallback alert also failed:', e2.message);
    }
  }

  // -------------------------------------------------------------------------------
  // _alertAdminMailFailure(ended, err)
  // The report failing to send must never be silent — mail a plain-text alert
  // to ADMIN_EMAIL. Best-effort: if even this fails, log and move on (the
  // archive is already committed; nothing here may throw).
  // -------------------------------------------------------------------------------
  function _alertAdminMailFailure(ended, err) {
    try {
      MailApp.sendEmail({
        to      : CONFIG.ADMIN_EMAIL,
        subject : `⚠️ Parts Dept — shift report FAILED to send (${ended.shift} ${ended.target.key})`,
        body    : `The ${ended.shift} shift report for op day ${ended.target.key} failed to send ` +
                  `to the manager list.\n\nError: ${err.message}\n\n` +
                  'The archive and clean slate completed normally — the workbook snapshot and its ' +
                  'full event history are intact in the backup folder. The report can be ' +
                  'regenerated from the archived copy if needed.',
        name    : 'Parts Request Workflow',
      });
    } catch (e2) {
      console.error('Turnover._alertAdminMailFailure: fallback alert also failed:', e2.message);
    }
  }

  // -------------------------------------------------------------------------------
  // _endedShift(now) → { shift, target:{ key, date } }
  // The shift that most recently ENDED, relative to "now":
  //   live shift NIGHT → the DAY shift of the SAME op day ended at 17:00.
  //     (Holds even after midnight: opDayKey() folds 00:00–06:00 into the prior
  //      calendar date's op day, so "same op day" is still correct.)
  //   live shift DAY   → the NIGHT shift of the PREVIOUS op day ended at 06:00.
  //     DST-SAFE: derived by calendar-day arithmetic on the op-day key (the key
  //     IS the calendar date the op day started), anchored at noon so the
  //     23/25-hour DST transition days can't skew it — never epoch-ms minus 24h.
  //
  // NOTE: this infers the boundary from the CLOCK. A run that somehow fires
  // BEFORE its boundary resolves to the previous one, where _alreadyRan() then
  // swallows it silently. That is why the trigger schedule keeps its wide
  // post-boundary margin — see the rev 3.3 header note.
  // -------------------------------------------------------------------------------
  function _endedShift(now) {
    if (ShiftReport.shiftOf(now) === ShiftReport.SHIFTS.NIGHT) {
      return {
        shift  : ShiftReport.SHIFTS.DAY,
        target : { key: ShiftReport.opDayKey(now), date: now },
      };
    }
    const prevDate = _dateFromKey(ShiftReport.opDayKey(now), -1);   // previous op day, at noon
    return {
      shift  : ShiftReport.SHIFTS.NIGHT,
      target : { key: _keyFromDate(prevDate), date: prevDate },
    };
  }

  // 'yyyy-MM-dd' key → local Date at NOON of that calendar day (+ offsetDays).
  // Noon anchoring makes ±1-day arithmetic immune to DST transitions.
  function _dateFromKey(key, offsetDays) {
    const p = key.split('-');
    return new Date(+p[0], +p[1] - 1, +p[2] + (offsetDays || 0), 12);
  }

  function _keyFromDate(date) {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
  }

  // Parse an ISO-ish stamp ('YYYY-MM-DDTHH:MM:SS') or a real Date → Date.
  // Component parse avoids the UTC interpretation new Date(str) applies to a
  // zoneless date-time; falls back to the engine parser for anything else.
  function _parseTs(raw) {
    if (raw instanceof Date) return isNaN(raw.getTime()) ? null : raw;
    const s = String(raw).trim();
    if (!s) return null;
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  // -------------------------------------------------------------------------------
  // Once-per-(op day, shift) guard
  // -------------------------------------------------------------------------------
  function _runKey(ended) { return `${ended.target.key}|${ended.shift}`; }

  function _alreadyRan(ended) {
    const last = PropertiesService.getScriptProperties().getProperty('TURNOVER_LAST_RUN');
    return last === _runKey(ended);
  }

  function _stampRun(ended) {
    PropertiesService.getScriptProperties().setProperty('TURNOVER_LAST_RUN', _runKey(ended));
  }

  // -------------------------------------------------------------------------------
  // Dated-folder helpers
  // Walks/creates: BackupRoot / YYYY / MM / Week N, relative to the spreadsheet's
  // parent folder. `date` is the ARCHIVED op day so files land in the right tree.
  // -------------------------------------------------------------------------------
  function _getOrCreateDatedFolder(date) {
    const tz    = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const year  = Utilities.formatDate(date, tz, 'yyyy');
    const month = Utilities.formatDate(date, tz, 'MM');

    let weekLabel;
    if (CONFIG.BACKUP_WEEK_OF_MONTH) {
      const dayOfMonth = parseInt(Utilities.formatDate(date, tz, 'd'), 10);
      weekLabel = `Week ${Math.ceil(dayOfMonth / 7)}`;
    } else {
      weekLabel = `W${Utilities.formatDate(date, tz, 'w').padStart(2, '0')}`;
    }

    const ssFile   = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
    const ssParent = ssFile.getParents().next();

    const backupRoot  = _findOrCreateFolder(ssParent,   CONFIG.BACKUP_FOLDER);
    const yearFolder  = _findOrCreateFolder(backupRoot,  year);
    const monthFolder = _findOrCreateFolder(yearFolder,  month);
    return _findOrCreateFolder(monthFolder, weekLabel);
  }

  function _findOrCreateFolder(parent, name) {
    const iter = parent.getFoldersByName(name);
    return iter.hasNext() ? iter.next() : parent.createFolder(name);
  }

  // -------------------------------------------------------------------------------
  // Public surface
  // -------------------------------------------------------------------------------
  return { run };

})();

// Top-level stub for menu item and BOTH time-driven triggers (17:00 and 06:00)
function Turnover_run() { Turnover.run(); }

// Legacy alias — keeps any old menu entries or stale triggers pointing at
// Eod_runEod working until Setup.initialize() re-installs the trigger set.
function Eod_runEod() { Turnover.run(); }
