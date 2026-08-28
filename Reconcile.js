/**
 * Reconcile.gs
 * Supersedes re-fulfilled discards so discard metrics reflect reality.
 *
 * THE PROBLEM: a request gets DISCARDED, the same stock number is submitted
 * again later in the shift and COMPLETED. Both terminal rows sit in _Requests,
 * so the shift report counts the work as discarded AND completed — inflating
 * the discard percentage for work that was ultimately fulfilled.
 *
 * THE FIX: a DISCARDED row whose stock number has a COMPLETED row with a
 * LATER complete_ts is flipped to status SUPERSEDED (a terminal status that
 * the report does not count as discarded). The flip is audited: the row's
 * notes gain a breadcrumb and a SUPERSEDED event is appended to _EventLog.
 *
 * Matching rule (deliberately time-ordered):
 *   DISCARDED(stock S, t1)  is superseded by  COMPLETED(stock S, t2)  iff t2 > t1.
 *   • One completion can supersede MULTIPLE earlier discards of the same
 *     stock — they were all attempts at the same ultimately-fulfilled work.
 *   • A discard AFTER the completion is NOT superseded — that is genuinely
 *     new discarded work and must still count.
 *   • Rows with unparsable timestamps are left untouched (never guess).
 *
 * WHEN IT RUNS:
 *   • Automatically at every shift turnover, INSIDE the turnover lock and
 *     BEFORE the stats snapshot — so every emailed report is already clean
 *     and the archive copy carries the corrected statuses.
 *   • Manually from the manager menu (♻️ Reconcile Re-fulfilled Discards)
 *     any time mid-shift, so live data views read clean too.
 *
 * WHY DIRECT CELL WRITES (not Ledger.upsertRequest): the upsert merge maps
 * absent fields to '' and overwrites anything defined, so a status-only
 * upsert would blank bins, picker, cycle time, and complete_ts. The flip
 * therefore writes exactly two cells (status, notes) by name-mapped column.
 *
 * Idempotent: SUPERSEDED rows are terminal and never re-match; running twice
 * changes nothing. Post-clean-slate scope: _Requests only holds the current
 * shift (+ carryover), so scans stay small and the sweep is cheap.
 *
 * Public API:
 *   Reconcile.reconcile() → count of rows superseded. LOCK-FREE — caller
 *                           must hold the script lock (Turnover does).
 *   Reconcile.run()       → interactive/trigger entry: takes the lock,
 *                           reconciles, toasts/alerts the result.
 */

const Reconcile = (() => {

  // -------------------------------------------------------------------------------
  // reconcile() → number of discards superseded
  // Caller MUST hold the script lock (mirrors the Intake.handleWalkIn pattern:
  // lock discipline lives at the entry points, never nested).
  // -------------------------------------------------------------------------------
  function reconcile() {
    const H  = CONFIG.HEADERS;
    const E  = CONFIG.ENUMS;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheet = Sheets.getByName(CONFIG.SHEETS.REQUESTS, ss);
    if (!sheet || sheet.getLastRow() < 2) return 0;

    const data  = sheet.getDataRange().getValues();
    const clean = data[0].map(Utils.cleanHeader);
    const col = (h) => clean.indexOf(Utils.cleanHeader(h));

    const iId     = col(H.RQ_REQUEST_ID);
    const iStock  = col(H.RQ_STOCK);
    const iStatus = col(H.RQ_STATUS);
    const iComp   = col(H.RQ_COMPLETE_TS);
    const iNotes  = col(H.RQ_NOTES);
    if (iStock === -1 || iStatus === -1 || iComp === -1) {
      console.warn('Reconcile: required _Requests columns missing — nothing done.');
      return 0;
    }

    // --- Pass 1: index completions by normalized stock number ---
    // completionsByStock: stock → [{ ts, id }] (only rows with parseable ts)
    const completionsByStock = {};
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][iStatus]).trim() !== E.STATUS_COMPLETED) continue;
      const stock = _normStock(data[i][iStock]);
      const ts    = _parseTs(data[i][iComp]);
      if (!stock || !ts) continue;
      (completionsByStock[stock] = completionsByStock[stock] || []).push({
        ts : ts,
        id : iId !== -1 ? String(data[i][iId]).trim() : '',
      });
    }
    if (Object.keys(completionsByStock).length === 0) return 0;

    // --- Pass 2: flip DISCARDED rows that a later completion supersedes ---
    let flipped = 0;
    const tz = ss.getSpreadsheetTimeZone();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][iStatus]).trim() !== E.STATUS_DISCARDED) continue;

      const stock     = _normStock(data[i][iStock]);
      const discardTs = _parseTs(data[i][iComp]);
      if (!stock || !discardTs) continue;                      // never guess on bad stamps

      const laters = (completionsByStock[stock] || [])
        .filter(c => c.ts.getTime() > discardTs.getTime());
      if (laters.length === 0) continue;

      // Earliest later completion is the superseder (the re-submission that fulfilled it).
      laters.sort((a, b) => a.ts.getTime() - b.ts.getTime());
      const by     = laters[0];
      const rowNum = i + 1;
      const stamp  = Utilities.formatDate(by.ts, tz, 'yyyy-MM-dd HH:mm');

      // Two-cell write; see header for why not Ledger.upsertRequest.
      sheet.getRange(rowNum, iStatus + 1).setValue(E.STATUS_SUPERSEDED);
      if (iNotes !== -1) {
        const prev = String(data[i][iNotes] || '').trim();
        const note = `Superseded by ${by.id || 'completion'} @ ${stamp}`;
        sheet.getRange(rowNum, iNotes + 1).setValue(prev ? `${prev} | ${note}` : note);
      }

      // Audit trail — appendEvent auto-fills event_id and event_ts.
      Ledger.appendEvent({
        request_id   : iId !== -1 ? String(data[i][iId]).trim() : '',
        event_type   : E.EVENT_SUPERSEDED,
        actor_email  : Session.getEffectiveUser().getEmail(),
        actor_display: 'System (Reconcile)',
        stock_number : stock,
        notes        : `Discard superseded by later completion ${by.id} @ ${stamp}`,
      });

      flipped++;
    }

    if (flipped > 0) {
      console.log(`Reconcile: superseded ${flipped} re-fulfilled discard(s).`);
      Ledger.clearCache();
    }
    return flipped;
  }

  // -------------------------------------------------------------------------------
  // run() — menu / time-trigger entry point. Owns the lock.
  // -------------------------------------------------------------------------------
  function run() {
    let ui = null;
    try { ui = SpreadsheetApp.getUi(); } catch (_) { /* time-driven — no UI */ }

    let flipped = 0;
    Utils.withLock(() => { flipped = reconcile(); });

    const msg = flipped > 0
      ? `${flipped} discarded request(s) were later fulfilled and have been marked SUPERSEDED. ` +
        'They no longer count toward the discard rate.'
      : 'No re-fulfilled discards found — discard counts are already accurate.';
    if (ui) ui.alert('♻️ Reconcile Complete', msg, ui.ButtonSet.OK);
    else    Utils.toast(msg, '♻️ Reconcile', 5);
  }

  // -------------------------------------------------------------------------------
  // Helpers
  // -------------------------------------------------------------------------------
  function _normStock(raw) {
    return String(raw === null || raw === undefined ? '' : raw).trim().toUpperCase();
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
  // Public surface
  // -------------------------------------------------------------------------------
  return { reconcile, run };

})();

// Top-level stub for the menu item (and an optional hourly time trigger)
function Reconcile_run() { Reconcile.run(); }
