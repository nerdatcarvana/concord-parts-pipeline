/**
 * Rebuild.gs
 * One-call regeneration of the entire workbook artifact from code.
 *
 * Purpose ("crystallize so it can be regenerated"): if a fatal mistake corrupts
 * the workbook — a deleted sheet, mangled protections, wrecked formatting —
 * point the script at the (damaged or blank) spreadsheet and run this once to
 * get a fully working system back. It composes the existing idempotent builders
 * in strict dependency order.
 *
 * What it does NOT touch:
 *   _EventLog and _Requests (the permanent event history) and OPEN board rows are
 *   PRESERVED — this rebuilds STRUCTURE, not history.
 *
 * NOTE: the former reporting layer (Analytics / Calc_Data / Dashboard) has been
 * retired in favor of the emailed ShiftReport, which is stateless — there is
 * nothing to rebuild for it. Reporting "recovery" is simply the next EOD run
 * or a menu-driven "📧 Email Shift Report".
 *
 * Phases (each idempotent; later phases depend on earlier ones):
 *   1. Structural spine — Setup.bootstrapStructural: sheets, headers, hidden
 *      flags, Processing row budget, triggers, gid registry (self-heals drifted
 *      tab names first).
 *   2. Presentation/access — Formatting (+ conditional formatting) → Protections.
 *
 * Public API:
 *   Rebuild.everything()
 */

const Rebuild = (() => {

  function everything() {
    let ui = null;
    try { ui = SpreadsheetApp.getUi(); } catch (_) { /* time-driven / headless */ }

    if (ui) {
      const resp = ui.alert(
        '♻️ FULL WORKBOOK REBUILD',
        'Reconstructs every managed sheet, trigger, protection, and format rule from code.\n\n' +
        'Event history (_EventLog, _Requests) and OPEN board rows are PRESERVED.\n\n' +
        'Run as the script OWNER. Proceed?',
        ui.ButtonSet.YES_NO
      );
      if (resp !== ui.Button.YES) return;
    }

    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const log = [];

    // Run a named phase; on failure, record it and stop the sequence (later
    // phases depend on earlier ones, so continuing would cascade noise).
    const phase = (label, fn) => {
      try {
        fn();
        log.push(`✓ ${label}`);
        return true;
      } catch (e) {
        const detail = (e && e.message) ? e.message : String(e);
        log.push(`✗ ${label} — ${detail}`);
        console.error(`Rebuild halted at "${label}": ${(e && e.stack) ? e.stack : detail}`);
        return false;
      }
    };

    const ok =
      // 1. Structural spine.
      phase('Structural bootstrap (sheets, headers, triggers, gid registry)',
            () => Setup.bootstrapStructural(ss)) &&
      // 2. Presentation + access.
      phase('Formatting + conditional formatting', () => Formatting.applyFormatting()) &&
      phase('Protections', () => Access.applyProtections());

    SpreadsheetApp.flush();

    if (ok) {
      const msg = 'Workbook fully rebuilt:\n\n' + log.join('\n');
      if (ui) ui.alert('✅ Rebuild Complete', msg, ui.ButtonSet.OK);
      else Utils.toast('Workbook rebuilt.', '♻️ Rebuild', 6);
    } else {
      const msg = 'Rebuild stopped at a failing phase:\n\n' + log.join('\n') +
                  '\n\nFix the cause and re-run — the rebuild is idempotent.';
      if (ui) ui.alert('⚠️ Rebuild Incomplete', msg, ui.ButtonSet.OK);
      else console.error(msg);
    }
  }

  return { everything };

})();

// Top-level stub for the menu item
function Rebuild_everything() { Rebuild.everything(); }
