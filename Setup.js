/**
 * Setup.gs
 * Idempotent bootstrap. Run once as the owner after every `clasp push`.
 * Safe to re-run — creates what's missing, skips what already exists.
 *
 * Steps:
 *   1. Create / verify all sheets with correct headers.
 *   2. Hide reference sheets (_Identity, _Roles, _EventLog, _Requests).
 *   3. Ensure Processing has the minimum row budget.
 *   4. Install / replace triggers (onEdit, onFormSubmit, health monitor,
 *      shift turnover ×2). Also deletes retired triggers (Eod_runEod,
 *      ShiftReport_sendDayShift) left over from earlier versions.
 *   5. Apply protections (Access.applyProtections).
 *   6. Apply formatting (Formatting.applyFormatting).
 *
 * Public API:
 *   Setup.initialize()
 */

const Setup = (() => {

  // -------------------------------------------------------------------------------
  // Sheet schema definitions — resolved at call time inside initialize()
  // -------------------------------------------------------------------------------
  function _getSheetSchemas() {
    const S = CONFIG.SHEETS;
    const H = CONFIG.HEADERS;
    return {
      [S.PROCESSING]: [
        H.PROC_REQUESTED,
        H.PROC_REQUESTER,
        H.PROC_STOCK,
        H.PROC_BUILD,
        H.PROC_LOT,
        H.PROC_BIN,
        H.PROC_REQ_ID,
        H.PROC_NOTES,
      ],
      [S.EVENT_LOG]: [
        H.EL_EVENT_ID,
        H.EL_REQUEST_ID,
        H.EL_EVENT_TYPE,
        H.EL_EVENT_TS,
        H.EL_ACTOR_EMAIL,
        H.EL_ACTOR_DISPLAY,
        H.EL_INTAKE_SRC,
        H.EL_STOCK,
        H.EL_BUILD,
        H.EL_LOT,
        H.EL_BIN_CODE,
        H.EL_LOC_TYPE,
        H.EL_REQ_DEPT,
        H.EL_RAW_INPUT,
        H.EL_NOTES,
      ],
      [S.REQUESTS]: [
        H.RQ_REQUEST_ID,
        H.RQ_INTAKE_SRC,
        H.RQ_REQUESTED_TS,
        H.RQ_REQ_EMAIL,
        H.RQ_REQ_DISPLAY,
        H.RQ_REQ_DEPT,
        H.RQ_STOCK,
        H.RQ_BUILD,
        H.RQ_LOT,
        H.RQ_BIN_COUNT,
        H.RQ_BINS,
        H.RQ_FIRST_PICK_TS,
        H.RQ_COMPLETE_TS,
        H.RQ_PICKER_EMAIL,
        H.RQ_PICKER_DISP,
        H.RQ_CYCLE_MIN,
        H.RQ_STATUS,
        H.RQ_NOTES,
      ],
      [S.IDENTITY]: [
        H.ID_ALIAS,
        H.ID_CANONICAL,
        H.ID_DISPLAY,
        H.ID_DEPARTMENT,
        H.ID_DROP_LOC,
        H.ID_ACTIVE,
      ],
      [S.ROLES]: [
        H.RO_EMAIL,
        H.RO_ROLE,
        H.RO_NOTES,
      ],
    };
  }

  const HIDDEN_SHEETS = ['_Identity', '_Roles', '_EventLog', '_Requests'];

  // -------------------------------------------------------------------------------
  // bootstrapStructural(ss)
  // Non-interactive structural core, shared by initialize() and Rebuild.everything().
  // Self-heals drifted tab names FIRST (so the name-based create/adopt below can't
  // spawn a duplicate of a renamed sheet), then ensures sheets/headers, hides
  // reference sheets, sets the row budget, installs triggers, and registers every
  // sheet's gid so resolution is rename-proof thereafter.
  // -------------------------------------------------------------------------------
  function bootstrapStructural(ss) {
    ss = ss || SpreadsheetApp.getActiveSpreadsheet();
    Sheets.repairAll(ss);
    _ensureSheets(ss);
    _hideReferenceSheets(ss);
    _ensureProcessingRowBudget(ss);
    _installTriggers(ss);
    Sheets.syncRegistry(ss);
  }

  // -------------------------------------------------------------------------------
  // initialize()
  // -------------------------------------------------------------------------------
  function initialize() {
    const ui = _getUiSafe();

    if (ui) {
      const resp = ui.alert(
        '🔧 Run Setup',
        'This will create/verify all sheets, reinstall triggers, reapply protections and formatting.\n\nRun as the script OWNER. Proceed?',
        ui.ButtonSet.YES_NO
      );
      if (resp !== ui.Button.YES) return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    bootstrapStructural(ss);
    Access.applyProtections();
    Formatting.applyFormatting();

    SpreadsheetApp.flush();

    if (ui) {
      ui.alert('✅ Setup Complete', 'All sheets, triggers, protections, and formatting are in place.', ui.ButtonSet.OK);
    }
  }

  // -------------------------------------------------------------------------------
  // _ensureSheets(ss)
  // -------------------------------------------------------------------------------
  function _ensureSheets(ss) {
    const schemas       = _getSheetSchemas();
    const existingNames = ss.getSheets().map(s => s.getName());

    for (const [sheetName, headers] of Object.entries(schemas)) {
      let sheet;
      if (!existingNames.includes(sheetName)) {
        sheet = ss.insertSheet(sheetName);
      } else {
        sheet = ss.getSheetByName(sheetName);
      }
      _ensureHeaders(sheet, headers);
    }
  }

  // -------------------------------------------------------------------------------
  // _ensureHeaders(sheet, expectedHeaders)
  // -------------------------------------------------------------------------------
  function _ensureHeaders(sheet, expectedHeaders) {
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      return;
    }

    const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(Utils.cleanHeader);
    const missing  = expectedHeaders.filter(h => !existing.includes(Utils.cleanHeader(h)));
    if (missing.length > 0) {
      const startCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
      console.warn(`Setup._ensureHeaders: appended missing columns to "${sheet.getName()}": ${missing.join(', ')}`);
    }
  }

  // -------------------------------------------------------------------------------
  // _hideReferenceSheets(ss)
  // -------------------------------------------------------------------------------
  function _hideReferenceSheets(ss) {
    HIDDEN_SHEETS.forEach(name => {
      const sheet = Sheets.getByName(name, ss);
      if (sheet) sheet.hideSheet();
    });
  }

  // -------------------------------------------------------------------------------
  // _ensureProcessingRowBudget(ss)
  // -------------------------------------------------------------------------------
  function _ensureProcessingRowBudget(ss) {
    const sheet = Sheets.get('PROCESSING', ss);
    if (!sheet) return;
    const MIN = CONFIG.PROCESSING_MIN_ROWS + 1;
    if (sheet.getMaxRows() < MIN) {
      sheet.insertRowsAfter(sheet.getMaxRows(), MIN - sheet.getMaxRows());
    }
  }

  // -------------------------------------------------------------------------------
  // _installTriggers(ss)
  // Clears all existing managed triggers then reinstalls them.
  //
  // Trigger inventory:
  //   onEditTrigger           — spreadsheet onEdit        — picker board interactions
  //   onFormSubmitTrigger     — spreadsheet onFormSubmit  — auto-sync on form submission
  //   onFormSubmitHeartbeat   — spreadsheet onFormSubmit  — schedules a delayed sync re-check
  //   Health_monitor          — time-driven HOURLY        — backstop sweep (stuck rows + unmapped aliases);
  //                                                         the 2-min heartbeat is the primary safety net
  //   Turnover_run            — time-driven daily ~17:00  — DAY-shift turnover: report + archive + purge,
  //                                                         clean _Requests slate for nightshift
  //   Turnover_run            — time-driven daily ~6 AM   — NIGHT-shift turnover: report + archive + purge,
  //                                                         clean _Requests slate for dayshift
  //
  // Retired (deleted on every run, never reinstalled): Eod_runEod,
  // ShiftReport_sendDayShift — both boundaries are handled by Turnover now.
  //
  // Both turnover triggers pull their hours from CONFIG.SHIFT so a
  // shift-boundary change propagates here automatically on the next Setup run.
  // -------------------------------------------------------------------------------
  function _installTriggers(ss) {
    const TRIGGER_FUNCTIONS = [
      'onFormSubmitHeartbeat',
      'Health_heartbeatCheck',
      'onEditTrigger',
      'onFormSubmitTrigger',
      'Health_monitor',
      'Turnover_run',
      // retired handlers — listed so stale triggers from earlier versions get removed
      'Eod_runEod',
      'ShiftReport_sendDayShift',
    ];

    // Remove all existing managed triggers to prevent duplicates on re-run
    ScriptApp.getProjectTriggers().forEach(trigger => {
      if (TRIGGER_FUNCTIONS.includes(trigger.getHandlerFunction())) {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    // Picker board interactions: walk-in stock entry, bin location pick gate
    ScriptApp.newTrigger('onEditTrigger')
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    // Auto-sync: fires the moment a Google Form response lands in FR1.
    // This is what drives real-time form intake without manual menu intervention.
    ScriptApp.newTrigger('onFormSubmitTrigger')
      .forSpreadsheet(ss)
      .onFormSubmit()
      .create();

    // Heartbeat: fires on each form submit, schedules a 2-min delayed check
    // to verify the sync trigger actually ran. Safety net for silent failures.
    ScriptApp.newTrigger('onFormSubmitHeartbeat')
      .forSpreadsheet(ss)
      .onFormSubmit()
      .create();

    // Health monitor backstop: stuck responses + unmapped alias backlog.
    // HOURLY (was every 5 min): the 2-minute post-submit heartbeat is the
    // primary net for stuck rows; this sweep only has to catch what the
    // heartbeat itself missed, which does not need 288 executions a day.
    ScriptApp.newTrigger('Health_monitor')
      .timeBased()
      .everyHours(1)
      .create();

    // Shift turnover — DAY boundary: fires within [SECOND_HR, SECOND_HR+1),
    // just after the 6:00–17:00 day shift ends. Reports, archives, and purges
    // dayshift's finished work so nightshift starts with a clean _Requests.
    ScriptApp.newTrigger('Turnover_run')
      .timeBased()
      .atHour(CONFIG.SHIFT.SECOND_HR)
      .everyDays(1)
      .inTimezone(ss.getSpreadsheetTimeZone())
      .create();

    // Shift turnover — NIGHT boundary: fires within [START_HR, START_HR+1),
    // just after the 17:00–6:00 night shift ends. Reports, archives, and purges
    // nightshift's finished work so dayshift starts with a clean _Requests.
    ScriptApp.newTrigger('Turnover_run')
      .timeBased()
      .atHour(CONFIG.SHIFT.START_HR)
      .everyDays(1)
      .inTimezone(ss.getSpreadsheetTimeZone())
      .create();
  }

  // -------------------------------------------------------------------------------
  // _getUiSafe()
  // -------------------------------------------------------------------------------
  function _getUiSafe() {
    try { return SpreadsheetApp.getUi(); } catch (_) { return null; }
  }

  // -------------------------------------------------------------------------------
  // Public surface
  // -------------------------------------------------------------------------------
  return { initialize, bootstrapStructural };

})();

// Top-level stubs
function Setup_initialize() { Setup.initialize(); }

// -------------------------------------------------------------------------------
// onFormSubmitTrigger(e)
// Installed as an onFormSubmit trigger by Setup.initialize().
// Fires automatically when a Google Form response is submitted.
// Calls the same syncFormResponses() used by the manual menu item —
// no special handling needed, the idempotency flag prevents double-processing.
// -------------------------------------------------------------------------------
function onFormSubmitTrigger(e) {
  Intake.syncFormResponses();
}
