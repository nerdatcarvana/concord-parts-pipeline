/**
 * Access.gs
 * Role-based access control, menu construction, and sheet protections.
 *
 * Public API:
 *   Access.currentUserEmail()    → string email of the active user
 *   Access.isManager(email)      → boolean
 *   Access.getManagerEmails()    → string[] of MANAGER emails from _Roles
 *   Access.buildMenu()           → constructs the UI menu (called from onOpen)
 *   Access.applyProtections()    → locks all non-entry ranges to owner + managers
 */

const Access = (() => {

  // --------------------------------------------------------------------------
  // currentUserEmail
  // Returns the active user's email. In a time-driven trigger this returns ''.
  // Only use this for interactive (onEdit / menu) paths.
  // --------------------------------------------------------------------------
  function currentUserEmail() {
    return Session.getActiveUser().getEmail();
  }

  // --------------------------------------------------------------------------
  // getManagerEmails() — reads _Roles, returns array of MANAGER emails
  // --------------------------------------------------------------------------
  function getManagerEmails() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = Sheets.getByName(CONFIG.SHEETS.ROLES, ss);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(Utils.cleanHeader);
    const emailIdx = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RO_EMAIL));
    const roleIdx  = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RO_ROLE));
    if (emailIdx === -1 || roleIdx === -1) return [];

    return data.slice(1)
      .filter(row => String(row[roleIdx]).trim().toUpperCase() === CONFIG.ENUMS.ROLE_MANAGER)
      .map(row => String(row[emailIdx]).trim().toLowerCase())
      .filter(e => e !== '');
  }

  // --------------------------------------------------------------------------
  // isManager(email)
  // --------------------------------------------------------------------------
  function isManager(email) {
    if (!email) return false;
    const managers = getManagerEmails();
    return managers.includes(String(email).trim().toLowerCase());
  }

  // --------------------------------------------------------------------------
  // buildMenu()
  // Constructs a role-appropriate menu. Pickers get minimal items;
  // managers get the full set including EOD archive and maintenance tools.
  // Safe to call from onOpen (simple trigger) — no lock, no heavy reads.
  // --------------------------------------------------------------------------
  function buildMenu() {
    const ui   = SpreadsheetApp.getUi();
    const menu = ui.createMenu('🛠️ Parts Management');
    const user = currentUserEmail();

    // Picker items — always visible
    menu.addItem('📥 Sync Form Submissions', 'Intake_syncFormResponses')
    .addItem('📧 Email Shift Report', 'ShiftReport_emailMe');

    if (isManager(user)) {
      menu
        .addSeparator()
        .addItem('🗄️ End of Day Archive', 'Eod_runEod')
        .addSeparator()
        .addItem('🔁 Re-apply Formatting', 'Formatting_applyFormatting')
        .addItem('🔒 Re-apply Protections', 'Access_applyProtections')
        .addSeparator()
        .addItem('👤 Clear Unmapped Alias Queue', 'Identity_clearUnmapped')
        .addItem('🔧 Re-run Setup (initialize)', 'Setup_initialize')
        .addSeparator()
        .addItem('♻️ Rebuild Entire Workbook', 'Rebuild_everything')
	      .addItem('♻️ Reconcile Re-fulfilled Discards', 'Reconcile_run');
    }

    menu.addToUi();
  }

  // --------------------------------------------------------------------------
  // applyProtections()
  // Locks every sheet except the picker-editable columns on Processing.
  // Idempotent: removes all existing protections before re-applying.
  //
  // Protected (owner + managers only):
  //   _EventLog, _Requests, _Identity, _Roles   — entire sheet
  //   Processing                                 — all columns EXCEPT those below
  //
  // Picker-editable on Processing:
  //   Bin Location  — the pick entry (always)
  //   Stock Number  — walk-in entry
  //   Requester     — walk-in entry
  //   Notes         — free-text annotation
  // --------------------------------------------------------------------------
  function applyProtections() {
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const managers = getManagerEmails();
    const owner    = Session.getEffectiveUser().getEmail();

    // Build the allowed-editors list: owner is implicit as script runner;
    // addEditor() calls add on top of the sheet-level permission.
    const privilegedEditors = managers.filter(e => e !== owner.toLowerCase());

    // --- Clear all existing protections first ---
    ss.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    ss.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

    // --- Protect ledger and reference sheets entirely ---
    const lockedSheets = [
      CONFIG.SHEETS.EVENT_LOG,
      CONFIG.SHEETS.REQUESTS,
      CONFIG.SHEETS.IDENTITY,
      CONFIG.SHEETS.ROLES,
    ];
    lockedSheets.forEach(name => {
      const sheet = Sheets.getByName(name, ss);
      if (!sheet) return;
      const prot = sheet.protect().setDescription(`Protected: ${name}`);
      prot.removeEditors(prot.getEditors());
      if (privilegedEditors.length > 0) prot.addEditors(privilegedEditors);
    });

    // --- Protect Processing: entire sheet, then punch holes for pickers ---
    const procSheet = Sheets.getByName(CONFIG.SHEETS.PROCESSING, ss);
    if (procSheet) {
      const sheetProt = procSheet.protect().setDescription('Processing: sheet-level lock');
      sheetProt.removeEditors(sheetProt.getEditors());
      if (privilegedEditors.length > 0) sheetProt.addEditors(privilegedEditors);

      // Picker-editable columns:
      //   Bin Location  — always (the pick entry)
      //   Stock Number  — walk-in entry
      //   Requester     — walk-in entry (picker types the requester's name/alias)
      const pickerCols = [];
      try { pickerCols.push(Utils.getColIndex(procSheet, CONFIG.HEADERS.PROC_BIN)); }       catch(_) {}
      try { pickerCols.push(Utils.getColIndex(procSheet, CONFIG.HEADERS.PROC_STOCK)); }     catch(_) {}
      try { pickerCols.push(Utils.getColIndex(procSheet, CONFIG.HEADERS.PROC_REQUESTER)); } catch(_) {}
      try { pickerCols.push(Utils.getColIndex(procSheet, CONFIG.HEADERS.PROC_BUILD)); }     catch(_) {}
      try { pickerCols.push(Utils.getColIndex(procSheet, CONFIG.HEADERS.PROC_LOT)); }       catch(_) {}
      try { pickerCols.push(Utils.getColIndex(procSheet, CONFIG.HEADERS.PROC_NOTES)); }     catch(_) {}

      if (pickerCols.length > 0) {
        const dataRows = Math.max(procSheet.getLastRow() - 1, CONFIG.PROCESSING_MIN_ROWS);
        const unprotectedRanges = pickerCols.map(col =>
          procSheet.getRange(2, col, dataRows, 1)
        );
        sheetProt.setUnprotectedRanges(unprotectedRanges);
      }
    }

    Utils.toast('Protections applied.', '🔒 Access', 3);
  }

  // --------------------------------------------------------------------------
  // Public surface
  // --------------------------------------------------------------------------
  return {
    currentUserEmail,
    isManager,
    getManagerEmails,
    buildMenu,
    applyProtections,
  };

})();

// ---------------------------------------------------------------------------
// Top-level stubs required by the menu item strings.
// Apps Script menu items must reference top-level function names.
// ---------------------------------------------------------------------------
function onOpen()                    { Access.buildMenu(); }
function Access_applyProtections()   { Access.applyProtections(); }
