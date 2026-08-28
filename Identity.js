/**
 * Identity.gs
 * Resolves any requester reference (alias, nickname, email, typo) to a
 * canonical {email, display, department} triple via the _Identity sheet.
 *
 * Public API:
 *   Identity.resolve(aliasOrEmail)   → {email, display, department} | UNMAPPED sentinel
 *   Identity.isUnmapped(result)      → boolean
 *   Identity.flagUnmapped(rawValue)  → logs the alias to the unmapped backlog column
 *   Identity.listUnmapped()          → string[] of pending unmapped aliases
 *   Identity.clearUnmapped()         → removes resolved entries (called by manager)
 */

const Identity = (() => {

  // Column name used to hold the unmapped-alias backlog in _Identity.
  // This is an internal system column, not part of the canonical schema.
  const UNMAPPED_COL = '_unmapped_queue';

  // --------------------------------------------------------------------------
  // resolve(aliasOrEmail)
  // --------------------------------------------------------------------------
  function resolve(aliasOrEmail) {
    if (!aliasOrEmail || String(aliasOrEmail).trim() === '') {
      return CONFIG.ENUMS.IDENTITY_UNMAPPED;
    }

    const raw = String(aliasOrEmail).trim().toLowerCase();
    const sheet = _getSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return CONFIG.ENUMS.IDENTITY_UNMAPPED;

    const headers = data[0].map(Utils.cleanHeader);
    const aliasIdx   = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.ID_ALIAS));
    const emailIdx   = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.ID_CANONICAL));
    const displayIdx = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.ID_DISPLAY));
    const deptIdx    = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.ID_DEPARTMENT));
    const activeIdx  = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.ID_ACTIVE));

    if (aliasIdx === -1 || emailIdx === -1) {
      throw new Error('Identity._getSheet: required headers missing from _Identity sheet.');
    }

    // Build a result helper to avoid repetition
    const buildResult = (row) => ({
      email      : String(row[emailIdx]).trim(),
      display    : displayIdx !== -1 ? String(row[displayIdx]).trim() : String(row[emailIdx]).trim(),
      department : deptIdx !== -1 ? String(row[deptIdx]).trim() : '',
    });

    const isActive = (row) => {
      if (activeIdx === -1) return true;
      const active = row[activeIdx];
      return active !== false && String(active).toLowerCase() !== 'false';
    };

    // Pass 1: match against the alias column (canonical lookup)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!isActive(row)) continue;
      const alias = String(row[aliasIdx]).trim().toLowerCase();
      if (alias === raw) return buildResult(row);
    }

    // Pass 2: match against display_name — catches pickers who select from the
    // dropdown (which shows display names) or type a display name directly.
    // Only runs if Pass 1 found nothing, so aliases always take precedence.
    if (displayIdx !== -1) {
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!isActive(row)) continue;
        const display = String(row[displayIdx]).trim().toLowerCase();
        if (display === raw) return buildResult(row);
      }
    }

    flagUnmapped(aliasOrEmail);
    return CONFIG.ENUMS.IDENTITY_UNMAPPED;
  }

  // --------------------------------------------------------------------------
  // isUnmapped(resolveResult)
  // --------------------------------------------------------------------------
  function isUnmapped(result) {
    return result === CONFIG.ENUMS.IDENTITY_UNMAPPED;
  }

  // --------------------------------------------------------------------------
  // flagUnmapped(rawValue)
  // --------------------------------------------------------------------------
  function flagUnmapped(rawValue) {
    if (!rawValue) return;
    const trimmed = String(rawValue).trim();
    if (!trimmed) return;

    const sheet = _getSheet();
    const data  = sheet.getDataRange().getValues();
    const headers = data[0].map(Utils.cleanHeader);

    let queueColIdx = headers.indexOf(Utils.cleanHeader(UNMAPPED_COL));
    if (queueColIdx === -1) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(UNMAPPED_COL);
      queueColIdx = newCol - 1;
    }

    const queueSheetCol = queueColIdx + 1;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const existing = sheet.getRange(2, queueSheetCol, lastRow - 1, 1).getValues().flat();
      if (existing.map(v => String(v).trim().toLowerCase()).includes(trimmed.toLowerCase())) {
        return;
      }
    }

    let targetRow = lastRow + 1;
    for (let i = 1; i < data.length; i++) {
      if (!data[i][queueColIdx] || String(data[i][queueColIdx]).trim() === '') {
        targetRow = i + 1;
        break;
      }
    }
    sheet.getRange(targetRow, queueSheetCol).setValue(trimmed);
  }

  // --------------------------------------------------------------------------
  // listUnmapped()
  // --------------------------------------------------------------------------
  function listUnmapped() {
    const sheet = _getSheet();
    const data  = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[0].map(Utils.cleanHeader);
    const queueColIdx = headers.indexOf(Utils.cleanHeader(UNMAPPED_COL));
    if (queueColIdx === -1) return [];

    return data.slice(1)
      .map(row => String(row[queueColIdx]).trim())
      .filter(v => v !== '');
  }

  // --------------------------------------------------------------------------
  // clearUnmapped()
  // --------------------------------------------------------------------------
  function clearUnmapped() {
    const sheet = _getSheet();
    const data  = sheet.getDataRange().getValues();
    const headers = data[0].map(Utils.cleanHeader);
    const queueColIdx = headers.indexOf(Utils.cleanHeader(UNMAPPED_COL));
    if (queueColIdx === -1) return;

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, queueColIdx + 1, lastRow - 1, 1).clearContent();
    }
  }

  // --------------------------------------------------------------------------
  // Internal helpers
  // --------------------------------------------------------------------------
  function _getSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = Sheets.getByName(CONFIG.SHEETS.IDENTITY, ss);
    if (!sheet) throw new Error('Identity: _Identity sheet not found. Run Setup.initialize() first.');
    return sheet;
  }

  // --------------------------------------------------------------------------
  // Public surface
  // --------------------------------------------------------------------------
  return {
    resolve,
    isUnmapped,
    flagUnmapped,
    listUnmapped,
    clearUnmapped,
  };

})();

// Top-level stub for menu item
function Identity_clearUnmapped() { Identity.clearUnmapped(); }
