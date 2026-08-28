/**
 * Formatting.gs
 * All visual styling, data validation, and conditional formatting for the
 * Processing board. Idempotent — safe to re-run after any EOD reset or on
 * demand from the manager menu.
 *
 * RULE PRECEDENCE (Google Sheets semantics): when several conditional format
 * rules match a cell, the FIRST matching rule in the list wins — later rules
 * are ignored for that cell. Rule order in _applyConditionalFormatting is
 * therefore load-bearing: PENDING > aging (crit, warn) > Lot Status chips.
 *
 * PENDING STATE: a row whose Bin Location holds anything renders in the
 * pending tint INSTANTLY (conditional formatting is client-side — no trigger
 * latency). The tint clears itself when the pick execution clears the row or
 * the bin cell; a tint that LINGERS is a visible "this didn't go through"
 * flag (e.g. lock-busy left the bin in place for a retry).
 *
 * Public API:
 *   Formatting.applyFormatting()  → full idempotent format pass on Processing
 */

const Formatting = (() => {

  // --------------------------------------------------------------------------
  // Design tokens
  // --------------------------------------------------------------------------
  const DESIGN = {
    HEADER_BG        : '#1a1a2e',
    HEADER_FG        : '#ffffff',
    HEADER_FONT_SIZE : 10,
    ROW_HEIGHT       : 26,
    BAND_ODD         : '#f8f9fc',
    BAND_EVEN        : '#ffffff',
    FONT_FAMILY      : 'Inter',
    FONT_SIZE        : 10,

    DEPT_COLORS: {
      'Reconditioning' : '#e8f4fd',
      'Body Shop'      : '#fef9e7',
      'Mechanical'     : '#e9f7ef',
      'Paint'          : '#fdf2f8',
      'Detail'         : '#f0f4f8',
      'default'        : '#ffffff',
    },

    // Lot Status chips — semantic fills on the Lot Status column (conditional fmt).
    // Soft palette, same family as DEPT_COLORS so the board reads as one system.
    LOT_COLORS: {
      'Parts Hold'      : '#fff3cd',  // caution / awaiting
      'Sublet Complete' : '#d4edda',  // done
      'Rework'          : '#f8d7da',  // problem
      'Paint Parts'     : '#e7e0f7',  // paint lane
      'Other'           : '#e9ecef',  // neutral
      'Walk In'         : '#d8e8f5',  // walk-in intake
    },

    // Aging escalation backgrounds for stale OPEN-on-board rows (conditional fmt).
    AGE_WARN_BG : '#fde9b8',  // amber — open longer than PROCESSING_STALE_WARN_HRS
    AGE_CRIT_BG : '#f7c5cb',  // red   — open longer than PROCESSING_STALE_CRIT_HRS
    AGE_CRIT_FG : '#7a1620',  // dark red text for contrast on the red row

    // Pending-pick tint: bin entered, trigger execution not yet landed. Instant
    // client-side feedback (no trigger in the loop). Distinct from every lot
    // chip and both aging bands so a glance reads unambiguously.
    PENDING_BG  : '#d7e3fc',  // soft blue — "submitted, processing…"
    PENDING_FG  : '#1b3a6b',  // dark blue text for contrast
  };

  // --------------------------------------------------------------------------
  // applyFormatting()
  // --------------------------------------------------------------------------
  function applyFormatting() {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = Sheets.getByName(CONFIG.SHEETS.PROCESSING, ss);
    if (!sheet) throw new Error('Formatting: Processing sheet not found.');

    const lastCol = sheet.getLastColumn();
    const maxRows = Math.max(sheet.getMaxRows(), CONFIG.PROCESSING_MIN_ROWS + 1);

    // Resolve column widths here (safe — called at runtime, not load time)
    const H = CONFIG.HEADERS;
    const colWidths = {
      [H.PROC_REQUESTED] : 140,
      [H.PROC_REQUESTER] : 160,
      [H.PROC_STOCK]     : 100,
      [H.PROC_BUILD]     : 130,
      [H.PROC_LOT]       : 140,
      [H.PROC_BIN]       : 160,
      [H.PROC_REQ_ID]    : 240,
      [H.PROC_NOTES]     : 280,
    };

    sheet.clearConditionalFormatRules();

    _formatHeader(sheet, lastCol);
    sheet.setFrozenRows(1);
    sheet.setRowHeightsForced(1, maxRows, DESIGN.ROW_HEIGHT);
    _applyColumnWidths(sheet, colWidths);

    // Gate on maxRows, not getLastRow(): after a turnover wipe of an empty
    // board getLastRow() is 1, but the fixed row budget must stay formatted.
    if (maxRows > 1) {
      sheet.getRange(2, 1, maxRows - 1, lastCol)
        .setFontFamily(DESIGN.FONT_FAMILY)
        .setFontSize(DESIGN.FONT_SIZE)
        .setVerticalAlignment('middle');
    }

    _applyBanding(sheet, lastCol, maxRows);
    _applyConditionalFormatting(sheet, lastCol, maxRows);
    _applyDataValidation(sheet, maxRows);
    _applyNotesFormatting(sheet, maxRows);

    SpreadsheetApp.flush();
  }

  // --------------------------------------------------------------------------
  // _formatHeader
  // --------------------------------------------------------------------------
  function _formatHeader(sheet, lastCol) {
    sheet.getRange(1, 1, 1, lastCol)
      .setBackground(DESIGN.HEADER_BG)
      .setFontColor(DESIGN.HEADER_FG)
      .setFontSize(DESIGN.HEADER_FONT_SIZE)
      .setFontWeight('bold')
      .setFontFamily(DESIGN.FONT_FAMILY)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
  }

  // --------------------------------------------------------------------------
  // _applyColumnWidths
  // --------------------------------------------------------------------------
  function _applyColumnWidths(sheet, colWidths) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(Utils.cleanHeader);
    Object.entries(colWidths).forEach(([headerVal, width]) => {
      const idx = headers.indexOf(Utils.cleanHeader(headerVal));
      if (idx !== -1) sheet.setColumnWidth(idx + 1, width);
    });
  }

  // --------------------------------------------------------------------------
  // _applyBanding
  // --------------------------------------------------------------------------
  function _applyBanding(sheet, lastCol, maxRows) {
    sheet.getBandings().forEach(b => b.remove());
    if (maxRows < 2) return;

    sheet.getRange(2, 1, maxRows - 1, lastCol)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false)
      .setFirstRowColor(DESIGN.BAND_ODD)
      .setSecondRowColor(DESIGN.BAND_EVEN)
      .setHeaderRowColor(null)
      .setFooterRowColor(null);
  }

  // --------------------------------------------------------------------------
  // _applyConditionalFormatting
  // --------------------------------------------------------------------------
  function _applyConditionalFormatting(sheet, lastCol, maxRows) {
    if (maxRows < 2) return;

    const H           = CONFIG.HEADERS;
    const lastDataRow = maxRows - 1;
    const dataRange   = sheet.getRange(2, 1, lastDataRow, lastCol);

    // Resolve columns by header name → A1 letter (board layout is position-independent).
    const colLetter = (hc) => {
      try { return _colToLetter(Utils.getColIndex(sheet, hc)); } catch (_) { return null; }
    };
    const reqTsL = colLetter(H.PROC_REQUESTED);
    const binL   = colLetter(H.PROC_BIN);
    const idL    = colLetter(H.PROC_REQ_ID);

    // ORDER IS LOAD-BEARING. Sheets applies the FIRST matching rule per cell
    // and ignores the rest (there is no "later overrides earlier"). Precedence
    // here, top to bottom:
    //   1. PENDING  — bin entered, execution in flight. Wins over everything:
    //      a stale row that just got a bin is being handled NOW.
    //   2. AGING    — crit before warn is not required for correctness (the
    //      bands are mutually exclusive) but whole-row urgency must beat the
    //      per-cell lot chips, so both sit above them.
    //   3. LOT CHIPS — per-value fills on the Lot cell of ordinary open rows.
    // (rev note: chips previously sat FIRST with a comment claiming the aging
    // rules "below" would win on the Lot cell — backwards; a stale row showed
    // a chip-colored hole in its urgency band. Reordered.)
    const rules = [];

    // --- 1. Pending pick (whole row) ------------------------------------------
    // Condition is simply "Bin holds anything": the tint appears the instant
    // the picker commits the cell — zero trigger latency — and disappears when
    // the pick execution clears the row (success), clears the bin (reject), or
    // never, which makes a stuck/lock-busy row visibly demand a retry instead
    // of failing silently. Presentation only: no data, ordering, or lock impact.
    if (binL) {
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=$${binL}2<>""`)
          .setBackground(DESIGN.PENDING_BG)
          .setFontColor(DESIGN.PENDING_FG)
          .setItalic(true)
          .setRanges([dataRange])
          .build()
      );
    }

    // --- 2. Aging escalation on OPEN-on-board rows (whole row) -----------------
    // OPEN-on-board := has a request_id AND no bin entered yet (completed rows are
    // cleared off the board, so this is a reliable "live, unpicked" signal). Age is
    // measured from the Requested timestamp, coerced from ISO text — or a real
    // datetime — to a serial with the same pattern Analytics uses. An unparseable
    // stamp yields age 0 (IFERROR→NOW()), so it never false-flags. NOW() makes
    // these rules volatile; negligible on a ≤~50-row board, and the point is a
    // live-aging board now that open work persists across the op-day boundary.
    if (reqTsL && binL && idL) {
      const serial =
        `IF(ISNUMBER($${reqTsL}2),$${reqTsL}2,` +
        `DATEVALUE(LEFT($${reqTsL}2&"",10))+IFERROR(TIMEVALUE(MID($${reqTsL}2&"",12,8)),0))`;
      const ageHrs = `(NOW()-IFERROR(${serial},NOW()))*24`;
      const open   = `$${idL}2<>"", $${binL}2=""`;

      const WARN = CONFIG.PROCESSING_STALE_WARN_HRS;
      const CRIT = CONFIG.PROCESSING_STALE_CRIT_HRS;

      // Critical band: age > CRIT.
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(${open}, ${ageHrs}>${CRIT})`)
          .setBackground(DESIGN.AGE_CRIT_BG)
          .setFontColor(DESIGN.AGE_CRIT_FG)
          .setRanges([dataRange])
          .build()
      );
      // Warn band: WARN < age <= CRIT. Bounded so it's mutually exclusive with the
      // critical band — final colour is order-independent between the two.
      rules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(${open}, ${ageHrs}>${WARN}, ${ageHrs}<=${CRIT})`)
          .setBackground(DESIGN.AGE_WARN_BG)
          .setRanges([dataRange])
          .build()
      );
    }

    // --- 3. Lot Status semantic chips (per-value, scoped to the Lot column) ----
    // LAST on purpose: first-match-wins means the pending and aging whole-row
    // rules above take the Lot cell on rows they claim, keeping urgency rows a
    // single uniform colour. Chips render only on ordinary open rows.
    let lotCol;
    try { lotCol = Utils.getColIndex(sheet, H.PROC_LOT); } catch (_) { lotCol = null; }
    if (lotCol !== null) {
      const lotRange = sheet.getRange(2, lotCol, lastDataRow, 1);
      Object.entries(DESIGN.LOT_COLORS).forEach(([value, bg]) => {
        rules.push(
          SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(value)
            .setBackground(bg)
            .setRanges([lotRange])
            .build()
        );
      });
    }

    if (rules.length > 0) sheet.setConditionalFormatRules(rules);
  }

  // --------------------------------------------------------------------------
  // _applyDataValidation
  // --------------------------------------------------------------------------
  function _applyDataValidation(sheet, maxRows) {
    if (maxRows < 2) return;
    const dataRows = maxRows - 1;

    _applyDropdown(sheet, CONFIG.HEADERS.PROC_BUILD, dataRows, CONFIG.BUILD_LEVELS);
    _applyDropdown(sheet, CONFIG.HEADERS.PROC_LOT, dataRows, [...CONFIG.LOT_STATUSES, CONFIG.LOT_WALKIN]);

    const requesterNames = _getIdentityDisplayNames();
    if (requesterNames.length > 0) {
      _applyDropdown(sheet, CONFIG.HEADERS.PROC_REQUESTER, dataRows, requesterNames);
    }
  }

  // --------------------------------------------------------------------------
  // _applyDropdown
  // --------------------------------------------------------------------------
  function _applyDropdown(sheet, headerConstant, dataRows, values) {
    let col;
    try { col = Utils.getColIndex(sheet, headerConstant); } catch (_) { return; }

    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, col, dataRows, 1).setDataValidation(validation);
  }

  // --------------------------------------------------------------------------
  // _getIdentityDisplayNames
  // --------------------------------------------------------------------------
  function _getIdentityDisplayNames() {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = Sheets.getByName(CONFIG.SHEETS.IDENTITY, ss);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(Utils.cleanHeader);
    const dispIdx = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.ID_DISPLAY));
    const actIdx  = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.ID_ACTIVE));
    if (dispIdx === -1) return [];

    const names = new Set();
    for (let i = 1; i < data.length; i++) {
      if (actIdx !== -1 && data[i][actIdx] === false) continue;
      const name = String(data[i][dispIdx]).trim();
      if (name) names.add(name);
    }
    return [...names].sort();
  }

  // --------------------------------------------------------------------------
  // _colToLetter(colNum) — 1-based column number → A1 letter(s)
  // --------------------------------------------------------------------------
  function _colToLetter(col) {
    let letter = '';
    while (col > 0) {
      const rem = (col - 1) % 26;
      letter = String.fromCharCode(65 + rem) + letter;
      col = Math.floor((col - 1) / 26);
    }
    return letter;
  }

  // --------------------------------------------------------------------------
  // _applyNotesFormatting
  // Word-wrap the Notes column and align text to the top so long notes are
  // fully readable without expanding rows manually.
  // --------------------------------------------------------------------------
  function _applyNotesFormatting(sheet, maxRows) {
    if (maxRows < 2) return;
    let notesCol;
    try {
      notesCol = Utils.getColIndex(sheet, CONFIG.HEADERS.PROC_NOTES);
    } catch (_) {
      return; // Notes column not present yet — skip
    }
    sheet.getRange(2, notesCol, maxRows - 1, 1)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment('top');
  }

  // --------------------------------------------------------------------------
  // Public surface
  // --------------------------------------------------------------------------
  return { applyFormatting };

})();

// Top-level stub for menu item
function Formatting_applyFormatting() { Formatting.applyFormatting(); }
