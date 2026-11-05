/**
 * Utils.gs
 * Stateless utility belt. No sheet writes, no side effects.
 *
 * Public API:
 *   Utils.getColIndex(sheet, headerConstant)  → 1-based column index (throws if missing)
 *   Utils.getColMap(sheet, ...headerConstants) → { headerConstant: colIndex, … }
 *   Utils.isoNow()                             → ISO 8601 string of current time
 *   Utils.toIso(date)                          → ISO 8601 string from a Date object
 *   Utils.newRequestId()                       → full UUID string
 *   Utils.newEventId()                         → full UUID string
 *   Utils.withLock(fn)                         → runs fn() under ScriptLock; returns result
 *   Utils.toast(msg, title, duration)          → SpreadsheetApp.toast wrapper
 *   Utils.normalizeBuildLevel(raw)             → canonical build level or null
 *   Utils.normalizeLotStatus(raw)              → canonical lot status or raw trimmed string
 *   Utils.cleanHeader(str)                     → trimmed, invisible-char-stripped string
 */

const Utils = (() => {

  // --------------------------------------------------------------------------
  // Internal: strip zero-width and bidi control characters, then trim.
  // This is the fix for the stray invisible char in "Build Load Level".
  // --------------------------------------------------------------------------
  function cleanHeader(str) {
    if (str === null || str === undefined) return '';
    // Remove zero-width space, ZWSP, BOM, bidi marks, soft hyphen, etc.
    return String(str)
      .replace(/[\u200B-\u200D\uFEFF\u00AD\u061C\u200E\u200F\u202A-\u202E\u2060-\u2064]/g, '')
      .trim();
  }

  // --------------------------------------------------------------------------
  // Header → 1-based column index.
  // Reads the first row of `sheet` once and caches nothing (idempotent reads
  // are cheap; caching across executions in Apps Script is error-prone).
  // Throws a descriptive error if the header is not found so bugs surface fast.
  // --------------------------------------------------------------------------
  function getColIndex(sheet, headerValue) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const cleaned = headers.map(cleanHeader);
    const idx = cleaned.indexOf(cleanHeader(headerValue));
    if (idx === -1) {
      throw new Error(
        `Utils.getColIndex: header "${headerValue}" not found in sheet "${sheet.getName()}". ` +
        `Found: [${cleaned.join(' | ')}]`
      );
    }
    return idx + 1; // 1-based
  }

  // --------------------------------------------------------------------------
  // Convenience: resolve multiple headers in one call.
  // Returns { headerValue: colIndex, … } for each argument passed.
  // --------------------------------------------------------------------------
  function getColMap(sheet, ...headerValues) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const cleaned = headers.map(cleanHeader);
    const map = {};
    for (const hv of headerValues) {
      const idx = cleaned.indexOf(cleanHeader(hv));
      if (idx === -1) {
        throw new Error(
          `Utils.getColMap: header "${hv}" not found in sheet "${sheet.getName()}". ` +
          `Found: [${cleaned.join(' | ')}]`
        );
      }
      map[hv] = idx + 1;
    }
    return map;
  }

  // --------------------------------------------------------------------------
  // ISO 8601 timestamp helpers.
  // Apps Script's Utilities.formatDate requires a timezone string; using the
  // spreadsheet's own timezone keeps times consistent with what operators see.
  // --------------------------------------------------------------------------
  function isoNow() {
    return toIso(new Date());
  }

  function toIso(date) {
    const tz  = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const d   = date instanceof Date ? date : new Date(date);
    // Build the ISO string manually — Utilities.formatDate can drop the 'T'
    // literal in certain locale configurations, producing a space instead.
    const datePart = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    const timePart = Utilities.formatDate(d, tz, 'HH:mm:ss');
    return datePart + 'T' + timePart;
  }

  // --------------------------------------------------------------------------
  // Collision-resistant surrogate keys — full UUID (not truncated).
  // --------------------------------------------------------------------------
  function newRequestId() {
    return Utilities.getUuid();
  }

  function newEventId() {
    return Utilities.getUuid();
  }

  // --------------------------------------------------------------------------
  // LockService wrapper.
  // Runs `fn` under a ScriptLock. Releases the lock in a finally block.
  // Throws if the lock cannot be acquired within CONFIG.LOCK_TIMEOUT_MS.
  // --------------------------------------------------------------------------
  function withLock(fn) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      throw new Error('Utils.withLock: could not acquire lock within timeout. System busy.');
    }
    try {
      return fn();
    } finally {
      lock.releaseLock();
    }
  }

  // --------------------------------------------------------------------------
  // Toast wrapper — safe to call from time-driven triggers (getUi() throws
  // in that context; SpreadsheetApp.toast does not).
  // --------------------------------------------------------------------------
  function toast(message, title, durationSeconds) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title || '', durationSeconds || 4);
  }

  // --------------------------------------------------------------------------
  // Build Load Level normalization.
  // Returns the canonical value from CONFIG.BUILD_LEVEL_MAP, or the uppercased
  // input if it already matches a known level, or null if unrecognizable.
  // --------------------------------------------------------------------------
  function normalizeBuildLevel(raw) {
    if (!raw) return null;
    const upper = String(raw).trim().toUpperCase();
    if (CONFIG.BUILD_LEVEL_MAP[upper]) return CONFIG.BUILD_LEVEL_MAP[upper];
    // Direct match against canonical list (case-insensitive)
    const match = CONFIG.BUILD_LEVELS.find(lvl => lvl.toUpperCase() === upper);
    return match || null;
  }

  // --------------------------------------------------------------------------
  // Lot Status normalization.
  // Form values arrive clean (multiple-choice), but walk-ins may vary.
  // Returns the canonical value if found, otherwise returns the trimmed input.
  // --------------------------------------------------------------------------

function normalizeLotStatus(raw) {
    if (!raw) return '';
    const trimmed = String(raw).trim();
    // Case-insensitive match against known statuses + Walk In
    const all = [...CONFIG.LOT_STATUSES, CONFIG.LOT_WALKIN];
    const match = all.find(s => s.toLowerCase() === trimmed.toLowerCase());
    return match || trimmed; // pass through unknowns rather than dropping them
  }

  // --------------------------------------------------------------------------
  // Public surface
  // --------------------------------------------------------------------------
  return {
    cleanHeader,
    getColIndex,
    getColMap,
    isoNow,
    toIso,
    newRequestId,
    newEventId,
    withLock,
    toast,
    normalizeBuildLevel,
    normalizeLotStatus,
  };

})();
