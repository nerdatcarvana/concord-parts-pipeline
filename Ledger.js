/**
 * Ledger.gs
 * The ONLY module that writes to _EventLog and _Requests.
 * Both sheets are append-only and NEVER cleared.
 *
 * Public API:
 *   Ledger.appendEvent(eventObj)                    → appends one row to _EventLog
 *   Ledger.upsertRequest(requestObj)                → insert or update a row in _Requests
 *   Ledger.getRequest(requestId)                    → returns the _Requests row object or null
 *   Ledger.findOpenRequest(stockNumber, email)      → returns request_id of matching OPEN row or null
 *   Ledger.commitPick(reqId, events, requestObj)    → batched pick write: N appendEvents + upsertRequest
 *                                                     sharing one sheet-ref + one header read.
 *                                                     Use instead of individual calls
 *                                                     from the pick hot path.
 *   Ledger.clearCache()                             → force-expire the execution-scoped sheet cache
 *                                                     (tests / Health / EOD only; pick path never needs it)
 *
 * ROW LOOKUP STRATEGY (the long-term latency fix):
 *   _Requests is append-only and never cleared, so it grows without bound. The
 *   old implementation read the ENTIRE sheet (getDataRange().getValues()) on
 *   every upsert/get/probe just to locate one row — O(n) per pick, with n
 *   growing forever. All lookups now use createTextFinder() scoped to a single
 *   column: the search runs server-side and returns just the matching cell, so
 *   a pick costs the same on day 1 and day 1,000. Only the one matched row is
 *   then read/written. request_id is a full UUID stored as plain text, so
 *   matchEntireCell(true) is exact and collision-free.
 */

const Ledger = (() => {

  // --------------------------------------------------------------------------
  // Execution-scoped sheet + header cache
  //
  // GAS executions are single-threaded and short-lived. Caching sheet references
  // and header arrays at module scope eliminates redundant Sheets.get() calls
  // (each of which does a PropertiesService read + ss.getSheets() scan) and
  // redundant getRange(1,1,1,lastCol).getValues() header reads that fire once
  // per appendEvent() call. Combined with TextFinder lookups, a bin pick with
  // N bins costs:
  //   1 × Sheets.get(_EventLog) + 1 × header read + 1 × Sheets.get(_Requests)
  //   + 1 × header read + 1 × TextFinder + 1 × single-row read + writes
  // — regardless of bin count AND regardless of how large _Requests has grown.
  //
  // Safety: cache entries are keyed by sheet name and hold the Sheet object plus
  // the header array. Sheet objects are stable JS references within one execution;
  // they do not go stale. Header arrays are read once from the live sheet; they
  // are correct as long as no code adds/reorders columns mid-execution (Setup
  // is the only thing that does, and it never runs concurrently with a pick).
  // clearCache() is exposed for the rare caller that needs a guaranteed fresh read.
  // --------------------------------------------------------------------------
  const _cache = {};   // { [sheetName]: { sheet, headers } }

  function _getCached(sheetName) {
    if (_cache[sheetName]) return _cache[sheetName];
    const sheet = Sheets.getByName(sheetName);
    if (!sheet) throw new Error(`Ledger: sheet "${sheetName}" not found. Run Setup.initialize() first.`);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(Utils.cleanHeader);
    _cache[sheetName] = { sheet, headers };
    return _cache[sheetName];
  }

  function clearCache() {
    Object.keys(_cache).forEach(k => delete _cache[k]);
  }

  // --------------------------------------------------------------------------
  // Internal: locate the row number of a request_id via TextFinder.
  // Scoped to the id column only; matchEntireCell guarantees exactness.
  // Returns the 1-based sheet row number, or -1 if absent.
  // --------------------------------------------------------------------------
  function _findRowById(sheet, idCol, requestId) {
    const id = String(requestId || '').trim();
    if (!id) return -1;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return -1;
    const hit = sheet.getRange(2, idCol, lastRow - 1, 1)
      .createTextFinder(id)
      .matchEntireCell(true)
      .findNext();
    return hit ? hit.getRow() : -1;
  }

  // --------------------------------------------------------------------------
  // appendEvent(eventObj)
  // --------------------------------------------------------------------------
  function appendEvent(eventObj) {
    const { sheet, headers } = _getCached(CONFIG.SHEETS.EVENT_LOG);
    // Project onto the live header order (name-mapped), mirroring _Requests. A
    // fixed positional array silently misaligns if _EventLog is ever reordered or
    // re-created with columns in a different sequence; _ensureHeaders only appends
    // missing columns, it never reorders, so position is not a safe contract.
    sheet.appendRow(_buildEventRow(eventObj, headers));
  }

  // --------------------------------------------------------------------------
  // upsertRequest(requestObj)
  // --------------------------------------------------------------------------
  function upsertRequest(requestObj) {
    const { sheet, headers } = _getCached(CONFIG.SHEETS.REQUESTS);
    _upsertInto(sheet, headers, requestObj);
  }

  // --------------------------------------------------------------------------
  // getRequest(requestId)
  // --------------------------------------------------------------------------
  function getRequest(requestId) {
    const { sheet, headers } = _getCached(CONFIG.SHEETS.REQUESTS);
    const idColIdx = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RQ_REQUEST_ID));
    if (idColIdx === -1) return null;

    const rowNum = _findRowById(sheet, idColIdx + 1, requestId);
    if (rowNum === -1) return null;

    const row = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
    return _rowToObj(row, headers);
  }

  // --------------------------------------------------------------------------
  // findOpenRequest(stockNumber, requesterEmail)
  // Durable idempotency probe for the walk-in path. Returns the request_id of an
  // existing OPEN _Requests row matching BOTH stock number and requester email,
  // or null. _Requests is the source of truth (board is a derived view), so this
  // survives the GAS cross-execution stale-read window that defeats an in-cell
  // claim flag. Matching on (stock, requester, OPEN) — not stock alone — means a
  // genuine re-request of the same stock by a DIFFERENT requester is never
  // collapsed, and a re-request by the SAME requester after the prior one was
  // PICKED/COMPLETED/DISCARDED (no longer OPEN) is correctly treated as new.
  //
  // TextFinder narrows to rows whose stock cell DISPLAYS exactly `stock`; each
  // candidate is then re-verified against the underlying values (walked
  // bottom-up so the newest matching OPEN row wins), so a display-vs-value
  // formatting difference can only cost a missed probe — never a false match.
  // A missed probe degrades to the pre-probe behavior (a new request_id), which
  // the SYNC_PROCESSED fast path already bounds to the rare concurrency window.
  // --------------------------------------------------------------------------
  function findOpenRequest(stockNumber, requesterEmail) {
    const stock = String(stockNumber || '').trim();
    const email = String(requesterEmail || '').trim().toLowerCase();
    if (!stock || !email) return null;

    const { sheet, headers } = _getCached(CONFIG.SHEETS.REQUESTS);
    const idColIdx     = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RQ_REQUEST_ID));
    const stockColIdx  = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RQ_STOCK));
    const emailColIdx  = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RQ_REQ_EMAIL));
    const statusColIdx = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RQ_STATUS));
    if (idColIdx === -1 || stockColIdx === -1 || emailColIdx === -1 || statusColIdx === -1) {
      return null;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    const matches = sheet.getRange(2, stockColIdx + 1, lastRow - 1, 1)
      .createTextFinder(stock)
      .matchEntireCell(true)
      .findAll();

    for (let i = matches.length - 1; i >= 0; i--) {
      const rowNum = matches[i].getRow();
      const row    = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
      if (String(row[statusColIdx]).trim() !== CONFIG.ENUMS.STATUS_OPEN)       continue;
      if (String(row[stockColIdx]).trim()  !== stock)                          continue;
      if (String(row[emailColIdx]).trim().toLowerCase() !== email)             continue;
      return String(row[idColIdx]).trim();
    }
    return null;
  }

  // --------------------------------------------------------------------------
  // commitPick(reqId, events, requestObj)
  //
  // Batched write for the pick hot path. Replaces the pattern of:
  //   for (bin of bins) Ledger.appendEvent(PICKED)
  //   Ledger.appendEvent(COMPLETED)
  //   Ledger.upsertRequest(...)
  //
  // Benefits vs. individual calls:
  //   • _EventLog sheet ref + header array: resolved once (cached), not per-event.
  //   • _Requests row located via TextFinder (single-column server-side search)
  //     and only that one row is read — no full-table getDataRange(), regardless
  //     of how many requests the ledger has accumulated.
  //   • appendRow calls are still individual — GAS does not support true batch
  //     appends — but the per-call overhead drops to just the network write,
  //     not a sheet resolve + header read + write.
  //
  // `events`     — array of event objects (PICKED×N + COMPLETED), in order.
  // `requestObj` — the upsertRequest payload.
  //
  // Returns the status string of the request row before the write, or null if
  // no existing row was found. Callers use this to detect the already-completed
  // case without a separate getRequest() call.
  // --------------------------------------------------------------------------
  function commitPick(reqId, events, requestObj) {
    // --- Locate + read the _Requests row FIRST, before writing anything ---
    // The already-COMPLETED guard has to run before the events are appended,
    // not after. Appending first and reporting the prior status afterwards
    // means a second board row for the same request_id writes a full duplicate
    // PICKED×N + COMPLETED set into _EventLog before the caller can bail —
    // silently inflating pick counts, bin counts and cycle-time percentiles in
    // every shift report. The single-read optimisation is preserved: this row
    // is read once here and reused for the merge below.
    const { sheet: rqSheet, headers: rqHeaders } = _getCached(CONFIG.SHEETS.REQUESTS);

    const idColIdx = rqHeaders.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RQ_REQUEST_ID));
    if (idColIdx === -1) {
      throw new Error('Ledger.commitPick: request_id column not found in _Requests.');
    }
    const statusColIdx = rqHeaders.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RQ_STATUS));

    const rowNum      = _findRowById(rqSheet, idColIdx + 1, reqId);
    let   existingRow = null;
    let   priorStatus = null;
    if (rowNum !== -1) {
      existingRow = rqSheet.getRange(rowNum, 1, 1, rqHeaders.length).getValues()[0];
      priorStatus = statusColIdx !== -1 ? String(existingRow[statusColIdx]).trim() : null;
    }

    // Terminal already — write nothing at all and let the caller clean up.
    if (priorStatus === CONFIG.ENUMS.STATUS_COMPLETED) return priorStatus;

    // --- _EventLog: resolve once, append all events ---
    const { sheet: elSheet, headers: elHeaders } = _getCached(CONFIG.SHEETS.EVENT_LOG);
    for (const ev of events) {
      elSheet.appendRow(_buildEventRow(ev, elHeaders));
    }

    // --- _Requests: write using the row already located above ---
    if (rowNum === -1) {
      rqSheet.appendRow(_buildRequestRow(requestObj, rqHeaders));
    } else {
      const updated = _mergeRequestRow(existingRow, rqHeaders, requestObj);
      rqSheet.getRange(rowNum, 1, 1, updated.length).setValues([updated]);
    }

    return priorStatus;
  }

  // --------------------------------------------------------------------------
  // statusOf(requestId) → current status string, or null if no row exists.
  // Cheap single-row probe for callers that must not write when a request has
  // already reached a terminal state (see _handleDiscard).
  // --------------------------------------------------------------------------
  function statusOf(requestId) {
    const { sheet, headers } = _getCached(CONFIG.SHEETS.REQUESTS);
    const idColIdx     = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RQ_REQUEST_ID));
    const statusColIdx = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RQ_STATUS));
    if (idColIdx === -1 || statusColIdx === -1) return null;

    const rowNum = _findRowById(sheet, idColIdx + 1, requestId);
    if (rowNum === -1) return null;

    return String(sheet.getRange(rowNum, statusColIdx + 1).getValue()).trim();
  }

  // --------------------------------------------------------------------------
  // Internal: upsert requestObj into _Requests.
  // Locates the target row via TextFinder on the request_id column, reads ONLY
  // that row, merges, writes it back. Appends when absent.
  // Returns the existing status string (before write), or null if inserting new.
  // Shared by upsertRequest() and commitPick() to avoid duplicating the lookup.
  // --------------------------------------------------------------------------
  function _upsertInto(sheet, headers, requestObj) {
    const idColIdx = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RQ_REQUEST_ID));
    if (idColIdx === -1) {
      throw new Error('Ledger.upsertRequest: request_id column not found in _Requests.');
    }

    const statusColIdx = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RQ_STATUS));

    const rowNum = _findRowById(sheet, idColIdx + 1, requestObj.request_id);

    if (rowNum === -1) {
      sheet.appendRow(_buildRequestRow(requestObj, headers));
      return null;
    }

    const existingRow    = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
    const existingStatus = statusColIdx !== -1 ? String(existingRow[statusColIdx]).trim() : null;
    const updated        = _mergeRequestRow(existingRow, headers, requestObj);
    sheet.getRange(rowNum, 1, 1, updated.length).setValues([updated]);

    return existingStatus;
  }

  // --------------------------------------------------------------------------
  // Internal: event field → cleaned-header map (defaults applied here)
  // --------------------------------------------------------------------------
  function _eventObjToMap(e) {
    const H = CONFIG.HEADERS;
    const m = {};
    m[Utils.cleanHeader(H.EL_EVENT_ID)]       = e.event_id             || Utils.newEventId();
    m[Utils.cleanHeader(H.EL_REQUEST_ID)]     = e.request_id           || '';
    m[Utils.cleanHeader(H.EL_EVENT_TYPE)]     = e.event_type           || '';
    m[Utils.cleanHeader(H.EL_EVENT_TS)]       = e.event_ts             || Utils.isoNow();
    m[Utils.cleanHeader(H.EL_ACTOR_EMAIL)]    = e.actor_email          || '';
    m[Utils.cleanHeader(H.EL_ACTOR_DISPLAY)]  = e.actor_display        || '';
    m[Utils.cleanHeader(H.EL_INTAKE_SRC)]     = e.intake_source        || '';
    m[Utils.cleanHeader(H.EL_STOCK)]          = e.stock_number         || '';
    m[Utils.cleanHeader(H.EL_BUILD)]          = e.build_load_level     || '';
    m[Utils.cleanHeader(H.EL_LOT)]            = e.lot_status           || '';
    m[Utils.cleanHeader(H.EL_BIN_CODE)]       = e.bin_code             || '';
    m[Utils.cleanHeader(H.EL_LOC_TYPE)]       = e.location_type        || '';
    m[Utils.cleanHeader(H.EL_REQ_DEPT)]       = e.requester_department || '';
    m[Utils.cleanHeader(H.EL_RAW_INPUT)]      = e.raw_input            || '';
    m[Utils.cleanHeader(H.EL_NOTES)]          = e.notes                || '';
    return m;
  }

  // --------------------------------------------------------------------------
  // Internal: build _EventLog row by projecting the field map onto live headers.
  // Columns present in the sheet but absent from the map resolve to '' (safe);
  // map fields whose header isn't in the sheet are dropped (would be appended by
  // Setup._ensureHeaders). Order follows the sheet, not this function.
  // --------------------------------------------------------------------------
  function _buildEventRow(e, headers) {
    const map = _eventObjToMap(e);
    return headers.map(h => {
      const val = map[h];
      return val !== undefined ? val : '';
    });
  }

  // --------------------------------------------------------------------------
  // Internal: build full ordered values array for a new _Requests row
  // --------------------------------------------------------------------------
  function _buildRequestRow(r, headers) {
    const map = _requestObjToMap(r);
    return headers.map(h => {
      const val = map[h];
      return val !== undefined ? val : '';
    });
  }

  // --------------------------------------------------------------------------
  // Internal: merge incoming fields onto existing row; immutable fields preserved
  // --------------------------------------------------------------------------
  function _mergeRequestRow(existingRow, headers, updates) {
    const IMMUTABLE = new Set([
      Utils.cleanHeader(CONFIG.HEADERS.RQ_REQUEST_ID),
      Utils.cleanHeader(CONFIG.HEADERS.RQ_INTAKE_SRC),
      Utils.cleanHeader(CONFIG.HEADERS.RQ_REQUESTED_TS),
      Utils.cleanHeader(CONFIG.HEADERS.RQ_REQ_EMAIL),
      Utils.cleanHeader(CONFIG.HEADERS.RQ_REQ_DISPLAY),
      Utils.cleanHeader(CONFIG.HEADERS.RQ_REQ_DEPT),
      Utils.cleanHeader(CONFIG.HEADERS.RQ_STOCK),
      Utils.cleanHeader(CONFIG.HEADERS.RQ_BUILD),
      Utils.cleanHeader(CONFIG.HEADERS.RQ_LOT),
    ]);

    const updateMap = _requestObjToMap(updates);
    return headers.map((h, i) => {
      if (IMMUTABLE.has(h)) return existingRow[i];
      const incoming = updateMap[h];
      return incoming !== undefined ? incoming : existingRow[i];
    });
  }

  // --------------------------------------------------------------------------
  // Internal: requestObj → map keyed by clean header strings
  // --------------------------------------------------------------------------
  function _requestObjToMap(r) {
    const H = CONFIG.HEADERS;
    return {
      [Utils.cleanHeader(H.RQ_REQUEST_ID)]   : r.request_id           || '',
      [Utils.cleanHeader(H.RQ_INTAKE_SRC)]   : r.intake_source         || '',
      [Utils.cleanHeader(H.RQ_REQUESTED_TS)] : r.requested_ts          || '',
      [Utils.cleanHeader(H.RQ_REQ_EMAIL)]    : r.requester_email       || '',
      [Utils.cleanHeader(H.RQ_REQ_DISPLAY)]  : r.requester_display     || '',
      [Utils.cleanHeader(H.RQ_REQ_DEPT)]     : r.requester_department  || '',
      [Utils.cleanHeader(H.RQ_STOCK)]        : r.stock_number          || '',
      [Utils.cleanHeader(H.RQ_BUILD)]        : r.build_load_level      || '',
      [Utils.cleanHeader(H.RQ_LOT)]          : r.lot_status            || '',
      [Utils.cleanHeader(H.RQ_BIN_COUNT)]    : r.bin_count             !== undefined ? r.bin_count : '',
      [Utils.cleanHeader(H.RQ_BINS)]         : r.bins                  || '',
      [Utils.cleanHeader(H.RQ_FIRST_PICK_TS)]: r.first_pick_ts         || '',
      [Utils.cleanHeader(H.RQ_COMPLETE_TS)]  : r.complete_ts           || '',
      [Utils.cleanHeader(H.RQ_PICKER_EMAIL)] : r.picker_email          || '',
      [Utils.cleanHeader(H.RQ_PICKER_DISP)]  : r.picker_display        || '',
      [Utils.cleanHeader(H.RQ_CYCLE_MIN)]    : r.cycle_time_min        !== undefined ? r.cycle_time_min : '',
      [Utils.cleanHeader(H.RQ_STATUS)]       : r.status                || CONFIG.ENUMS.STATUS_OPEN,
      [Utils.cleanHeader(H.RQ_NOTES)]        : r.notes                 || '',
    };
  }

  // --------------------------------------------------------------------------
  // Internal: row array → plain object keyed by header strings
  // --------------------------------------------------------------------------
  function _rowToObj(row, headers) {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  }

  // --------------------------------------------------------------------------
  // Public surface
  // --------------------------------------------------------------------------
  return {
    appendEvent,
    upsertRequest,
    getRequest,
    findOpenRequest,
    commitPick,
    statusOf,
    clearCache,
  };

})();
