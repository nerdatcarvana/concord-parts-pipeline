/**
 * HourlyReport.gs
 * Emails the IN-PROGRESS shift report every hour to subscribed people.
 *
 * Origin: night shift lead asked for the "📧 Email Shift Report" snapshot to
 * arrive hourly without clicking. Built as a per-person setting rather than a
 * hardcoded address so anyone else can opt in later with no code change.
 *
 * SUBSCRIPTION STORE: the `hourly_report` column on _Roles.
 *   blank  — not subscribed
 *   DAY | NIGHT | BOTH — mail goes out only while the LIVE shift matches
 *   anything else — skipped + console.warn'd (Setup puts a dropdown on the
 *   column, allowInvalid, so this is rare)
 * _Roles, not _Identity, because _Roles has one row per EMAIL (_Identity has
 * one per alias, and every _Identity row feeds the Processing Requester
 * dropdown). Access only ever grants privileges to role MANAGER, so a
 * non-manager who wants the report gets a row with any other role (PICKER,
 * or a descriptive one like VIEWER) — it carries the subscription and grants
 * nothing. If one email has several rows, it is subscribed to a shift when
 * ANY of its rows matches, and the menu writes every one of its rows.
 *
 * Two ways in: a MANAGER self-subscribes from the menu (writes their own
 * row), or any manager sets the cell on someone's row by hand. Configuration,
 * not data: Turnover's clean slate never touches _Roles.
 *
 * WHAT ONE RUN DOES (time-driven, hourly — installed by Setup):
 *   1. Kill switch: CONFIG.HOURLY_REPORT.ENABLED.
 *   2. Resolve subscribers for the LIVE shift. None → exit, nothing stamped.
 *   3. UNDER THE SCRIPT LOCK: once-per-clock-hour guard, stats snapshot, stamp.
 *      The lock is what keeps this from reading _Requests halfway through a
 *      Turnover clean slate (clearContent → rewrite) and mailing zeros. It is
 *      held for one getDataRange() read — no longer than a pick. If Turnover
 *      holds it past LOCK_TIMEOUT_MS, this hour is skipped; the next one runs.
 *      The guard lives inside the lock so two overlapping firings cannot both
 *      pass it.
 *   4. Skip-empty (CONFIG.HOURLY_REPORT.SKIP_EMPTY, default off).
 *   5. OUTSIDE the lock: build the PDF once, then one email PER recipient —
 *      subscribers don't see each other's addresses, and one bad send doesn't
 *      sink the rest. Failures are logged; ADMIN_EMAIL is alerted at most once
 *      per op day so a quota outage cannot turn into an hourly alert storm.
 *
 * TIMING: Apps Script's everyHours(1) picks a minute offset and holds it (the
 * same behavior that pins Turnover at ~17:54 / ~06:13). Emails land "about once
 * an hour" at that minute, not on the hour. The live shift is judged at fire
 * time, so the first NIGHT email may arrive a few minutes after 17:00 with
 * little in it — that is the honest in-progress number.
 *
 * Public API:
 *   HourlyReport.run()                   → trigger entry point
 *   HourlyReport.subscribeMe()           → menu: set hourly_report on the clicking manager's row(s)
 *   HourlyReport.unsubscribeMe()         → menu: clear hourly_report on the clicking user's row(s)
 *   HourlyReport.listSubscriptions()     → [{ email, shifts, row }] non-blank rows (diagnostics)
 *   HourlyReport.subscribersFor(shift)   → [{ email, shifts }] subscribed + matching, deduped
 */

const HourlyReport = (() => {

  const PROP_LAST_RUN   = 'HOURLY_REPORT_LAST_RUN';    // 'yyyy-MM-ddTHH' of the last send
  const PROP_LAST_ALERT = 'HOURLY_REPORT_LAST_ALERT';  // op-day key of the last admin alert

  // -------------------------------------------------------------------------------
  // run() — hourly trigger entry point
  // -------------------------------------------------------------------------------
  function run() {
    if (!CONFIG.HOURLY_REPORT || CONFIG.HOURLY_REPORT.ENABLED !== true) {
      console.log('HourlyReport.run: disabled via CONFIG.HOURLY_REPORT.ENABLED.');
      return;
    }

    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const tz     = ss.getSpreadsheetTimeZone();
    const now    = new Date();
    const shift  = ShiftReport.shiftOf(now);
    const target = { key: ShiftReport.opDayKey(now), date: now };

    const subs = subscribersFor(shift);
    if (subs.length === 0) {
      console.log(`HourlyReport.run: no subscribers for the ${shift} shift.`);
      return;
    }

    const hourKey = Utilities.formatDate(now, tz, "yyyy-MM-dd'T'HH");
    const props   = PropertiesService.getScriptProperties();
    let stats     = null;

    try {
      Utils.withLock(() => {
        if (props.getProperty(PROP_LAST_RUN) === hourKey) return;   // already sent this hour
        stats = ShiftReport.computeShiftStats(target, shift);
        props.setProperty(PROP_LAST_RUN, hourKey);
      });
    } catch (err) {
      if (String(err.message).includes('could not acquire lock')) {
        console.warn('HourlyReport.run: script lock busy (turnover in progress?) — skipping this hour.');
        return;
      }
      throw err;
    }

    if (!stats) {
      console.log(`HourlyReport.run: already sent for ${hourKey} — duplicate firing ignored.`);
      return;
    }

    const isEmpty = stats.requests === 0 && stats.completed === 0 &&
                    stats.discarded === 0 && stats.openTotal === 0;
    if (CONFIG.HOURLY_REPORT.SKIP_EMPTY === true && isEmpty) {
      console.log(`HourlyReport.run: ${target.key}|${shift} has no activity yet — suppressed (SKIP_EMPTY).`);
      return;
    }

    // One render shared by every recipient. A failure means everyone gets the
    // email without the attachment — never N retries of a failing converter.
    let pdf       = null;
    let attachPdf = CONFIG.HOURLY_REPORT.ATTACH_PDF === true;
    if (attachPdf) {
      try {
        pdf = ShiftReport.buildPdf(target, shift, stats, { partial: true });
      } catch (err) {
        console.error('HourlyReport.run: PDF render failed, sending without it:', err.message);
        attachPdf = false;
      }
    }

    let sent = 0;
    const failures = [];
    subs.forEach(sub => {
      try {
        sent += ShiftReport.sendShiftReport(target, shift, [sub.email], {
          partial    : true,
          stats      : stats,
          pdf        : pdf,
          attachPdf  : attachPdf,
          footerNote : _footerNote(sub),
        });
      } catch (err) {
        failures.push(`${sub.email}: ${err.message}`);
      }
    });

    console.log(`HourlyReport.run: ${target.key}|${shift} @ ${hourKey} — sent to ${sent} of ${subs.length}.`);
    if (failures.length > 0) {
      console.error('HourlyReport.run: send failures:\n' + failures.join('\n'));
      _alertAdminOncePerDay(target.key, failures);
    }
  }

  // -------------------------------------------------------------------------------
  // Subscription reads
  // -------------------------------------------------------------------------------

  // Every _Roles row that has a non-blank hourly_report.
  // → [{ email, shifts, row }]; shifts === null means unrecognized.
  // [] when _Roles or the column is missing (i.e. before Setup has run).
  function listSubscriptions() {
    const cols = _rolesCols();
    if (!cols || cols.sheet.getLastRow() < 2) return [];

    const data = cols.sheet.getDataRange().getValues();
    const out  = [];
    for (let i = 1; i < data.length; i++) {
      const email = String(data[i][cols.iEmail] || '').trim();
      const raw   = String(data[i][cols.iHourly] === undefined ? '' : data[i][cols.iHourly]).trim();
      if (!email || raw === '') continue;
      out.push({ email: email, shifts: _normShifts(raw), row: i + 1 });
    }
    return out;
  }

  function subscribersFor(shift) {
    const E    = CONFIG.ENUMS;
    const seen = new Set();
    const out  = [];
    listSubscriptions().forEach(s => {
      if (s.email.indexOf('@') === -1) {
        console.warn(`HourlyReport: _Roles row ${s.row} email "${s.email}" is not an address — skipped.`);
        return;
      }
      if (s.shifts === null) {
        console.warn(`HourlyReport: _Roles row ${s.row} (${s.email}) has an unrecognized ` +
                     'hourly_report value — use DAY, NIGHT, or BOTH. Skipped.');
        return;
      }
      if (s.shifts !== E.SUB_SHIFT_BOTH && s.shifts !== shift) return;
      const key = s.email.toLowerCase();
      if (seen.has(key)) return;           // several rows for one email → one email sent
      seen.add(key);
      out.push({ email: s.email, shifts: s.shifts });
    });
    return out;
  }

  // -------------------------------------------------------------------------------
  // Menu: subscribeMe() — managers only; prompts for the shift scope
  // -------------------------------------------------------------------------------
  function subscribeMe() {
    const ui   = SpreadsheetApp.getUi();
    const user = Access.currentUserEmail();
    if (!user) {
      ui.alert('🔔 Hourly Reports', 'Could not determine your email address.', ui.ButtonSet.OK);
      return;
    }
    if (!Access.isManager(user)) {
      ui.alert('🔔 Hourly Reports',
               'Only managers can subscribe from the menu. Ask a manager to set hourly_report ' +
               'on your row in _Roles.', ui.ButtonSet.OK);
      return;
    }
    if (!_rolesCols()) {
      ui.alert('🔔 Hourly Reports',
               '_Roles has no hourly_report column yet. Run Parts Management → Re-run Setup first.',
               ui.ButtonSet.OK);
      return;
    }

    const E    = CONFIG.ENUMS;
    const resp = ui.prompt(
      '🔔 Subscribe to Hourly Reports',
      `Which shift should email you the in-progress report every hour?\n\n` +
      `Type ${E.SUB_SHIFT_DAY}, ${E.SUB_SHIFT_NIGHT}, or ${E.SUB_SHIFT_BOTH}.`,
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    const text   = String(resp.getResponseText() || '').trim();
    const shifts = text === '' ? null : _normShifts(text);
    if (shifts === null) {
      ui.alert('🔔 Hourly Reports',
               `"${text}" isn't a shift. Use ${E.SUB_SHIFT_DAY}, ${E.SUB_SHIFT_NIGHT}, ` +
               `or ${E.SUB_SHIFT_BOTH}.`, ui.ButtonSet.OK);
      return;
    }

    let written = 0;
    Utils.withLock(() => { written = _setHourly(user, shifts); });

    ui.alert('🔔 Hourly Reports',
             written > 0
               ? `Subscribed: ${user} will receive the ${_scopeLabel(shifts)} in-progress report ` +
                 'about once an hour.\n\nUse "🔕 Unsubscribe from Hourly Reports" to stop.'
               : `Couldn't find ${user} in _Roles — nothing changed.`,
             ui.ButtonSet.OK);
  }

  // -------------------------------------------------------------------------------
  // Menu: unsubscribeMe() — clears hourly_report on the clicking user's rows
  // -------------------------------------------------------------------------------
  function unsubscribeMe() {
    const ui   = SpreadsheetApp.getUi();
    const user = Access.currentUserEmail();
    if (!user) {
      ui.alert('🔕 Hourly Reports', 'Could not determine your email address.', ui.ButtonSet.OK);
      return;
    }

    let cleared = 0;
    const cols  = _rolesCols();
    if (cols) {
      Utils.withLock(() => {
        const target = user.trim().toLowerCase();
        listSubscriptions().forEach(s => {
          if (s.email.toLowerCase() !== target) return;
          cols.sheet.getRange(s.row, cols.iHourly + 1).setValue('');
          cleared++;
        });
        if (cleared > 0) SpreadsheetApp.flush();
      });
    }

    ui.alert('🔕 Hourly Reports',
             cleared > 0 ? `${user} will no longer receive hourly reports.`
                         : `${user} wasn't subscribed to hourly reports.`,
             ui.ButtonSet.OK);
  }

  // -------------------------------------------------------------------------------
  // Writes (caller holds the lock)
  // -------------------------------------------------------------------------------

  // Set hourly_report on EVERY _Roles row for `email` (so no stale duplicate row
  // can disagree). Never appends: the caller has already checked isManager, so
  // a row exists. Returns rows written.
  function _setHourly(email, shifts) {
    const cols = _rolesCols();
    if (!cols) throw new Error('HourlyReport: _Roles has no hourly_report column. Run Setup.');
    const target = String(email).trim().toLowerCase();
    const data   = cols.sheet.getDataRange().getValues();
    let n = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][cols.iEmail] || '').trim().toLowerCase() !== target) continue;
      cols.sheet.getRange(i + 1, cols.iHourly + 1).setValue(shifts);
      n++;
    }
    if (n > 0) SpreadsheetApp.flush();
    return n;
  }

  // _Roles sheet + 0-based column indexes, or null if the sheet or either
  // column is missing.
  function _rolesCols() {
    const sheet = Sheets.getByName(CONFIG.SHEETS.ROLES);
    if (!sheet || sheet.getLastColumn() < 1) return null;
    const hdr     = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(Utils.cleanHeader);
    const iEmail  = hdr.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RO_EMAIL));
    const iHourly = hdr.indexOf(Utils.cleanHeader(CONFIG.HEADERS.RO_HOURLY));
    if (iEmail === -1 || iHourly === -1) return null;
    return { sheet, iEmail, iHourly };
  }

  // -------------------------------------------------------------------------------
  // Helpers
  // -------------------------------------------------------------------------------

  // DAY/NIGHT/BOTH case-insensitive → canonical; anything else → null.
  // (Blank is filtered out by listSubscriptions before this is reached.)
  function _normShifts(raw) {
    const E = CONFIG.ENUMS;
    const s = String(raw === null || raw === undefined ? '' : raw).trim().toUpperCase();
    if (s === E.SUB_SHIFT_DAY || s === E.SUB_SHIFT_NIGHT || s === E.SUB_SHIFT_BOTH) return s;
    return null;
  }

  function _scopeLabel(shifts) {
    const E = CONFIG.ENUMS;
    if (shifts === E.SUB_SHIFT_DAY)   return 'Day shift';
    if (shifts === E.SUB_SHIFT_NIGHT) return 'Night shift';
    return 'Day and Night shift';
  }

  // Managers can unsubscribe from the menu; non-manager subscribers can't edit
  // _Roles (protected), so the footer names both routes.
  function _footerNote(sub) {
    return `You're receiving this hourly update because ${sub.email} is subscribed to the ` +
           `${_scopeLabel(sub.shifts)} hourly Parts report. To stop: managers can use ` +
           'Parts Management → "🔕 Unsubscribe from Hourly Reports"; anyone else, ask a ' +
           'manager to clear hourly_report on your row in _Roles.';
  }

  // At most one alert per op day: if MailApp quota is exhausted, an hourly
  // alert would fail too, and when it isn't, one heads-up a day is enough.
  function _alertAdminOncePerDay(opDayKey, failures) {
    try {
      const props = PropertiesService.getScriptProperties();
      if (props.getProperty(PROP_LAST_ALERT) === opDayKey) return;
      props.setProperty(PROP_LAST_ALERT, opDayKey);
      MailApp.sendEmail({
        to      : CONFIG.ADMIN_EMAIL,
        subject : `⚠️ Parts Workflow — hourly report send failures (${opDayKey})`,
        body    : 'Hourly in-progress report emails failed for:\n\n' + failures.join('\n') +
                  '\n\nFurther failures today are logged to the Apps Script execution log only.\n' +
                  `Sheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`,
      });
    } catch (err) {
      console.error('HourlyReport: admin alert failed:', err.message);
    }
  }

  // -------------------------------------------------------------------------------
  // Public surface
  // -------------------------------------------------------------------------------
  return { run, subscribeMe, unsubscribeMe, listSubscriptions, subscribersFor };

})();

// Top-level stubs (trigger + menu items must reference top-level names)
function HourlyReport_run()           { HourlyReport.run(); }
function HourlyReport_subscribeMe()   { HourlyReport.subscribeMe(); }
function HourlyReport_unsubscribeMe() { HourlyReport.unsubscribeMe(); }
