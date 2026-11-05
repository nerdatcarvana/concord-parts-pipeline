/**
 * ShiftReport.gs   (rev 2.1 — carryover is shift-windowed)
 * Shift-scoped, at-a-glance completed-work report delivered as an HTML email.
 *
 * REV 2.1 — OPEN/CARRYOVER FALSE POSITIVES FIXED. The rev-2 tripwire counted
 * EVERY OPEN row in the ledger, with no timestamp filter. The turnover trigger
 * fires in a window AFTER the boundary (17:00 run lands 17:0x–18:00), so
 * requests submitted after the boundary — the incoming shift's live queue —
 * were already OPEN at stats time and got flagged as the ENDED shift's
 * carryover. openTotal now counts only OPEN rows whose requested_ts falls in
 * or before the reported shift's window; later-stamped OPEN rows are judged
 * at their own boundary. OPEN rows with an unparseable requested_ts still
 * count (the tripwire must never under-report) and are console.warn'd.
 *
 * Shift model (constants live in CONFIG.SHIFT — the single source of truth
 * shared with Turnover): the operating day starts at START_HR (06:00).
 *   DAY   shift = [START_HR, SECOND_HR)          → 06:00 – 17:00
 *   NIGHT shift = [SECOND_HR, START_HR next day) → 17:00 – 06:00; wraps past
 *                 midnight and collapses into the op day of the DAY shift it
 *                 followed.
 *
 * Department process rule: NO OPEN WORK CROSSES A SHIFT BOUNDARY. Each shift
 * finishes everything it starts, so "Open / Carryover" should always be 0.
 * A nonzero count is rendered as a red ⚠ anomaly banner, not a neutral stat.
 *
 * Every report covers exactly ONE shift of ONE op day. Delivery paths:
 *   1. Shift turnover (Turnover.run, triggered ~17:00 and ~06:00): the report
 *      for the shift that just ENDED, mailed to all MANAGERs (fallback:
 *      ADMIN_EMAIL). Turnover pre-computes the stats BEFORE it purges
 *      _Requests and passes the snapshot in via opts.stats — never recompute
 *      after the purge.
 *   2. On demand (menu → "📧 Email Shift Report"): in-progress snapshot of the
 *      CURRENT shift, sent to whoever clicked. Read-only — never purges.
 *
 * Metrics (deliberately lean — at-a-glance completed work):
 *   Total Requests, Total Completions, Total Reworks, Avg Turnaround, plus a
 *   per-Build-Load-Level breakdown ("Rework" is a canonical build level — its
 *   row IS the rework count) and a small footer (discarded / bins picked /
 *   median / p90 / open carryover).
 *
 * Public API:
 *   ShiftReport.SHIFTS                            → { DAY:'DAY', NIGHT:'NIGHT' }
 *   ShiftReport.opDayKey(date)                    → 'YYYY-MM-DD' op-day key
 *   ShiftReport.shiftOf(date)                     → 'DAY' | 'NIGHT'
 *   ShiftReport.computeShiftStats(target, shift)  → stats object (one ledger read)
 *   ShiftReport.sendShiftReport(target, shift, recipients, opts) → count mailed
 *       target     : { key:'YYYY-MM-DD', date:Date }
 *       recipients : string[] of email addresses
 *       opts       : { partial:boolean, stats:object }
 *                    partial=true labels the report "in progress";
 *                    stats = precomputed snapshot (skips the ledger read)
 *   ShiftReport.emailMe()                         → menu entry point
 *   ShiftReport.defaultRecipients()               → MANAGERs (fallback ADMIN_EMAIL)
 */

const ShiftReport = (() => {

  const SHIFTS = { DAY: 'DAY', NIGHT: 'NIGHT' };

  // -------------------------------------------------------------------------------
  // Time helpers (constants from CONFIG.SHIFT — never hardcode hours here)
  // -------------------------------------------------------------------------------
  function _tz() {
    return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  }

  // Op-day key for an instant: shift back to the boundary, then take the date.
  function opDayKey(date) {
    const shifted = new Date(date.getTime() - CONFIG.SHIFT.START_HR * 3600 * 1000);
    return Utilities.formatDate(shifted, _tz(), 'yyyy-MM-dd');
  }

  // DAY for [START_HR, SECOND_HR), otherwise NIGHT (wraps past midnight).
  function shiftOf(date) {
    const hr = parseInt(Utilities.formatDate(date, _tz(), 'H'), 10);
    return (hr >= CONFIG.SHIFT.START_HR && hr < CONFIG.SHIFT.SECOND_HR) ? SHIFTS.DAY : SHIFTS.NIGHT;
  }

  function _shiftName(shift)  { return shift === SHIFTS.DAY ? 'Day' : 'Night'; }
  function _shiftWindow(shift) {
    const s = CONFIG.SHIFT;
    return shift === SHIFTS.DAY
      ? `${s.START_HR}:00 – ${s.SECOND_HR}:00`
      : `${s.SECOND_HR}:00 – ${s.START_HR}:00 (+1)`;
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

  // An OPEN row is CARRYOVER for the reported shift only if it was REQUESTED
  // during or before that shift's window. The turnover trigger fires minutes
  // (sometimes longer) AFTER the boundary, so requests submitted after it —
  // the incoming shift's live queue — are already OPEN in the ledger when the
  // ended shift's stats are computed, but they are NOT that shift's carryover.
  // Op-day keys are 'yyyy-MM-dd', so lexicographic compare is chronological;
  // within one op day, DAY precedes NIGHT.
  function _requestedInOrBefore(ts, target, shift) {
    const k = opDayKey(ts);
    if (k !== target.key) return k < target.key;
    return shift === SHIFTS.NIGHT || shiftOf(ts) === shift;
  }

  // Normalize a raw build-load-level to its canonical form via CONFIG.
  function _normBuild(raw) {
    const s = String(raw || '').trim();
    if (!s) return '';
    return CONFIG.BUILD_LEVEL_MAP[s.toUpperCase()] || s;
  }

  // -------------------------------------------------------------------------------
  // Recipients — all MANAGERs from _Roles; fallback ADMIN_EMAIL so a report is
  // never silently dropped. (Shared with Turnover.)
  // -------------------------------------------------------------------------------
  function defaultRecipients() {
    let managers = [];
    try { managers = Access.getManagerEmails(); } catch (_) { /* fall through */ }
    return managers.length > 0 ? managers : [CONFIG.ADMIN_EMAIL];
  }

  // -------------------------------------------------------------------------------
  // Entry points
  // -------------------------------------------------------------------------------
  function sendShiftReport(target, shift, recipients, opts) {
    opts = opts || {};
    const to = (recipients || []).map(e => String(e).trim()).filter(e => e !== '');
    if (to.length === 0) {
      console.error('ShiftReport.sendShiftReport: no recipients — nothing sent.');
      return 0;
    }

    // Turnover passes a pre-purge snapshot in opts.stats; everyone else computes live.
    const stats   = opts.stats || computeShiftStats(target, shift);
    const subject = `Parts Dept — ${_shiftName(shift)} Shift Report ${target.key}` +
                    (opts.partial ? ' (in progress)' : '') +
                    (!opts.partial && stats.openTotal > 0 ? ' ⚠ OPEN CARRYOVER' : '');

    // --- Print-ready PDF attachment -----------------------------------------
    // Decorative: the HTML body remains the primary artifact and carries every
    // figure. Generation is therefore best-effort — a converter hiccup must not
    // stop a turnover report from going out, so failures are logged and the
    // email sends without the attachment.
    let attachments = [];
    try {
      const tz = _tz();

      // Label the OP DAY, not the wall clock. emailMe() passes target.date =
      // now, and on a night shift past midnight the op day is still the
      // previous calendar date — using target.date printed "July 31" under a
      // header reading "Op Day 2026-07-30". Rebuild the date from the key and
      // anchor it at noon so no timezone or DST edge can shift it a day.
      const kp = String(target.key || '').split('-').map(Number);
      const opDate = (kp.length === 3 && kp.every(n => !isNaN(n)))
        ? new Date(kp[0], kp[1] - 1, kp[2], 12, 0, 0)
        : target.date;

      attachments = [ReportPdf.build(target, shift, stats, {
        partial        : !!opts.partial,
        shiftName      : _shiftName(shift),
        shiftWindow    : _shiftWindow(shift),
        dayLabel       : Utilities.formatDate(opDate, tz, 'EEEE, MMMM d yyyy'),
        generatedLabel : Utilities.formatDate(new Date(), tz, 'EEE MMM d, h:mm a'),
      })];
    } catch (err) {
      console.error('ShiftReport: PDF attachment failed, sending without it:', err.message);
    }

    MailApp.sendEmail({
      to          : to.join(','),
      subject     : subject,
      body        : _renderText(target, shift, stats, opts),   // plain-text fallback
      htmlBody    : _renderHtml(target, shift, stats,
                                Object.assign({}, opts, { hasPdf: attachments.length > 0 })),
      name        : 'Parts Request Workflow',
      attachments : attachments,
    });
    return to.length;
  }

  // Menu path: in-progress snapshot of the CURRENT shift, to whoever clicked.
  function emailMe() {
    const user = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    if (!user) {
      Utils.toast('Could not determine your email address.', '📧 Shift Report', 5);
      return;
    }
    const now    = new Date();
    const target = { key: opDayKey(now), date: now };
    const shift  = shiftOf(now);
    sendShiftReport(target, shift, [user], { partial: true });
    Utils.toast(`${_shiftName(shift)} shift report for ${target.key} sent to ${user}.`, '📧 Shift Report', 5);
  }

  // -------------------------------------------------------------------------------
  // computeShiftStats(target, shift)
  // ONE getDataRange() read of _Requests; everything derived in memory.
  // "Requests" bucket on requested_ts; "Completed/Discarded" bucket on
  // complete_ts. A row only counts when its timestamp lands in BOTH the target
  // op day AND the target shift — this is what makes the report shift-aware.
  // -------------------------------------------------------------------------------
  function computeShiftStats(target, shift) {
    const H = CONFIG.HEADERS;
    const E = CONFIG.ENUMS;

    const stats = {
      requests: 0, completed: 0, discarded: 0, reworks: 0,
      openTotal: 0,           // OPEN rows REQUESTED in/before this shift's window —
                              // true carryover; should ALWAYS be 0 at turnover.
                              // OPEN rows requested AFTER the window (the next
                              // shift's live queue, visible because the trigger
                              // fires after the boundary) are excluded.
      binsPicked: 0,
      cycles: [],             // cycle_time_min (turnaround) of this shift's completions
      byBuild: {},            // canonical level → { requests, completed, cycles: [] }
    };
    CONFIG.BUILD_LEVELS.forEach(l => { stats.byBuild[l] = { requests: 0, completed: 0, cycles: [] }; });

    const sheet = Sheets.getByName(CONFIG.SHEETS.REQUESTS, SpreadsheetApp.getActiveSpreadsheet());
    if (!sheet || sheet.getLastRow() < 2) return stats;

    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(Utils.cleanHeader);
    const col = (h) => headers.indexOf(Utils.cleanHeader(h));

    const iStatus = col(H.RQ_STATUS);
    const iReqTs  = col(H.RQ_REQUESTED_TS);
    const iCompTs = col(H.RQ_COMPLETE_TS);
    const iCycle  = col(H.RQ_CYCLE_MIN);
    const iBuild  = col(H.RQ_BUILD);
    const iBinCnt = col(H.RQ_BIN_COUNT);
    if (iStatus === -1) return stats;

    const buildBucket = (level) => {
      if (!level) return null;
      if (!stats.byBuild[level]) stats.byBuild[level] = { requests: 0, completed: 0, cycles: [] };
      return stats.byBuild[level];
    };

    let openNoTs = 0;   // OPEN rows with no parseable requested_ts (counted, warned)

    for (let i = 1; i < data.length; i++) {
      const row    = data[i];
      const status = String(row[iStatus] || '').trim();
      if (!status) continue;

      const level = iBuild !== -1 ? _normBuild(row[iBuild]) : '';
      const reqTs = iReqTs !== -1 ? _parseTs(row[iReqTs]) : null;

      // --- Open / Carryover: only OPEN rows requested in/before this shift's
      //     window. OPEN rows stamped after it belong to the shift now on —
      //     they'll be judged at THEIR boundary. A row with no parseable
      //     requested_ts can't be attributed to the next shift, so it counts
      //     (conservative: the tripwire must not under-report) and is warned.
      if (status === E.STATUS_OPEN) {
        if (!reqTs) { stats.openTotal++; openNoTs++; }
        else if (_requestedInOrBefore(reqTs, target, shift)) stats.openTotal++;
      }

      // --- Intake side: requests received during the target op day + shift ---
      if (reqTs && opDayKey(reqTs) === target.key && shiftOf(reqTs) === shift) {
        stats.requests++;
        const b = buildBucket(level);
        if (b) b.requests++;
      }

      // --- Outcome side: terminal events landing in the target op day + shift ---
      const compTs = iCompTs !== -1 ? _parseTs(row[iCompTs]) : null;
      if (!compTs || opDayKey(compTs) !== target.key || shiftOf(compTs) !== shift) continue;

      if (status === E.STATUS_COMPLETED) {
        stats.completed++;
        if (level === 'Rework') stats.reworks++;

        const cyc = parseFloat(row[iCycle]);
        const b   = buildBucket(level);
        if (b) b.completed++;
        if (iCycle !== -1 && !isNaN(cyc)) {
          stats.cycles.push(cyc);
          if (b) b.cycles.push(cyc);
        }

        const bins = parseInt(row[iBinCnt], 10);
        if (iBinCnt !== -1 && !isNaN(bins)) stats.binsPicked += bins;
      } else if (status === E.STATUS_DISCARDED) {
        stats.discarded++;
      }
    }

    if (openNoTs > 0) {
      console.warn(`ShiftReport.computeShiftStats: ${openNoTs} OPEN row(s) had no parseable ` +
                   'requested_ts — counted as carryover conservatively.');
    }

    return stats;
  }

  // -------------------------------------------------------------------------------
  // Numeric helpers
  // -------------------------------------------------------------------------------
  function _avg(arr)  { return arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0; }
  function _pct(arr, p) {
    if (!arr.length) return 0;
    const s = arr.slice().sort((a, b) => a - b);
    const idx = Math.min(s.length - 1, Math.max(0, Math.ceil(p * s.length) - 1));
    return s[idx];
  }
  function _fmt(n)    { return (Math.round(n * 10) / 10).toFixed(1); }

  // -------------------------------------------------------------------------------
  // HTML rendering — table-based layout with inline styles (email-client safe).
  // Palette mirrors Formatting.js DESIGN so the report reads as the same system.
  // -------------------------------------------------------------------------------
  const C = {
    DARK   : '#1a1a2e',
    MID    : '#16213e',
    ACCENT : '#e94560',
    TILE   : '#f0f4f8',
    BORDER : '#dde3ea',
    MUTED  : '#6b7280',
  };

  function _renderHtml(target, shift, s, opts) {
    const tz     = _tz();
    const dayLbl = Utilities.formatDate(target.date, tz, 'EEE MMM d, yyyy');
    const genLbl = Utilities.formatDate(new Date(), tz, 'EEE MMM d, h:mm a');
    const ssUrl  = SpreadsheetApp.getActiveSpreadsheet().getUrl();

    const kpi = (label, value, accent) => `
      <td style="background:${C.TILE};border:1px solid ${C.BORDER};border-radius:6px;padding:10px 14px;text-align:center;">
        <div style="font-size:10px;letter-spacing:.08em;color:${C.MUTED};text-transform:uppercase;">${label}</div>
        <div style="font-size:24px;font-weight:700;color:${accent || C.DARK};padding-top:2px;">${value}</div>
      </td><td style="width:8px;"></td>`;

    const th = (t, align) =>
      `<th style="background:${C.DARK};color:#fff;font-size:11px;padding:6px 10px;text-align:${align || 'left'};">${t}</th>`;
    const td = (t, align) =>
      `<td style="border-bottom:1px solid ${C.BORDER};font-size:12px;padding:6px 10px;text-align:${align || 'right'};">${t}</td>`;

    // --- No-overlap tripwire: OPEN work at a shift boundary is an anomaly.
    //     (Suppressed on "in progress" snapshots — open work is normal mid-shift.)
    const carryoverBanner = (!opts.partial && s.openTotal > 0) ? `
    <div style="background:#fdecef;border:1px solid ${C.ACCENT};border-radius:6px;padding:10px 14px;margin-bottom:14px;">
      <span style="color:${C.ACCENT};font-weight:700;font-size:13px;">⚠ ${s.openTotal} OPEN request(s) crossed the shift boundary.</span>
      <span style="font-size:12px;color:${C.DARK};"> Per department process, no open work carries over — check the Processing board for what was left behind.</span>
    </div>` : '';

    // --- Build Load Level breakdown: canonical levels always shown (fixed
    //     shape scans faster shift-over-shift); unknown levels appended if hit.
    const levels = CONFIG.BUILD_LEVELS.concat(
      Object.keys(s.byBuild).filter(k => CONFIG.BUILD_LEVELS.indexOf(k) === -1)
    );
    const buildRows = levels.map(l => {
      const b = s.byBuild[l] || { requests: 0, completed: 0, cycles: [] };
      const emphasis = l === 'Rework' && b.completed > 0 ? `color:${C.ACCENT};font-weight:700;` : '';
      return `<tr>
        <td style="border-bottom:1px solid ${C.BORDER};font-size:12px;padding:6px 10px;text-align:left;${emphasis}">${l}</td>
        ${td(b.requests)}${td(b.completed)}${td(b.cycles.length ? _fmt(_avg(b.cycles)) + ' min' : '—')}
      </tr>`;
    }).join('');

    return `
<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;margin:0 auto;color:${C.DARK};">
  <div style="background:${C.DARK};color:#fff;padding:16px 20px;border-radius:6px 6px 0 0;">
    <div style="font-size:16px;font-weight:700;">CARVANA &nbsp;|&nbsp; PARTS DEPARTMENT — ${_shiftName(shift).toUpperCase()} SHIFT REPORT</div>
    <div style="font-size:11px;color:#aaaacc;padding-top:4px;">
      Op Day: ${target.key} (${dayLbl}) &nbsp;·&nbsp; ${_shiftName(shift)} Shift ${_shiftWindow(shift)}
      ${opts.partial ? ` &nbsp;·&nbsp; <span style="color:${C.ACCENT};font-weight:700;">IN PROGRESS</span>` : ''}
    </div>
  </div>

  <div style="padding:16px 0;">
    ${carryoverBanner}
    <table cellspacing="0" cellpadding="0" style="width:100%;"><tr>
      ${kpi('Requests', s.requests)}
      ${kpi('Completed', s.completed)}
      ${kpi('Reworks', s.reworks, s.reworks > 0 ? C.ACCENT : null)}
      ${kpi('Avg Turnaround', _fmt(_avg(s.cycles)) + ' min')}
    </tr></table>

    <div style="margin-top:18px;">
      <div style="font-size:13px;font-weight:700;color:${C.MID};padding-bottom:6px;">By Build Load Level</div>
      <table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;border:1px solid ${C.BORDER};">
        <tr>${th('Build Load Level')}${th('Requests', 'right')}${th('Completed', 'right')}${th('Avg Turnaround', 'right')}</tr>
        ${buildRows}
      </table>
    </div>

    <div style="font-size:11px;color:${C.MUTED};padding-top:12px;">
      Discarded: <b>${s.discarded}</b> &nbsp;·&nbsp;
      Bins picked: <b>${s.binsPicked}</b> &nbsp;·&nbsp;
      Median turnaround: <b>${_fmt(_pct(s.cycles, 0.5))} min</b> &nbsp;·&nbsp;
      P90 turnaround: <b>${_fmt(_pct(s.cycles, 0.9))} min</b> &nbsp;·&nbsp;
      Open / Carryover: <b style="${s.openTotal > 0 ? 'color:' + C.ACCENT + ';' : ''}">${s.openTotal}</b>
    </div>

    <div style="font-size:10px;color:${C.MUTED};padding-top:18px;border-top:1px solid ${C.BORDER};margin-top:18px;">
      ${opts.hasPdf ? `<b style="color:${C.MID};">📎 A print-ready PDF of this report is attached.</b><br>` : ''}
      Generated ${genLbl} · <a href="${ssUrl}" style="color:${C.MID};">Open the Parts Request workbook</a><br>
      Counts only include activity inside this shift's window. Per department process, no open work
      crosses a shift boundary — Open/Carryover above 0 on an end-of-shift report means something was left behind.
    </div>
  </div>
</div>`;
  }

  // Plain-text fallback for clients that strip HTML.
  function _renderText(target, shift, s, opts) {
    const lines = [
      `PARTS DEPARTMENT — ${_shiftName(shift).toUpperCase()} SHIFT REPORT — Op Day ${target.key} (${_shiftWindow(shift)})${opts.partial ? ' (in progress)' : ''}`,
      '',
    ];
    if (!opts.partial && s.openTotal > 0) {
      lines.push(`*** WARNING: ${s.openTotal} OPEN request(s) crossed the shift boundary — no open work should carry over. ***`, '');
    }
    lines.push(
      `Requests: ${s.requests}   Completed: ${s.completed}   Reworks: ${s.reworks}   Avg turnaround: ${_fmt(_avg(s.cycles))} min`,
      `Discarded: ${s.discarded}   Bins picked: ${s.binsPicked}   Median: ${_fmt(_pct(s.cycles, 0.5))}   P90: ${_fmt(_pct(s.cycles, 0.9))}   Open/Carryover: ${s.openTotal}`,
      '',
      'By Build Load Level (requests / completed / avg turnaround):'
    );
    const levels = CONFIG.BUILD_LEVELS.concat(
      Object.keys(s.byBuild).filter(k => CONFIG.BUILD_LEVELS.indexOf(k) === -1)
    );
    levels.forEach(l => {
      const b = s.byBuild[l] || { requests: 0, completed: 0, cycles: [] };
      lines.push(`  ${l}: ${b.requests} / ${b.completed} / ${b.cycles.length ? _fmt(_avg(b.cycles)) + ' min' : '—'}`);
    });
    return lines.join('\n');
  }

  // -------------------------------------------------------------------------------
  // Public surface
  // -------------------------------------------------------------------------------
  return { SHIFTS, opDayKey, shiftOf, computeShiftStats, sendShiftReport, emailMe, defaultRecipients };

})();

// Top-level stub for the menu item
function ShiftReport_emailMe() { ShiftReport.emailMe(); }
