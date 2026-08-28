/**
 * ReportPdf.gs
 * Renders the shift report as a print-designed PDF, attached to the shift
 * report email alongside the existing HTML body.
 *
 * WHY A SEPARATE ARTIFACT: the email body is tuned for a glance in an inbox.
 * Printed, it inherits the client's margins, link underlines and font
 * substitution, and reads as a screenshot of an email. This builds a document
 * meant for paper — fixed width, print-safe type scale, real charts.
 *
 * RENDERER CONSTRAINTS — read before editing:
 *   Google's HTML→PDF converter (Utilities.newBlob(...).getAs(PDF)) is a
 *   limited engine. It reliably handles tables, inline styles, background
 *   colours, borders and block text. It does NOT handle flexbox, CSS grid,
 *   SVG, external or data-URI images, web fonts, box-shadow or absolute
 *   positioning. Every chart here is therefore built from nested tables with
 *   percentage widths and background fills — which is why they render the same
 *   in the converter, in a browser, and on paper.
 *
 *   Keep it that way. A chart image would look better in one place and vanish
 *   in another.
 *
 * FAILURE POLICY: PDF generation is decorative. Any failure is swallowed by
 * the caller so the shift report email still sends — a turnover report that
 * does not arrive is a real problem; one without an attachment is not.
 *
 * Public API:
 *   ReportPdf.build(target, shift, stats, opts)      → Blob (throws on failure)
 *   ReportPdf.renderHtml(target, shift, stats, opts) → HTML string (testable)
 *   ReportPdf.fileName(target, shift, opts)          → 'Parts-Shift-Report_…pdf'
 */

const ReportPdf = (() => {

  // --------------------------------------------------------------------------
  // Palette — extends the tokens in Formatting.js / ShiftReport.js so the
  // printed report, the email and the board all read as one system.
  // --------------------------------------------------------------------------
  const P = {
    INK       : '#1a1a2e',   // headline / body ink (matches DESIGN.HEADER_BG)
    MID       : '#16213e',
    ACCENT    : '#e94560',   // rework / discard / alert
    MUTED     : '#6b7280',
    RULE      : '#dde3ea',
    TILE      : '#f4f6fa',
    PAPER     : '#ffffff',

    // Chart fills
    BAR_REQ   : '#c3cfe2',   // requests — pale navy
    BAR_DONE  : '#2b4c7e',   // completed — deep navy
    SEG_DONE  : '#2b6cb0',   // outcome mix: completed
    SEG_DISC  : '#e94560',   // outcome mix: discarded
    SEG_OPEN  : '#e8a33d',   // outcome mix: still open (carryover)
    HIST      : '#3f6493',   // turnaround histogram
  };

  const FONT = 'Helvetica, Arial, sans-serif';
  const PAGE_W = 720;        // px of content width; converter renders at letter

  // --------------------------------------------------------------------------
  // CA — print-color-adjust, prepended to the style of EVERY filled element.
  //
  // Google's converter is headless Chromium, and it prints with background
  // graphics OFF (Chromium's default). Without this declaration every
  // background-color is silently dropped: text and borders render, and every
  // chart bar, KPI tile and the masthead band come out blank white. Verified
  // against Chromium 141 — `print-color-adjust:exact` produces output
  // byte-identical to printBackground:true.
  //
  // It is applied INLINE rather than via a <style> block on purpose: inline
  // styles are proven to survive the converter (borders and text colours came
  // through on a live run), a stylesheet is not. The property inherits, but it
  // is repeated on each filled element so no wrapper the converter may inject
  // can break the chain.
  // --------------------------------------------------------------------------
  const CA = 'print-color-adjust:exact;-webkit-print-color-adjust:exact;';

  // --------------------------------------------------------------------------
  // Escaping — build level names come from sheet data and can be anything.
  // An unescaped '<' would silently corrupt the document.
  // --------------------------------------------------------------------------
  function _esc(v) {
    return String(v === null || v === undefined ? '' : v)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function _num(n) {
    const x = Number(n);
    return isFinite(x) ? x : 0;
  }

  function _fmt1(n) { return (Math.round(_num(n) * 10) / 10).toFixed(1); }

  // Percentage of a whole, clamped to [0,100] and safe when whole is 0.
  function _pctOf(part, whole) {
    const w = _num(whole);
    if (w <= 0) return 0;
    const p = (_num(part) / w) * 100;
    return Math.max(0, Math.min(100, p));
  }

  function _avg(arr)  {
    const a = (arr || []).map(_num);
    return a.length ? a.reduce((x, y) => x + y, 0) / a.length : 0;
  }
  function _pctile(arr, p) {
    const a = (arr || []).map(_num).sort((x, y) => x - y);
    if (!a.length) return 0;
    const i = Math.min(a.length - 1, Math.max(0, Math.ceil(p * a.length) - 1));
    return a[i];
  }

  // --------------------------------------------------------------------------
  // Chart primitives — nested tables only. See renderer constraints above.
  // --------------------------------------------------------------------------

  // A single horizontal bar occupying `pct` of the track width.
  function _bar(pct, colour, height) {
    const w = Math.round(_pctOf(pct, 100) * 100) / 100;
    const h = height || 10;
    // Two cells: the filled portion and the remainder. Percentage widths on
    // table cells are the one layout primitive the converter never gets wrong.
    return `<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;table-layout:fixed;">
      <tr>
        <td style="${CA}width:${w}%;height:${h}px;background:${colour};font-size:1px;line-height:${h}px;">&nbsp;</td>
        <td style="${CA}width:${100 - w}%;height:${h}px;background:${P.TILE};font-size:1px;line-height:${h}px;">&nbsp;</td>
      </tr>
    </table>`;
  }

  // A stacked bar from [{ pct, colour }] segments; remainder filled with TILE.
  function _stackedBar(segments, height) {
    const h = height || 18;
    const cells = segments
      .filter(s => s.pct > 0)
      .map(s => `<td style="${CA}width:${Math.round(s.pct * 100) / 100}%;height:${h}px;background:${s.colour};font-size:1px;line-height:${h}px;">&nbsp;</td>`)
      .join('');
    const used = segments.reduce((t, s) => t + (s.pct > 0 ? s.pct : 0), 0);
    const rest = Math.max(0, 100 - used);
    const tail = rest > 0
      ? `<td style="${CA}width:${Math.round(rest * 100) / 100}%;height:${h}px;background:${P.TILE};font-size:1px;line-height:${h}px;">&nbsp;</td>`
      : '';
    return `<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;table-layout:fixed;">
      <tr>${cells}${tail}</tr></table>`;
  }

  // Legend entry. The colour swatch is a table cell rather than an
  // inline-block span — the converter does not reliably honour inline-block,
  // and a collapsed swatch would leave the legend unreadable.
  function _legendDot(colour, label, value) {
    return `<td style="padding:0 18px 0 0;vertical-align:middle;">
      <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;"><tr>
        <td width="8" style="${CA}width:8px;height:8px;background:${colour};font-size:1px;line-height:8px;">&nbsp;</td>
        <td style="padding-left:6px;font-size:10px;color:${P.MUTED};white-space:nowrap;">${_esc(label)}${
          value === '' || value === null || value === undefined
            ? ''
            : ` <b style="color:${P.INK};">${_esc(value)}</b>`}</td>
      </tr></table>
    </td>`;
  }

  function _sectionTitle(t) {
    return `<div style="font-size:10px;font-weight:700;color:${P.MID};letter-spacing:.1em;
      text-transform:uppercase;padding:0 0 7px 0;border-bottom:2px solid ${P.INK};margin-bottom:11px;">${_esc(t)}</div>`;
  }

  function _kpi(label, value, sub, accent) {
    return `<td width="25%" style="vertical-align:top;">
      <table cellpadding="0" cellspacing="0" style="${CA}width:100%;border-collapse:collapse;background:${P.TILE};border:1px solid ${P.RULE};">
        <tr><td style="padding:11px 12px;text-align:center;">
          <div style="font-size:8.5px;letter-spacing:.11em;color:${P.MUTED};text-transform:uppercase;">${_esc(label)}</div>
          <div style="font-size:27px;font-weight:700;color:${accent || P.INK};padding:3px 0 0 0;line-height:1.05;">${_esc(value)}</div>
          <div style="font-size:8.5px;color:${P.MUTED};padding-top:2px;">${_esc(sub || '')}&nbsp;</div>
        </td></tr>
      </table>
    </td>`;
  }

  // --------------------------------------------------------------------------
  // Sections
  // --------------------------------------------------------------------------

  function _outcomeMix(s) {
    const done = _num(s.completed), disc = _num(s.discarded), open = _num(s.openTotal);
    const total = done + disc + open;

    if (total <= 0) {
      return _sectionTitle('Outcome Mix') +
        `<div style="font-size:11px;color:${P.MUTED};padding:6px 0 2px 0;">No terminal activity recorded in this window.</div>`;
    }

    const segs = [
      { pct: _pctOf(done, total), colour: P.SEG_DONE },
      { pct: _pctOf(disc, total), colour: P.SEG_DISC },
      { pct: _pctOf(open, total), colour: P.SEG_OPEN },
    ];

    const rate = Math.round(_pctOf(done, total));

    return _sectionTitle('Outcome Mix') +
      _stackedBar(segs, 20) +
      `<table cellpadding="0" cellspacing="0" style="width:100%;padding-top:8px;"><tr>
        ${_legendDot(P.SEG_DONE, 'Completed', done)}
        ${_legendDot(P.SEG_DISC, 'Discarded', disc)}
        ${_legendDot(P.SEG_OPEN, 'Open / carryover', open)}
        <td style="text-align:right;font-size:10px;color:${P.MUTED};">Completion rate
          <b style="color:${P.INK};font-size:12px;">${rate}%</b></td>
      </tr></table>`;
  }

  function _buildLevels(s) {
    const byBuild = s.byBuild || {};
    const known   = (typeof CONFIG !== 'undefined' && CONFIG.BUILD_LEVELS) ? CONFIG.BUILD_LEVELS : [];
    const levels  = known.concat(Object.keys(byBuild).filter(k => known.indexOf(k) === -1));

    let peak = 0;
    levels.forEach(l => {
      const b = byBuild[l] || {};
      peak = Math.max(peak, _num(b.requests), _num(b.completed));
    });

    if (peak <= 0) {
      return _sectionTitle('Throughput by Build Load Level') +
        `<div style="font-size:11px;color:${P.MUTED};padding:6px 0 2px 0;">No requests recorded in this window.</div>`;
    }

    const rows = levels.map(l => {
      const b   = byBuild[l] || { requests: 0, completed: 0, cycles: [] };
      const req = _num(b.requests), don = _num(b.completed);
      const cyc = (b.cycles && b.cycles.length) ? _fmt1(_avg(b.cycles)) + ' min' : '—';
      const isRework = (l === 'Rework' && don > 0);
      const nameCol  = isRework ? P.ACCENT : P.INK;

      return `<tr>
        <td width="72" style="font-size:11px;font-weight:700;color:${nameCol};padding:5px 8px 5px 0;vertical-align:middle;">${_esc(l)}</td>
        <td style="padding:5px 0;vertical-align:middle;">
          <table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;">
            <tr><td style="padding-bottom:3px;">${_bar(_pctOf(req, peak), P.BAR_REQ, 9)}</td></tr>
            <tr><td>${_bar(_pctOf(don, peak), P.BAR_DONE, 9)}</td></tr>
          </table>
        </td>
        <td width="112" style="font-size:10px;color:${P.MUTED};text-align:right;padding:5px 0 5px 12px;vertical-align:middle;white-space:nowrap;">
          <b style="color:${P.INK};">${req}</b> req &middot; <b style="color:${P.INK};">${don}</b> done
        </td>
        <td width="74" style="font-size:10px;color:${P.MUTED};text-align:right;padding:5px 0 5px 10px;vertical-align:middle;white-space:nowrap;">${_esc(cyc)}</td>
      </tr>`;
    }).join('');

    return _sectionTitle('Throughput by Build Load Level') +
      `<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;">${rows}</table>
       <table cellpadding="0" cellspacing="0" style="width:100%;padding-top:7px;"><tr>
         ${_legendDot(P.BAR_REQ, 'Requested', '')}
         ${_legendDot(P.BAR_DONE, 'Completed', '')}
         <td style="text-align:right;font-size:9px;color:${P.MUTED};">right column: average turnaround</td>
       </tr></table>`;
  }

  const BUCKETS = [
    { label: '\u2264 15 min',  test: (v) => v <= 15 },
    { label: '16 \u2013 30',   test: (v) => v > 15  && v <= 30 },
    { label: '31 \u2013 60',   test: (v) => v > 30  && v <= 60 },
    { label: '61 \u2013 120',  test: (v) => v > 60  && v <= 120 },
    { label: '> 120 min',      test: (v) => v > 120 },
  ];

  function _turnaround(s) {
    const cycles = (s.cycles || []).map(_num).filter(v => isFinite(v));

    if (!cycles.length) {
      return _sectionTitle('Turnaround Distribution') +
        `<div style="font-size:11px;color:${P.MUTED};padding:6px 0 2px 0;">No completed requests with a recorded turnaround.</div>`;
    }

    const counts = BUCKETS.map(b => cycles.filter(b.test).length);
    const peak   = Math.max.apply(null, counts);

    const rows = BUCKETS.map((b, i) => `
      <tr>
        <td width="78" style="font-size:10px;color:${P.MUTED};padding:3px 8px 3px 0;text-align:right;vertical-align:middle;white-space:nowrap;">${b.label}</td>
        <td style="padding:3px 0;vertical-align:middle;">${_bar(_pctOf(counts[i], peak), P.HIST, 11)}</td>
        <td width="34" style="font-size:10px;font-weight:700;color:${P.INK};text-align:right;padding:3px 0 3px 9px;vertical-align:middle;">${counts[i]}</td>
      </tr>`).join('');

    const stat = (lbl, val) => `
      <td style="text-align:center;padding:0 4px;">
        <div style="font-size:8.5px;letter-spacing:.09em;color:${P.MUTED};text-transform:uppercase;">${_esc(lbl)}</div>
        <div style="font-size:14px;font-weight:700;color:${P.INK};padding-top:1px;">${_esc(val)}</div>
      </td>`;

    return _sectionTitle('Turnaround Distribution') +
      `<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;">${rows}</table>
       <table cellpadding="0" cellspacing="0" style="${CA}width:100%;margin-top:10px;border-collapse:collapse;background:${P.TILE};border:1px solid ${P.RULE};">
         <tr><td style="padding:9px 6px;">
           <table cellpadding="0" cellspacing="0" style="width:100%;"><tr>
             ${stat('Fastest', _fmt1(Math.min.apply(null, cycles)))}
             ${stat('Median',  _fmt1(_pctile(cycles, 0.5)))}
             ${stat('Average', _fmt1(_avg(cycles)))}
             ${stat('P90',     _fmt1(_pctile(cycles, 0.9)))}
             ${stat('Slowest', _fmt1(Math.max.apply(null, cycles)))}
           </tr></table>
         </td></tr>
       </table>
       <div style="font-size:8.5px;color:${P.MUTED};padding-top:4px;text-align:right;">all values in minutes</div>`;
  }

  function _carryoverBanner(s, opts) {
    if (opts.partial || _num(s.openTotal) <= 0) return '';
    return `<table cellpadding="0" cellspacing="0" style="${CA}width:100%;border-collapse:collapse;background:#fdecef;border:1px solid ${P.ACCENT};margin-bottom:15px;">
      <tr>
        <td width="5" style="${CA}background:${P.ACCENT};font-size:1px;">&nbsp;</td>
        <td style="padding:9px 13px;">
          <div style="font-size:12px;font-weight:700;color:${P.ACCENT};">
            ${_num(s.openTotal)} OPEN request(s) crossed the shift boundary
          </div>
          <div style="font-size:10px;color:${P.INK};padding-top:2px;">
            Per department process no open work carries over. Check the Processing board for what was left behind.
          </div>
        </td>
      </tr>
    </table>`;
  }

  // --------------------------------------------------------------------------
  // renderHtml(target, shift, stats, opts)
  // --------------------------------------------------------------------------
  function renderHtml(target, shift, stats, opts) {
    opts = opts || {};
    const s = stats || {};

    const shiftName = opts.shiftName ||
      (String(shift).toUpperCase() === 'DAY' ? 'Day' : 'Night');

    const dayLbl = opts.dayLabel || _esc(target && target.key);
    const genLbl = opts.generatedLabel || '';
    const window = opts.shiftWindow || '';

    return `<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<style>*{print-color-adjust:exact;-webkit-print-color-adjust:exact;}</style>
<body style="${CA}margin:0;padding:0;background:${P.PAPER};">
<table cellpadding="0" cellspacing="0" style="${CA}width:${PAGE_W}px;margin:0 auto;border-collapse:collapse;font-family:${FONT};color:${P.INK};">

  <!-- masthead -->
  <tr><td style="${CA}background:${P.INK};padding:17px 22px;">
    <table cellpadding="0" cellspacing="0" style="width:100%;"><tr>
      <td style="vertical-align:middle;">
        <div style="font-size:19px;font-weight:700;color:#ffffff;letter-spacing:.16em;">CARVANA</div>
        <div style="font-size:9.5px;color:#9aa4c4;letter-spacing:.19em;padding-top:3px;">PARTS DEPARTMENT</div>
      </td>
      <td style="vertical-align:middle;text-align:right;">
        <div style="font-size:14px;font-weight:700;color:#ffffff;">${_esc(shiftName)} Shift Report</div>
        <div style="font-size:9.5px;color:#9aa4c4;padding-top:3px;">
          Op Day ${_esc(target && target.key)}${window ? ' &middot; ' + _esc(window) : ''}
        </div>
        ${opts.partial ? `<div style="font-size:9px;color:${P.ACCENT};font-weight:700;padding-top:3px;letter-spacing:.09em;">IN PROGRESS \u2014 SNAPSHOT</div>` : ''}
      </td>
    </tr></table>
  </td></tr>
  <tr><td style="${CA}height:3px;background:${P.ACCENT};font-size:1px;">&nbsp;</td></tr>

  <!-- body -->
  <tr><td style="padding:19px 22px 0 22px;">
    ${_carryoverBanner(s, opts)}
    ${dayLbl ? `<div style="font-size:10px;color:${P.MUTED};padding-bottom:13px;">${_esc(dayLbl)}</div>` : ''}

    <table cellpadding="0" cellspacing="5" style="width:100%;border-collapse:separate;margin-bottom:16px;"><tr>
      ${_kpi('Requests',  _num(s.requests),  'received this shift')}
      ${_kpi('Completed', _num(s.completed), _num(s.binsPicked) + ' bins picked')}
      ${_kpi('Reworks',   _num(s.reworks),   'rework level', _num(s.reworks) > 0 ? P.ACCENT : null)}
      ${_kpi('Avg Turnaround', _fmt1(_avg(s.cycles)), 'minutes')}
    </tr></table>

    <div style="margin-bottom:20px;">${_outcomeMix(s)}</div>
    <div style="margin-bottom:20px;">${_buildLevels(s)}</div>
    <div style="margin-bottom:14px;">${_turnaround(s)}</div>
  </td></tr>

  <!-- footer -->
  <tr><td style="padding:0 22px 18px 22px;">
    <div style="border-top:1px solid ${P.RULE};padding-top:9px;font-size:8.5px;color:${P.MUTED};line-height:1.5;">
      ${genLbl ? 'Generated ' + _esc(genLbl) + ' &middot; ' : ''}Parts Request Workflow${opts.partial ? ' &middot; in-progress snapshot, figures will change before end of shift' : ''}<br>
      Counts include only activity inside this shift's window. Per department process no open work crosses a
      shift boundary &mdash; an Open / carryover figure above zero on an end-of-shift report means work was left behind.
    </div>
  </td></tr>

</table>
</body></html>`;
  }

  // --------------------------------------------------------------------------
  // fileName / build
  // --------------------------------------------------------------------------
  function fileName(target, shift, opts) {
    opts = opts || {};
    const name = String(shift).toUpperCase() === 'DAY' ? 'Day' : 'Night';
    const key  = String((target && target.key) || 'report').replace(/[^0-9A-Za-z-]/g, '');
    return `Parts-Shift-Report_${key}_${name}${opts.partial ? '_in-progress' : ''}.pdf`;
  }

  function build(target, shift, stats, opts) {
    const html = renderHtml(target, shift, stats, opts);
    const name = fileName(target, shift, opts);
    return Utilities.newBlob(html, 'text/html', name)
      .getAs('application/pdf')
      .setName(name);
  }

  return { build, renderHtml, fileName, _palette: P };

})();
