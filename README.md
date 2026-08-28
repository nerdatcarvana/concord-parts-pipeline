# Parts Request Workflow

Google Apps Script backing the Parts Department request board. Pickers work a
single spreadsheet tab; the script handles intake, identity resolution, event
logging, shift reporting, and archival.

Two intake paths converge on one ledger:

```
Google Form ──► Form Responses 1 ──┐
                                   ├──► Intake ──► Ledger ──► _Requests + _EventLog
Walk-in (typed on the board) ──────┘                │
                                                    └──► Processing (the board)
                                                             │
                                          picker enters a bin │
                                                             ▼
                                              PickerTrigger ──► PICKED×n + COMPLETED
                                                                (or DISCARDED)
                                                             │
                                    ~06:00 / ~17:00 boundary │
                                                             ▼
                                              Turnover ──► report + archive + clean slate
```

`_Requests` is the request spine, one row per request. `_EventLog` is the
narrative — every state change, append-only within a shift. The Processing tab
is a **view of open work**, not a record: completed rows are cleared off it.

---

## The shift / op-day model

This is the concept everything else hangs off. Constants live in
`CONFIG.SHIFT` and are the single source of truth — never redeclare these
hours in another module.

| | Window | Notes |
|---|---|---|
| **Op day** | starts at `START_HR` (06:00) | The overnight NIGHT shift collapses into the op day of the DAY shift it *followed* |
| **DAY shift** | `[START_HR, SECOND_HR)` → 06:00–17:00 | |
| **NIGHT shift** | `[SECOND_HR, START_HR)` → 17:00–06:00 | Wraps past midnight |

An op day is keyed `YYYY-MM-DD` — the calendar date the op day *started*. Work
at 02:30 on the 27th belongs to op day 2026-08-26, NIGHT.

**The boundary is a computable instant, not the time a trigger fires.** DAY
ends at `opDay @ SECOND_HR`; NIGHT ends at `opDay+1 @ START_HR`. Apps Script
time triggers only guarantee an hour bucket, so the 17:00 trigger actually
lands around 17:54 and the 06:00 trigger around 06:13. Anything that decides
what belongs to which shift compares against the boundary instant, never
against `new Date()`. Code that got this wrong is what produced the rev-3.3 bug
described below.

**Department process rule: no open work crosses a shift boundary.** Each shift
finishes what it starts, so "Open / Carryover" should always read 0. A nonzero
count is treated as an anomaly — red banner on the report, warning in the
turnover message — not a neutral statistic.

---

## Sheets

| Sheet | Role | Cleared at turnover? |
|---|---|---|
| `Form Responses 1` | Google Form landing tab. Gains a `Sync Status` column the form does not know about. | Yes, with survivor rules |
| `Processing` | The picker-facing board. Open work only. | Yes, with survivor rules |
| `_Requests` | Request spine — one row per request, current status. | Yes, with survivor rules |
| `_EventLog` | Every state change. | Yes, with survivor rules |
| `_Identity` | alias → canonical email / display / department. Configuration. | **Never** |
| `_Roles` | email → MANAGER / PICKER. Configuration. | **Never** |

Reference sheets are hidden and protected. On `Processing`, pickers can edit
only Bin Location, Stock Number, Requester, Build, Lot Status, and Notes;
everything else is locked to the owner and managers.

Sheets are resolved **by gid, not by name** (`Sheets.js`). A picker once
renamed the Processing tab to a stock number and the whole pick flow stopped —
range protections don't help, because renaming a tab is a structural operation
available to any editor. The registry maps role → gid in Script Properties,
self-heals drifted names back to canonical, and survives renames entirely.

> `Sheets.getByName()` must never be pointed at an archive copy. On a gid miss
> it adopts by name and writes that sheet's id back into the shared registry —
> pointing it at a copy would overwrite the live workbook's gids. `Turnover`
> uses a local `_archSheet()` helper for exactly this reason.

---

## Modules

| File | Responsibility |
|---|---|
| `Config.js` | Sheet names, headers, enums, shift hours, thresholds. Change a value here and it propagates. Never hard-code a sheet name or header string elsewhere. |
| `Sheets.js` | Rename-proof sheet resolution by gid, with self-healing. |
| `Utils.js` | Stateless helpers: header→column index, ISO timestamps, UUIDs, `withLock`, toasts, value normalization. No side effects. |
| `Identity.js` | Resolves any requester reference (alias, nickname, email, typo) to a canonical triple. Unresolvable aliases go to a backlog the health monitor reports. |
| `Ledger.js` | The **only** module that writes `_EventLog` and `_Requests`. Row lookups use `createTextFinder` scoped to one column so a pick costs the same on day 1 and day 1,000. |
| `Intake.js` | Both intake paths — form sync and walk-in. Idempotent: the `Sync Status` flag prevents double-processing. |
| `PickerTrigger.js` | Installable `onEdit`. Stock Number → walk-in intake; Bin Location → parse bins, emit PICKED×n + COMPLETED (or DISCARDED). Runs as the owner, so it can write protected sheets the picker cannot. |
| `Reconcile.js` | Flips a DISCARDED request to SUPERSEDED when the same stock number was later COMPLETED, so the discard rate reflects work that was ultimately fulfilled. |
| `ShiftReport.js` | Shift-scoped stats and the HTML report email. |
| `ReportPdf.js` | Print-designed PDF attachment. Charts are nested tables, not images — Google's HTML→PDF converter has no flexbox, grid, SVG, or web fonts. Decorative: failures are swallowed so the email still sends. |
| `Turnover.js` | The boundary job: reconcile → stats → archive → clean slate → prune → email. |
| `Health.js` | Two-layer monitoring — a 2-minute post-submit heartbeat (primary) and an hourly sweep (backstop). |
| `Formatting.js` | All board styling, validation, and conditional formatting. Idempotent. Rule order is load-bearing (first match wins in Sheets). |
| `Access.js` | Roles, menu construction, sheet protections. |
| `Setup.js` | Idempotent bootstrap. Run once as the owner after every push. |
| `Rebuild.js` | Regenerates the whole workbook artifact from code. Structure only — history and open rows are preserved. |

---

## Triggers

Installed by `Setup._installTriggers`. Re-running Setup removes and reinstalls
the whole set, including retired handlers from earlier versions.

| Handler | Type | Purpose |
|---|---|---|
| `onEditTrigger` | onEdit | Board interactions — walk-in entry, bin pick |
| `onFormSubmitTrigger` | onFormSubmit | Sync the new response immediately |
| `onFormSubmitHeartbeat` | onFormSubmit | Schedule a 2-min verification check |
| `Health_heartbeatCheck` | one-time, +2 min | Verify the sync landed; retry and alert if not. Deletes its own trigger. |
| `Health_monitor` | hourly | Backstop sweep — stuck rows, unmapped aliases |
| `Turnover_run` | daily, `atHour(SECOND_HR)` | DAY boundary |
| `Turnover_run` | daily, `atHour(START_HR)` | NIGHT boundary |

**The turnover schedule is deliberately loose.** `.nearMinute()` is ±15
minutes, and anything that can fire *before* a boundary is worse than firing
late: `_endedShift()` would resolve to the previous boundary, `_alreadyRan()`
would match that morning's stamp, and the run would return silently — no
archive, no wipe, console log only, with the next boundary quietly absorbing
two shifts. The observed 17:54 slot has 54 minutes of margin and has never
fired early across 32 runs. Leave it.

The durable fix is self-identifying handlers (`Turnover_runDayBoundary` /
`Turnover_runNightBoundary` passing an expected shift) so `run()` validates
against intent rather than inferring the boundary from the clock. Not yet
implemented; it touches the Setup trigger inventory.

---

## Turnover: what one run does

Steps 0–4 hold the script lock. Steps 5–6 do not — the archive is committed
and the live workbook is done being modified, so holding the lock through more
work during shift change (the busiest minutes of the day for `onEdit` and
`onFormSubmit`) buys nothing.

0. **Reconcile** re-fulfilled discards, so the report's discard rate is right.
1. **Snapshot stats** for the ended shift. Must precede the wipe — it reads
   rows the wipe removes.
   - *Nothing-to-archive guard:* an automated run whose ended shift had zero
     activity, zero carryover, and no data predating the boundary makes no
     snapshot and wipes nothing. It only stamps the once-per guard.
2. **Archive** — a full Drive copy of the workbook, before anything is touched.
   If the copy throws, nothing is wiped. This is the ARCHIVE-FIRST invariant.
3. **Clean slate** — reset the four data sheets, honoring survivor rules.
4. Re-apply formatting; stamp the once-per-(op day, shift) guard.
5. **Prune** the archive copy down to the ended shift.
6. **Email** the report to all managers.

### What survives the wipe

Do not "simplify" these away.

- **OPEN requests**, plus their event rows and board rows. These should be zero
  by process rule; when they aren't, destroying live requests is not an
  acceptable failure mode.
- **Terminal rows completed in the current live shift.** The 17:00 run really
  fires at 17:5x, so picks completed in those minutes belong to the shift now
  on and are counted at the *next* boundary. Wipe them and the next report
  undercounts.
- **In-flight board rows** — content but no Request ID yet, i.e. a walk-in
  someone is mid-typing.
- **Form responses stamped at or after the boundary**, and **any response that
  hasn't reached `_Requests` yet**.

### Archive semantics

Membership is decided by **intake time**; state is **as-of the snapshot**. A
day request at 16:50 that a night picker completes at 17:20 stays in the DAY
archive, showing COMPLETED. `_EventLog`, `Processing`, and `Form Responses`
follow their *request*, not their own timestamps — filtering the log on
`event_ts` would leave that request in the DAY archive with its PICKED and
COMPLETED rows orphaned.

Archives therefore partition rather than overlap wholesale. The overlap that
remains is genuine: an OPEN row predating two consecutive boundaries appears in
both, which is exactly where you want to see it twice.

Archives are inert. Installable triggers do not survive `makeCopy()`, and the
duplicate Google Form that `makeCopy()` creates for a form-linked workbook is
unlinked and trashed immediately — otherwise a submission to the stray form
would write into an archived snapshot.

Files land at
`Parts Request Backups / YYYY / MM / Week N / Turnover_{opDay}_{SHIFT}_{HHmm}`.

A missing file means either "empty shift" or "failed run" — the Apps Script
trigger-failure email disambiguates.

---

## Configuration knobs

All in `Config.js`.

| Key | Default | Effect |
|---|---|---|
| `SHIFT.START_HR` / `SHIFT.SECOND_HR` | `6` / `17` | Op-day boundary and shift start. Propagates to the triggers on the next Setup run. |
| `TURNOVER_SKIP_EMPTY_REPORT` | `true` | Suppress the email when an automated run's shift had no activity and no carryover. Manual runs always send. |
| `TURNOVER_PRUNE_ARCHIVE` | `true` | Trim the archive to the ended shift. **Kill switch** — set `false` to fall back to whole-workbook snapshots without a redeploy. The fallback is over-inclusive, never lossy. |
| `LOCK_TIMEOUT_MS` | `15000` | `withLock` wait before giving up. |
| `HEALTH_STUCK_MINUTES` | `3` | Unprocessed response age before the hourly monitor alerts. |
| `PROCESSING_MIN_ROWS` | `50` | Row budget kept formatted on the board. |
| `PROCESSING_STALE_WARN_HRS` / `_CRIT_HRS` | `1` / `3` | Board aging bands — amber, then red. |
| `BACKUP_WEEK_OF_MONTH` | `true` | Week-of-month folders vs. ISO week numbers. |
| `ADMIN_EMAIL` | script owner | Override with a Script Property of the same name rather than hard-coding. |

---

## Menu

`🛠️ Parts Management` — built by `Access.buildMenu()` on open, role-aware.

Everyone:
- 📥 Sync Form Submissions
- 📧 Email Shift Report — in-progress snapshot of the *current* shift, to you only

Managers also get:
- 🗄️ End of Day Archive — a manual turnover. Confirms, always proceeds, always mails.
- 🔁 Re-apply Formatting / 🔒 Re-apply Protections
- 👤 Clear Unmapped Alias Queue
- 🔧 Re-run Setup
- ♻️ Rebuild Entire Workbook / ♻️ Reconcile Re-fulfilled Discards

---

## Operations

**After every push:** run `Setup.initialize()` once as the script owner. It is
idempotent — creates what's missing, skips what exists, reinstalls triggers,
reapplies protections and formatting.

**Adding a scope to `appsscript.json` requires re-authorization.** The trigger
owner must run any function once from the editor. Skip it and trigger runs fail
on auth instead of degrading. This bit us once already: the forms scope was
missing, `FormApp.openByUrl()` was denied, the best-effort catch swallowed it
into a log nobody read, and one stray "Copy of…" form accumulated per boundary.

**A stuck form response.** The 2-minute heartbeat retries and emails on
failure. If one persists, run Sync Form Submissions from the menu. Since rev
3.3 an unsynced response **survives turnover** instead of being silently
deleted, so it will keep the health monitor alerting until a human clears it.
That's the intended direction — loud beats lossy — but it means a permanently
unsyncable row needs attention rather than waiting.

**An unmapped alias.** Add it to `_Identity` (alias → canonical_email). The
hourly monitor lists the backlog.

**Workbook damaged.** `Rebuild.everything()` rebuilds structure, formatting,
protections, and triggers. `_EventLog`, `_Requests`, and OPEN board rows are
preserved — it rebuilds structure, not history.

**The board looks wrong.** A lingering pending tint (soft blue) means a bin was
entered but the pick execution never landed — retry the bin. Amber and red are
aging bands on open unpicked rows.

---

## Repo layout

```
.
├── .clasp.json          # not committed — see below
├── .claspignore
├── README.md
└── src/
    ├── appsscript.json
    ├── Access.js
    ├── Config.js
    ├── Formatting.js
    ├── Health.js
    ├── Identity.js
    ├── Intake.js
    ├── Ledger.js
    ├── PickerTrigger.js
    ├── Rebuild.js
    ├── Reconcile.js
    ├── ReportPdf.js
    ├── Setup.js
    ├── Sheets.js
    ├── ShiftReport.js
    ├── Turnover.js
    └── Utils.js
```

Apps Script has no module system — every file shares one global scope, and
load order is not guaranteed. Hence the IIFE-per-module pattern with a small
returned surface, and hence the rule that `CONFIG` values are read *inside*
functions at call time rather than captured at load time.

### Example `.clasp.json`

Keep this out of version control — it carries your script id.

```json
{
  "scriptId": "1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789_ExampleScriptId",
  "rootDir": "src",
  "scriptExtensions": [".js", ".gs"],
  "jsonExtensions": [".json"],
  "filePushOrder": ["src/Config.js", "src/Utils.js", "src/Sheets.js"],
  "fileExtensions": [".js", ".gs", ".json"]
}
```

`rootDir: "src"` keeps the README and repo tooling out of the push.
`filePushOrder` is cosmetic rather than load-bearing — the Apps Script runtime
resolves the global scope itself — but pushing the foundations first keeps the
editor's file list in a sensible order.

### `.claspignore`

```
**/**
!src/**/*.js
!src/appsscript.json
```

clasp's ignore syntax is inverted from `.gitignore`: exclude everything, then
allow back in.

### `.gitignore`

```
.clasp.json
.clasprc.json
node_modules/
```

---

## Known issues and drift

- **`Ledger.js` header is stale.** It says `_EventLog` and `_Requests` are
  "append-only and NEVER cleared." That was true before rev 3 turnover. They
  are now cleared at each boundary with survivor rules, and immutability lives
  in the archive chain rather than the live workbook. The code is correct; the
  comment is not.
- **Boundary inference is clock-based.** See the trigger section — a run that
  fires before its boundary is silently absorbed. Mitigated by margin, not
  fixed.
- **`Access.js` menu still points at `Eod_runEod`.** That top-level stub is a
  live alias for `Turnover.run()` and works fine, but the name predates the
  turnover rewrite.
