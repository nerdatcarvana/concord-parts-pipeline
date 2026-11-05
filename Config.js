/**
 * Config.gs
 * Single source of truth for all sheet names, column definitions, enums,
 * bin taxonomy constants, and tunable thresholds.
 *
 * Changing a value here propagates everywhere. Never hard-code sheet names,
 * header strings, or magic numbers in other modules.
 */

const CONFIG = {

  // ---------------------------------------------------------------------------------
  // Sheet names
  // ---------------------------------------------------------------------------------
  SHEETS: {
    FORM_RESPONSES : 'Form Responses 1',
    PROCESSING     : 'Processing',
    EVENT_LOG      : '_EventLog',
    REQUESTS       : '_Requests',
    IDENTITY       : '_Identity',
    ROLES          : '_Roles',
  },

  // ---------------------------------------------------------------------------------
  // Column header definitions — used by Utils.getColIndex()
  // These must match the actual sheet headers exactly (Utils strips invisible
  // chars and trims before comparing, so leading/trailing whitespace is fine).
  // ---------------------------------------------------------------------------------
  HEADERS: {

    // Form Responses 1
    FR_TIMESTAMP    : 'Timestamp',
    FR_EMAIL        : 'Email Address',
    FR_STOCK        : 'Stock Number',
    FR_BUILD        : 'Build Load Level',
    FR_LOT          : 'Lot Status',
    FR_SYNC_STATUS  : 'Sync Status',  // appended by this system, not the form

    // Processing board
    PROC_REQUESTED  : 'Requested',
    PROC_REQUESTER  : 'Requester',
    PROC_STOCK      : 'Stock Number',
    PROC_BUILD      : 'Build Load Level',
    PROC_LOT        : 'Lot Status',
    PROC_BIN        : 'Bin Location',
    PROC_REQ_ID     : 'Request ID',
    PROC_NOTES      : 'Notes',

    // _EventLog
    EL_EVENT_ID     : 'event_id',
    EL_REQUEST_ID   : 'request_id',
    EL_EVENT_TYPE   : 'event_type',
    EL_EVENT_TS     : 'event_ts',
    EL_ACTOR_EMAIL  : 'actor_email',
    EL_ACTOR_DISPLAY: 'actor_display',
    EL_INTAKE_SRC   : 'intake_source',
    EL_STOCK        : 'stock_number',
    EL_BUILD        : 'build_load_level',
    EL_LOT          : 'lot_status',
    EL_BIN_CODE     : 'bin_code',
    EL_LOC_TYPE     : 'location_type',
    EL_REQ_DEPT     : 'requester_department',
    EL_RAW_INPUT    : 'raw_input',
    EL_NOTES        : 'notes',

    // _Requests
    RQ_REQUEST_ID   : 'request_id',
    RQ_INTAKE_SRC   : 'intake_source',
    RQ_REQUESTED_TS : 'requested_ts',
    RQ_REQ_EMAIL    : 'requester_email',
    RQ_REQ_DISPLAY  : 'requester_display',
    RQ_REQ_DEPT     : 'requester_department',
    RQ_STOCK        : 'stock_number',
    RQ_BUILD        : 'build_load_level',
    RQ_LOT          : 'lot_status',
    RQ_BIN_COUNT    : 'bin_count',
    RQ_BINS         : 'bins',
    RQ_FIRST_PICK_TS: 'first_pick_ts',
    RQ_COMPLETE_TS  : 'complete_ts',
    RQ_PICKER_EMAIL : 'picker_email',
    RQ_PICKER_DISP  : 'picker_display',
    RQ_CYCLE_MIN    : 'cycle_time_min',
    RQ_STATUS       : 'status',
    RQ_NOTES        : 'notes',

    // _Identity
    ID_ALIAS        : 'alias',
    ID_CANONICAL    : 'canonical_email',
    ID_DISPLAY      : 'display_name',
    ID_DEPARTMENT   : 'department',
    ID_DROP_LOC     : 'drop_location',
    ID_ACTIVE       : 'active',

    // _Roles
    RO_EMAIL        : 'email',
    RO_ROLE         : 'role',
    RO_NOTES        : 'notes',
  },

  // ---------------------------------------------------------------------------------
  // Enums — use these constants everywhere, never bare strings
  // ---------------------------------------------------------------------------------
  ENUMS: {
    // Event types
    STATUS_SUPERSEDED : 'SUPERSEDED',  
    EVENT_SUPERSEDED  : 'SUPERSEDED',
    EVENT_REQUESTED : 'REQUESTED',
    EVENT_PICKED    : 'PICKED',
    EVENT_COMPLETED : 'COMPLETED',
    EVENT_DISCARDED : 'DISCARDED',

    // Intake sources
    INTAKE_FORM     : 'FORM',
    INTAKE_WALKIN   : 'WALK_IN',

    // Priority
    PRIORITY_NORMAL : 'NORMAL',
    PRIORITY_URGENT : 'URGENT',

    // Request status
    STATUS_OPEN      : 'OPEN',
    STATUS_COMPLETED : 'COMPLETED',
    STATUS_DISCARDED : 'DISCARDED',

    // Roles
    ROLE_MANAGER    : 'MANAGER',
    ROLE_PICKER     : 'PICKER',

    // Identity sentinel
    IDENTITY_UNMAPPED: 'UNMAPPED',

    // Sync flags written into Form Responses 1
    SYNC_PROCESSED  : 'PROCESSED',
    // Terminal marker for a response that cannot be synced as-is (no stock
    // number). Distinct from PROCESSED so the row is still visibly incomplete,
    // but it must be excluded from the "stuck" count — an unflagged row pins
    // Health._checkStuckResponses above zero forever, which turns the
    // heartbeat retry into an after-every-submission event.
    SYNC_SKIPPED    : 'SKIPPED_NO_STOCK',
  },

  // ---------------------------------------------------------------------------------
  // Shift / op-day model — THE single source of truth.
  // The operating day starts at START_HR (06:00): the overnight NIGHT shift
  // collapses into the op day of the DAY shift it followed.
  // DAY shift = [START_HR, SECOND_HR); NIGHT shift = the remainder (wraps
  // past midnight). Consumed by Turnover.js and ShiftReport.js — never
  // redeclare these hours in another module.
  // ---------------------------------------------------------------------------------
  SHIFT: {
    START_HR  : 6,   // op-day boundary AND day-shift start (06:00)
    SECOND_HR : 17,  // night-shift start (17:00)
  },

  // ---------------------------------------------------------------------------------
  // Turnover behavior
  // TURNOVER_SKIP_EMPTY_REPORT: when true, an AUTOMATED turnover run sends no
  // report email if the ended shift had zero requests, completions, and
  // discards AND zero open carryover (e.g., plant closed). If any OPEN work
  // crossed the boundary — a violation of the no-overlap rule — the report is
  // ALWAYS sent so the anomaly is seen. Manual runs always send.
  // ---------------------------------------------------------------------------------
  TURNOVER_SKIP_EMPTY_REPORT: true,

  // ---------------------------------------------------------------------------------
  // TURNOVER_PRUNE_ARCHIVE (rev 3.3): when true, the turnover trims its archive
  // copy down to the shift that actually ended, dropping rows belonging to the
  // shift that came on after the boundary. The trigger fires well after the
  // boundary (~17:54 and ~06:13 in practice), so without this a DAY snapshot
  // carries ~54 minutes of nightshift work under the DAY label.
  //
  // KILL SWITCH. Set false to fall back to rev-3.2 behavior — a whole-workbook
  // snapshot — without a redeploy. The fallback is over-inclusive, never lossy:
  // the extra rows also appear in their own shift's archive. Pruning is
  // best-effort regardless; a failure logs, emails ADMIN_EMAIL, and leaves the
  // untrimmed archive standing.
  //
  // Does NOT affect the clean slate. Survivor rules in _cleanSlate are
  // unconditional — this flag only governs what the archive keeps.
  // ---------------------------------------------------------------------------------
  TURNOVER_PRUNE_ARCHIVE: true,

  // ---------------------------------------------------------------------------------
  // Build Load Level — canonical values accepted from form + normalization map
  // ---------------------------------------------------------------------------------
  BUILD_LEVELS: ['L1', 'L2', 'L3', 'H1', 'H2', 'H3', 'Rework'],

  BUILD_LEVEL_MAP: {
    // Normalize common variants to canonical form
    'RW'    : 'Rework',
    'REWORK': 'Rework',
    'L1': 'L1', 'L2': 'L2', 'L3': 'L3',
    'H1': 'H1', 'H2': 'H2', 'H3': 'H3',
  },

  // ---------------------------------------------------------------------------------
  // Lot Status — the exact option set from the Google Form multiple-choice
  // ---------------------------------------------------------------------------------
  LOT_STATUSES: ['Parts Hold', 'Sublet Complete', 'Rework', 'Paint Parts', 'Other'],

  // Walk-In is injected by the system (not a form option) but is a valid value
  LOT_WALKIN: 'Walk In',

  // ---------------------------------------------------------------------------------
  // Bin taxonomy
  // Confirm / extend these with the owner before hard-blocking validation.
  // UNKNOWN location_type is accepted but flagged — never silently rejected.
  // ---------------------------------------------------------------------------------
  BIN: {
    // Delimiter used by pickers to enter multiple bins in one cell.
    // '/' is the documented separator; whitespace is also accepted because
    // pickers habitually type "B1.2 B2.3" (Day-2 beta request).
    DELIMITER_PATTERN: /[\/\s]+/,

    // Picker shorthand: B<n>.<n> denotes a Bulk bin — "B1.2" ≡ "BULK1.2"
    // (Day-2 beta request). Expanded at parse time so _Requests.bins,
    // _EventLog.bin_code, shift reports, and archive snapshots all carry the
    // full BULK form; what the picker actually typed is preserved in
    // _EventLog.raw_input. Anchored tight (B + digits + dot + digits) so it
    // cannot shadow rack codes, BULKHUB, or plain BULK entries.
    BULK_SHORTHAND_PATTERN     : /^B(\d+\.\d+)$/,
    BULK_SHORTHAND_REPLACEMENT : 'BULK$1',

    // Location type prefix map: prefix (uppercase) → location_type label.
    // Longest-match wins — URG must come before any shorter prefix it might shadow.
    // URG is a physical rack (shelves A-D, positions 1-5, spots a-f).
    // e.g. URGA4D, URGD5C, URGC4A — the full string is the bin code.
    LOCATION_PREFIXES: [
      { prefix: 'BULKHUB', type: 'BULK_HUB'  },
      { prefix: 'BULK',    type: 'BULK'       },
      { prefix: 'URG',     type: 'URG'        },
      { prefix: 'HUB',     type: 'HUB'        },
      { prefix: 'DASH',    type: 'DASH'       },
      { prefix: 'FL',      type: 'FLOOR'      },
    ],

    // General rack coordinate pattern: letter(s) then digits then letter(s) then digits, etc.
    // Kept loose — confirm exact format with real data and tighten if needed.
    RACK_PATTERN: /^[A-Z]\d+[A-Z]\d+$|^\d+[A-Z]\d+[A-Z]$/,

    // A token that is purely numeric and long is almost certainly a stock number
    // typed into the wrong cell. Reject and toast; produce no event.
    GARBAGE_PATTERN: /^\d{5,}$/,

    LOCATION_TYPE_UNKNOWN: 'UNKNOWN',
  },

  // ---------------------------------------------------------------------------------
  // Operational thresholds
  // ---------------------------------------------------------------------------------
  LOCK_TIMEOUT_MS         : 15000,  // LockService wait before giving up
  HEALTH_STUCK_MINUTES    : 3,      // unprocessed response age before alert (heartbeat fires at 2 min; hourly monitor is the backstop)
  PROCESSING_MIN_ROWS     : 50,     // minimum data rows kept on Processing board
  PROCESSING_STALE_WARN_HRS : 1,    // Processing CF: open-on-board age (hrs) → amber
  PROCESSING_STALE_CRIT_HRS : 3,    // Processing CF: open-on-board age (hrs) → red
  BACKUP_WEEK_OF_MONTH    : true,   // true = week-of-month; false = ISO week number

  // ---------------------------------------------------------------------------------
  // Backup folder name (relative to the spreadsheet's parent folder)
  // ---------------------------------------------------------------------------------
  BACKUP_FOLDER: 'Parts Request Backups',

  // ---------------------------------------------------------------------------------
  // Admin email — falls back to script owner (Session.getEffectiveUser()) if blank.
  // Override via Script Properties key 'ADMIN_EMAIL' to avoid hardcoding.
  // ---------------------------------------------------------------------------------
  get ADMIN_EMAIL() {
    const override = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    return override || Session.getEffectiveUser().getEmail();
  },
};
