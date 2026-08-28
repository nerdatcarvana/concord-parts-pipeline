/**
 * Health.gs
 * Monitors system health and emails the owner when action is needed.
 *
 * Two layers, deliberately unequal:
 *   PRIMARY — heartbeat: onFormSubmitHeartbeat(e) fires on every form submit
 *   alongside the sync trigger and schedules a one-time check 2 minutes later
 *   to verify the row actually got processed. If it didn't (trigger silently
 *   failed), it auto-retries the sync and emails the owner. This catches the
 *   real failure mode within minutes of it happening.
 *
 *   BACKSTOP — monitor(): time-driven HOURLY sweep for anything the heartbeat
 *   can't see (a stuck row whose heartbeat execution itself failed, plus the
 *   unmapped-alias backlog). This used to run every 5 minutes — 288
 *   executions/day of sheet reads guarding a failure mode the heartbeat
 *   already covers within 2 minutes. Hourly keeps the safety net while
 *   cutting the standing execution load ~96%. If you ever need it tighter,
 *   change ONE line in Setup._installTriggers (everyHours → everyMinutes)
 *   and re-run Setup.
 *
 * Checks (monitor):
 *   1. Stuck Form Responses — unprocessed rows older than CONFIG.HEALTH_STUCK_MINUTES
 *   2. Unmapped alias backlog — aliases that couldn't be resolved to a person
 *
 * Public API:
 *   Health.monitor()             → full periodic check (time-driven, hourly)
 *   Health.heartbeatCheck()      → lightweight stuck-row retry (one-time trigger)
 */

const Health = (() => {

  // --------------------------------------------------------------------------
  // monitor()
  // Time-driven, hourly. Checks for stuck rows and unmapped aliases.
  // Uses Session.getEffectiveUser() — getActiveUser() returns '' in time-driven.
  // --------------------------------------------------------------------------
  function monitor() {
    const issues = [];
    const ss     = SpreadsheetApp.getActiveSpreadsheet();

    // Proactively restore any drifted managed tab names (resolved by gid) so a
    // rename never lingers, even in the gaps between picker edits.
    Sheets.repairAll(ss);

    const stuckCount = _checkStuckResponses(ss);
    if (stuckCount > 0) {
      // Auto-retry before alerting — the trigger may have just had a transient failure
      try {
        Intake.syncFormResponses();
      } catch (err) {
        console.error('Health.monitor: auto-retry sync failed:', err.message);
      }

      // Re-check after retry
      const stillStuck = _checkStuckResponses(ss);
      if (stillStuck > 0) {
        issues.push(
          `⚠️ ${stillStuck} form submission(s) have been stuck unprocessed for ` +
          `more than ${CONFIG.HEALTH_STUCK_MINUTES} minutes. An auto-retry was attempted and failed.`
        );
      }
    }

    const unmappedAliases = Identity.listUnmapped();
    if (unmappedAliases.length > 0) {
      issues.push(
        `👤 ${unmappedAliases.length} unrecognized requester alias(es) in the backlog:\n` +
        unmappedAliases.map(a => `  • ${a}`).join('\n') + '\n' +
        `Add these to the _Identity sheet (alias → canonical_email) to resolve.`
      );
    }

    if (issues.length === 0) return;
    _sendAlert(ss, issues);
  }

  // --------------------------------------------------------------------------
  // heartbeatCheck()
  // Fired by a one-time trigger scheduled 2 minutes after each form submission
  // by onFormSubmitHeartbeat(). Verifies the triggering row got processed.
  // If it didn't, retries the sync and emails the owner.
  //
  // Cleans up its own one-time trigger after running.
  // --------------------------------------------------------------------------
  function heartbeatCheck() {
    // Clean up the one-time trigger that called this
    _deleteSelfTrigger('Health_heartbeatCheck');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Age-agnostic (0): the heartbeat fires ~2 min after submit, well under
    // HEALTH_STUCK_MINUTES, so an age-gated check would never see the row it was
    // scheduled to verify. Any still-unprocessed row means the submit-sync didn't land.
    const stuckCount = _checkStuckResponses(ss, 0);
    if (stuckCount === 0) return; // all good

    // Auto-retry (idempotent — PROCESSED flags prevent re-emission)
    try {
      Intake.syncFormResponses();
    } catch (err) {
      console.error('Health.heartbeatCheck: sync retry failed:', err.message);
    }

    // Check again after retry
    const stillStuck = _checkStuckResponses(ss, 0);
    if (stillStuck > 0) {
      _sendAlert(ss, [
        `⚠️ ${stillStuck} form submission(s) failed to sync automatically after a form submit. ` +
        `An auto-retry was attempted.\n` +
        `Please open the sheet and run "Sync Form Submissions" from the Parts Management menu.`
      ]);
    }
  }

  // --------------------------------------------------------------------------
  // _checkStuckResponses(ss) → count of unprocessed rows older than threshold
  // --------------------------------------------------------------------------
  function _checkStuckResponses(ss, minAgeMinutes) {
    const respSheet = Sheets.getByName(CONFIG.SHEETS.FORM_RESPONSES, ss);
    if (!respSheet || respSheet.getLastRow() < 2) return 0;

    const data    = respSheet.getDataRange().getValues();
    const headers = data[0].map(Utils.cleanHeader);
    const tsIdx   = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.FR_TIMESTAMP));
    const syncIdx = headers.indexOf(Utils.cleanHeader(CONFIG.HEADERS.FR_SYNC_STATUS));

    if (tsIdx === -1) return 0;

    const now         = new Date();
    // Age gate in minutes; callers may pass 0 to count ANY unprocessed row.
    const ageGate     = (minAgeMinutes === undefined) ? CONFIG.HEALTH_STUCK_MINUTES : minAgeMinutes;
    const thresholdMs = ageGate * 60 * 1000;
    let stuckCount    = 0;

    for (let i = 1; i < data.length; i++) {
      const row        = data[i];
      const syncStatus = syncIdx !== -1 ? String(row[syncIdx]).trim() : '';
      if (syncStatus === CONFIG.ENUMS.SYNC_PROCESSED) continue;
      // A row Intake could not sync (no stock number) is not "stuck" — it is
      // terminal. Counting it would keep this function permanently above zero,
      // firing the retry sync after every submission and alerting hourly.
      if (syncStatus === CONFIG.ENUMS.SYNC_SKIPPED)   continue;

      const rawTs = row[tsIdx];
      if (!rawTs) continue;

      const submitted = rawTs instanceof Date ? rawTs : new Date(rawTs);
      if (isNaN(submitted.getTime())) continue;

      if ((now - submitted) >= thresholdMs) stuckCount++;
    }

    return stuckCount;
  }

  // --------------------------------------------------------------------------
  // _sendAlert(ss, issues)
  // --------------------------------------------------------------------------
  function _sendAlert(ss, issues) {
    const adminEmail = CONFIG.ADMIN_EMAIL;
    const subject    = `🚨 Parts Workflow Alert — ${issues.length} issue(s) need attention`;

    const body = [
      'Parts Workflow Health Alert',
      '----------------------------',
      '',
      ...issues.map((issue, i) => `${i + 1}. ${issue}`),
      '',
      '----------------------------',
      'Action required:',
      '  • For stuck responses: open the sheet and use Parts Management → "Sync Form Submissions".',
      '  • For unmapped aliases: add the alias(es) to the _Identity sheet.',
      '',
      `Sheet: ${ss.getUrl()}`,
    ].join('\n');

    MailApp.sendEmail({ to: adminEmail, subject, body });
  }

  // --------------------------------------------------------------------------
  // _deleteSelfTrigger(fnName)
  // Removes all project triggers for the given function name.
  // Used by heartbeatCheck to clean up its own one-time trigger.
  // --------------------------------------------------------------------------
  function _deleteSelfTrigger(fnName) {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === fnName) ScriptApp.deleteTrigger(t);
    });
  }

  // --------------------------------------------------------------------------
  // Public surface
  // --------------------------------------------------------------------------
  return { monitor, heartbeatCheck };

})();

// Top-level stubs
function Health_monitor()        { Health.monitor(); }
function Health_heartbeatCheck() { Health.heartbeatCheck(); }

// ---------------------------------------------------------------------------
// onFormSubmitHeartbeat(e)
// Fires on every form submit (installed alongside onFormSubmitTrigger).
// Schedules a one-time trigger 2 minutes out to verify the sync actually ran.
// This is the safety net for silent trigger failures.
// ---------------------------------------------------------------------------
function onFormSubmitHeartbeat(e) {
  // Debounce: one pending heartbeat suffices. Creating a fresh one-time trigger per
  // submission would exhaust the per-user time-trigger quota under a burst and start
  // throwing on .create(). heartbeatCheck deletes its own trigger when it runs, so
  // the next submission after that re-arms it.
  const pending = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'Health_heartbeatCheck');
  if (pending) return;

  // Schedule a one-time check 2 minutes from now
  ScriptApp.newTrigger('Health_heartbeatCheck')
    .timeBased()
    .after(2 * 60 * 1000) // 2 minutes in milliseconds
    .create();
}
