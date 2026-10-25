/**
 * PROJECT ANTIGRAVITY: END-TO-END JOB OUTREACH INTELLIGENCE PLATFORM
 * Module 10: Trigger Management, Cron Automation & Daily Quota Maintenance
 * (10_TriggersAndAutomations.js)
 */

/**
 * Programmatically installs the 15-minute and nightly time-driven automation triggers.
 */
function installPlatformTriggers() {
  // Remove existing triggers first to ensure idempotency
  uninstallPlatformTriggers();

  // 1. Quota-safe mail dispatcher (every 15 minutes)
  ScriptApp.newTrigger('runQuotaSafeMailDispatcher')
    .timeBased()
    .everyMinutes(15)
    .create();

  // 2. Inbound reply, bounce detection and follow-up engine (every 15 minutes)
  ScriptApp.newTrigger('syncGmailResponsesAndKPIs')
    .timeBased()
    .everyMinutes(15)
    .create();

  // 3. Nightly KPI rollup & maintenance (every day at 23:45)
  ScriptApp.newTrigger('updateDailyKpiMetrics_')
    .timeBased()
    .atHour(23)
    .nearMinute(45)
    .everyDays(1)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Automation triggers successfully installed! Mail dispatcher & reply sync will run every 15 minutes.',
    'Triggers Configured',
    6
  );
}

/**
 * Deletes all project triggers to safely pause or reset automation.
 */
function uninstallPlatformTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  triggers.forEach(t => {
    const handler = t.getHandlerFunction();
    if (['runQuotaSafeMailDispatcher', 'syncGmailResponsesAndKPIs', 'updateDailyKpiMetrics_'].includes(handler)) {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Removed ' + count + ' active Antigravity automation triggers.',
    'Triggers Cleaned',
    4
  );
}

/**
 * Daily maintenance utility: Resets today's sent counter in PropertiesService.
 */
function resetDailyQuotaCounter() {
  const props = PropertiesService.getScriptProperties();
  const todayStr = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd');
  props.deleteProperty('SENT_COUNT_' + todayStr);
  SpreadsheetApp.getActiveSpreadsheet().toast('Daily sent counter reset to 0.', 'Success', 3);
}
