/**
 * Optra360 & Zoptro Automated Outreach & Opportunity Engine
 * 
 * 09_SidebarController.js: Bridge between HTML sidebar and server-side Apps Script functions.
 */

/**
 * Displays the Optra360 Command Sidebar in the active Google Spreadsheet.
 */
function showOptraSidebar() {
  var ui = SpreadsheetApp.getUi();
  var htmlOutput;

  try {
    htmlOutput = HtmlService.createHtmlOutputFromFile('08_SidebarUI')
      .setTitle('Optra360 Engine')
      .setWidth(340);
  } catch (e) {
    // Fallback if file lookup fails
    htmlOutput = HtmlService.createHtmlOutput(getOptraSidebarHtml_())
      .setTitle('Optra360 Engine')
      .setWidth(340);
  }

  ui.showSidebar(htmlOutput);
}

/**
 * Returns state object for the sidebar dashboard UI.
 * @return {Object}
 */
function getSidebarDashboardState() {
  var kpis = getDashboardKPIs();
  var contacts = getSheetDataAsObjects(OPTRA_CONFIG.TABS.CONTACTS);

  var recentDrafts = [];
  for (var i = contacts.length - 1; i >= 0; i--) {
    var c = contacts[i];
    if (c.Outreach_Status === 'Drafted' && c.LinkedIn_Note) {
      recentDrafts.push(c);
      if (recentDrafts.length >= 5) break;
    }
  }

  return {
    totalCompanies: kpis.totalCompanies,
    verifiedEmails: kpis.verifiedEmails,
    draftsPrepared: kpis.draftsPrepared,
    emailsSent: kpis.emailsSent,
    activeProfile: kpis.activeProfile,
    recentDrafts: recentDrafts
  };
}

/**
 * Saves user credentials in ScriptProperties.
 * @param {Object} keys 
 */
function saveSidebarApiKeys(keys) {
  var props = PropertiesService.getScriptProperties();
  if (keys.gemini) props.setProperty(OPTRA_CONFIG.PROPERTY_KEYS.GEMINI_KEY, keys.gemini);
  if (keys.hunter) props.setProperty(OPTRA_CONFIG.PROPERTY_KEYS.HUNTER_KEY, keys.hunter);
  if (keys.apollo) props.setProperty(OPTRA_CONFIG.PROPERTY_KEYS.APOLLO_KEY, keys.apollo);
  return { success: true };
}

/**
 * Fallback embedded HTML generator for self-contained execution.
 * @private
 */
function getOptraSidebarHtml_() {
  return HtmlService.createHtmlOutputFromFile('08_SidebarUI').getContent();
}
