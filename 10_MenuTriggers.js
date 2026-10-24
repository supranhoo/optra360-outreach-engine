/**
 * Optra360 & Zoptro Automated Outreach & Opportunity Engine
 * 
 * 10_MenuTriggers.js: Spreadsheet onOpen UI menu, automatic 08:00 AM IST triggers & sheet bootstrapping.
 */

/**
 * Custom spreadsheet onOpen trigger adding the Optra360 Engine menu.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Optra360 Engine')
    .addItem('🚀 Open Command Sidebar', 'showOptraSidebar')
    .addSeparator()
    .addItem('▶ Run Full Outreach Pipeline (Steps 1-5)', 'runFullDailyPipeline')
    .addItem('🛠 Seed 7-Tab CRM Architecture & Demos', 'setupInitialSheetArchitecture')
    .addSeparator()
    .addItem('⏰ Set Up Daily 08:00 AM IST Trigger', 'installDailyTrigger')
    .addItem('⏹ Remove Daily Triggers', 'removeDailyTriggers')
    .addToUi();
}

/**
 * Runs the full end-to-end outreach pipeline sequentially.
 */
function runFullDailyPipeline() {
  console.log('Starting Optra360 daily outreach pipeline...');

  // Step 1: Research
  var res1 = runDailyResearch();
  console.log('Step 1 (Research): ' + JSON.stringify(res1));

  // Step 2: Contact Enrichment & MX verification
  var res2 = enrichProspectContacts();
  console.log('Step 2 (Enrichment): ' + JSON.stringify(res2));

  // Step 3: AI Problem Diagnosis
  var res3 = mapProblemSolutions();
  console.log('Step 3 (Diagnosis): ' + JSON.stringify(res3));

  // Step 4: 2-Tier PDF Pitch Deck generation
  var res4 = buildAllPendingPitchDecks();
  console.log('Step 4 (Pitch Decks): ' + JSON.stringify(res4));

  // Step 5: Email Drafts & LinkedIn queue
  var res5 = processOutreachEmailBatch();
  console.log('Step 5 (Drafts): ' + JSON.stringify(res5));

  return {
    research: res1,
    enrichment: res2,
    diagnosis: res3,
    decks: res4,
    drafts: res5
  };
}

/**
 * Installs scheduled trigger at 08:00 AM IST every day.
 */
function installDailyTrigger() {
  removeDailyTriggers();

  ScriptApp.newTrigger('runFullDailyPipeline')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone('Asia/Kolkata')
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast('Daily 08:00 AM IST outreach trigger successfully active!', 'Optra360 Engine', 5);
}

/**
 * Removes existing time-driven triggers for this script.
 */
function removeDailyTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runFullDailyPipeline') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Bootstraps all 7 CRM tabs with headers and styling if not already present.
 */
function setupInitialSheetArchitecture() {
  var ss = getCRMSpreadsheet();
  var c = OPTRA_CONFIG.BRAND.COLORS;

  var schema = [
    {
      name: OPTRA_CONFIG.TABS.PROFILES,
      headers: ['Profile_ID', 'Mode', 'Name', 'Title', 'Email', 'Phone', 'LinkedIn_URL', 'Headline', 'Pitch_Focus']
    },
    {
      name: OPTRA_CONFIG.TABS.TARGET_COMPANIES,
      headers: ['Company_ID', 'Company_Name', 'Industry_Sector', 'Est_Headcount', 'Domain_Website', 'Identified_Trigger', 'Source_Signal', 'Signal_URL', 'Sourcing_Date', 'Status']
    },
    {
      name: OPTRA_CONFIG.TABS.PROBLEM_SOLUTION,
      headers: ['Company_ID', 'Company_Name', 'Identified_Trigger', 'Raw_Problem_Signal', 'Formulated_Pain_Statement', 'Recommended_Solution', 'Mapped_Delivery_Stage', 'Recommended_Engagement', 'Matched_Demo_ID', 'Expected_Transformation_Impact', 'Diagnosis_Status']
    },
    {
      name: OPTRA_CONFIG.TABS.CONTACTS,
      headers: ['Contact_ID', 'Company_ID', 'Company_Name', 'Prospect_Name', 'Designation', 'Email', 'Email_Status', 'LinkedIn_URL', 'Identified_Trigger', 'Outreach_Status', 'Deck_5_Slide_PDF', 'Deck_12_Slide_PDF', 'Deck_5_Slide_ID', 'Deck_12_Slide_ID', 'LinkedIn_Note', 'Draft_Date']
    },
    {
      name: OPTRA_CONFIG.TABS.DECK_TEMPLATES,
      headers: ['Template_ID', 'Deck_Type', 'Google_Slides_ID', 'Target_Audience', 'Merge_Tags_List', 'Status']
    },
    {
      name: OPTRA_CONFIG.TABS.DEMOS,
      headers: ['Demo_ID', 'Title', 'Category', 'Operating_Scale', 'Live_URL', 'Description', 'Trigger_Alignment']
    },
    {
      name: OPTRA_CONFIG.TABS.KPI,
      headers: ['Metric_ID', 'Metric_Name', 'Target_Daily', 'Current_Count', 'Conversion_Rate', 'Last_Updated']
    }
  ];

  for (var i = 0; i < schema.length; i++) {
    var def = schema[i];
    var sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(def.headers);
      var headerRange = sheet.getRange(1, 1, 1, def.headers.length);
      headerRange.setBackground(c.AUTHORITY_BLUE)
                 .setFontColor(c.CLARITY_IVORY)
                 .setFontWeight('bold')
                 .setFontFamily('Figtree');
      sheet.setFrozenRows(1);
    }
  }

  // Preload default profiles & demos if empty
  var profileData = getSheetDataAsObjects(OPTRA_CONFIG.TABS.PROFILES);
  if (profileData.length === 0) {
    var defaults = getAllProfiles();
    for (var p = 0; p < defaults.length; p++) {
      appendSheetRow(OPTRA_CONFIG.TABS.PROFILES, defaults[p]);
    }
  }

  seedDemoPortfolioSheet();
  SpreadsheetApp.getActiveSpreadsheet().toast('CRM architecture and initial records initialized!', 'Optra360 Engine', 5);
}
