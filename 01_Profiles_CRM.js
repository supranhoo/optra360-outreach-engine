/**
 * Optra360 & Zoptro Automated Outreach & Opportunity Engine
 * 
 * 01_Profiles_CRM.js: Relational CRM reader/writer and Dual-Mode Profile Engine.
 */

/**
 * Access the master CRM spreadsheet.
 * @return {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function getCRMSpreadsheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && ss.getId() === OPTRA_CONFIG.SPREADSHEET_ID) {
      return ss;
    }
    return SpreadsheetApp.openById(OPTRA_CONFIG.SPREADSHEET_ID);
  } catch (e) {
    console.warn('Could not open by hardcoded ID, falling back to active spreadsheet: ' + e.message);
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

/**
 * Converts a sheet's tabular data into an array of objects based on header row (row 1).
 * @param {string} sheetName 
 * @return {Array<Object>}
 */
function getSheetDataAsObjects(sheetName) {
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var results = [];

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var obj = { _rowIndex: r + 1 };
    var hasContent = false;

    for (var c = 0; c < headers.length; c++) {
      var key = headers[c];
      if (key) {
        obj[key] = row[c];
        if (row[c] !== '' && row[c] !== null && row[c] !== undefined) {
          hasContent = true;
        }
      }
    }
    if (hasContent) {
      results.push(obj);
    }
  }
  return results;
}

/**
 * Updates an existing row in a sheet matching a specific identifier column value.
 * @param {string} sheetName
 * @param {string} idColName
 * @param {*} idVal
 * @param {Object} updatesObj
 * @return {boolean}
 */
function updateSheetRow(sheetName, idColName, idVal, updatesObj) {
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return false;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return false;

  var headers = data[0].map(function(h) { return String(h).trim(); });
  var idColIdx = headers.indexOf(idColName);
  if (idColIdx === -1) return false;

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idColIdx]).trim() === String(idVal).trim()) {
      var rowNum = r + 1;
      for (var key in updatesObj) {
        var colIdx = headers.indexOf(key);
        if (colIdx !== -1) {
          sheet.getRange(rowNum, colIdx + 1).setValue(updatesObj[key]);
        }
      }
      return true;
    }
  }
  return false;
}

/**
 * Appends a new record to the sheet matching the header columns.
 * @param {string} sheetName
 * @param {Object} dataObj
 * @return {number} row number appended
 */
function appendSheetRow(sheetName, dataObj) {
  var ss = getCRMSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return -1;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { return String(h).trim(); });
  var row = [];

  for (var i = 0; i < headers.length; i++) {
    var val = dataObj[headers[i]];
    row.push(val !== undefined ? val : '');
  }

  sheet.appendRow(row);
  return sheet.getLastRow();
}

/**
 * Retrieves all configured profiles (Mode A B2B co-founders & Mode B Executive Job Seekers).
 * @return {Array<Object>}
 */
function getAllProfiles() {
  var profiles = getSheetDataAsObjects(OPTRA_CONFIG.TABS.PROFILES);
  if (profiles.length === 0) {
    // Default fallback profiles if sheet is blank
    return [
      {
        Profile_ID: 'PROF-01',
        Mode: 'Mode B (Executive Career & Advisory)',
        Name: 'Ankit Choudhary',
        Title: 'Co-founder & CTO | Head of HR & HRIS',
        Email: 'ankit.choudhary@optra360.com',
        Phone: '+91 99999 99999',
        LinkedIn_URL: 'https://linkedin.com/in/ankitchoudhary',
        Headline: 'Transforming HR, Industrial Engineering & HR Tech Architectures',
        Pitch_Focus: 'Executive HR, CTO, HR Tech Leadership & Strategic Advisory'
      },
      {
        Profile_ID: 'PROF-02',
        Mode: 'Mode A (B2B Optra360 Diagnostic Outreach)',
        Name: 'Jaspal Bhanker',
        Title: 'Co-founder & Practice Head (Operations)',
        Email: 'jaspal.bhanker@optra360.com',
        Phone: '',
        LinkedIn_URL: '',
        Headline: 'Enterprise Operations & 6-Stage Business Transformation',
        Pitch_Focus: 'Process Breakdown & Performance Stagnation Turnarounds'
      },
      {
        Profile_ID: 'PROF-03',
        Mode: 'Mode A (B2B Optra360 Diagnostic Outreach)',
        Name: 'Divyanee Gohite',
        Title: 'Co-founder & Business Architect',
        Email: 'divyanee.gohite@optra360.com',
        Phone: '',
        LinkedIn_URL: '',
        Headline: 'Organizational Design & Workforce Strategy',
        Pitch_Focus: 'Leadership Disconnect & Workforce Misalignment'
      },
      {
        Profile_ID: 'PROF-04',
        Mode: 'Mode A (B2B Optra360 Diagnostic Outreach)',
        Name: 'Sunehri Yadav',
        Title: 'Co-founder & Program Director',
        Email: 'sunehri.yadav@optra360.com',
        Phone: '',
        LinkedIn_URL: '',
        Headline: 'Digital Systems & Operations Governance',
        Pitch_Focus: 'System Fragmentation & Rapid Modernization'
      },
      {
        Profile_ID: 'PROF-05',
        Mode: 'Mode A (B2B Optra360 Diagnostic Outreach)',
        Name: 'Vivek Kumar Dansena',
        Title: 'Co-founder & Execution Lead',
        Email: 'vivek.dansena@optra360.com',
        Phone: '',
        LinkedIn_URL: '',
        Headline: 'Manufacturing & Heavy Industry Operational Excellence',
        Pitch_Focus: 'Shop-Floor Digitization & Controlled Transformation Pilots'
      }
    ];
  }
  return profiles;
}

/**
 * Returns the currently active profile.
 * @return {Object}
 */
function getActiveProfile() {
  var props = PropertiesService.getScriptProperties();
  var activeId = props.getProperty(OPTRA_CONFIG.PROPERTY_KEYS.ACTIVE_PROFILE) || 'PROF-01';
  var all = getAllProfiles();

  for (var i = 0; i < all.length; i++) {
    if (all[i].Profile_ID === activeId) {
      return all[i];
    }
  }
  return all[0];
}

/**
 * Sets the active profile ID.
 * @param {string} profileId 
 */
function setActiveProfileMode(profileId) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(OPTRA_CONFIG.PROPERTY_KEYS.ACTIVE_PROFILE, profileId);
  return getActiveProfile();
}

/**
 * Computes live KPI metrics across the pipeline.
 * @return {Object}
 */
function getDashboardKPIs() {
  var targets = getSheetDataAsObjects(OPTRA_CONFIG.TABS.TARGET_COMPANIES);
  var contacts = getSheetDataAsObjects(OPTRA_CONFIG.TABS.CONTACTS);

  var totalCompanies = targets.length;
  var totalContacts = contacts.length;
  var verifiedEmails = 0;
  var draftsPrepared = 0;
  var emailsSent = 0;
  var meetingsBooked = 0;

  for (var i = 0; i < contacts.length; i++) {
    var c = contacts[i];
    if (c.Verification_Status === 'Valid' || c.Email_Status === 'Verified') verifiedEmails++;
    if (c.Outreach_Status === 'Drafted') draftsPrepared++;
    if (c.Outreach_Status === 'Sent' || c.Outreach_Status === 'Followed Up') emailsSent++;
    if (c.Outreach_Status === 'Converted' || c.Outreach_Status === 'Meeting Scheduled') meetingsBooked++;
  }

  return {
    totalCompanies: totalCompanies,
    totalContacts: totalContacts,
    verifiedEmails: verifiedEmails,
    draftsPrepared: draftsPrepared,
    emailsSent: emailsSent,
    meetingsBooked: meetingsBooked,
    activeProfile: getActiveProfile()
  };
}
