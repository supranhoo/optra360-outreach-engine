/**
 * PROJECT ANTIGRAVITY: END-TO-END JOB OUTREACH INTELLIGENCE PLATFORM
 * Web App Controller & Client-Server API Bridge
 * (WebAppController.js)
 */

/**
 * Entry point for Google Apps Script Web App.
 * Serves the standalone responsive Single Page Application UI.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'syncAnkitProfile') {
    try {
      syncAnkitProfileNow();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: "Ankit Choudhary's profile, projects, and CV data synced to Sheet successfully!",
        spreadsheetUrl: ANTIGRAVITY_CONFIG.getSpreadsheetUrl()
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        error: err.message
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return HtmlService.createHtmlOutputFromFile('WebAppUI')
    .setTitle('⚡ Project Antigravity — Job Outreach Intelligence')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Helper to display the Web App inside Apps Script Editor or Google Sheets modal dialog.
 */
function openAntigravityUi() {
  const html = HtmlService.createHtmlOutputFromFile('WebAppUI')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, '⚡ Project Antigravity Command Center');
}

/**
 * API 1: Initial Dashboard Data
 */
function apiGetOverview() {
  try {
    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    const todayStr = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd');
    const todayKey = 'SENT_COUNT_' + todayStr;
    const sentToday = Number(props.getProperty(todayKey) || 0);

    let remainingQuota = 1500;
    try {
      remainingQuota = MailApp.getRemainingDailyQuota();
    } catch (e) {}

    let totalSent = 0;
    let totalReplies = 0;
    let positiveReplies = 0;
    let interviewsCount = 0;
    let draftCount = 0;
    let queuedCount = 0;

    let recentDispatches = [];
    let activeDeals = [];

    if (ss) {
      const outreachSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OUTREACH_HISTORY);
      const respSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.RESPONSES);
      const pipeSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.INTERVIEW_PIPELINE);

      if (outreachSheet) {
        const oData = outreachSheet.getDataRange().getValues().slice(1);
        totalSent = oData.filter(r => r[11] === 'SENT').length;
        draftCount = oData.filter(r => r[11] === 'DRAFT_READY').length;
        queuedCount = oData.filter(r => r[11] === 'APPROVED_QUEUED').length;

        recentDispatches = oData.slice(-8).reverse().map(r => ({
          outreachId: r[0],
          campaignId: r[1],
          contactId: r[2],
          subject: r[5],
          scheduledAt: r[7],
          sentAt: r[8],
          status: r[11]
        }));
      }

      if (respSheet) {
        const rData = respSheet.getDataRange().getValues().slice(1);
        totalReplies = rData.length;
        positiveReplies = rData.filter(r => r[6] === 'POSITIVE_INTERVIEW' || r[6] === 'REFERRAL_FORWARD').length;
      }

      if (pipeSheet) {
        const pData = pipeSheet.getDataRange().getValues().slice(1);
        interviewsCount = pData.length;
        activeDeals = pData.map(r => ({
          pipelineId: r[0],
          companyId: r[1],
          opportunityId: r[2],
          contactId: r[3],
          stage: r[4],
          nextMeeting: r[5],
          prepDocUrl: r[6],
          compTarget: r[7],
          winProb: r[8]
        }));
      }
    }

    const replyRate = totalSent > 0 ? ((totalReplies / totalSent) * 100).toFixed(1) + '%' : '0.0%';
    const posReplyRate = totalSent > 0 ? ((positiveReplies / totalSent) * 100).toFixed(1) + '%' : '0.0%';

    return {
      success: true,
      userEmail: ANTIGRAVITY_CONFIG.getCandidateEmail(),
      spreadsheetUrl: ANTIGRAVITY_CONFIG.getSpreadsheetUrl(),
      spreadsheetId: props.getProperty('SPREADSHEET_ID') || '',
      quota: {
        sentToday: sentToday,
        dailyHardLimit: ANTIGRAVITY_CONFIG.DAILY_HARD_LIMIT,
        remainingGoogleQuota: remainingQuota,
        draftCount: draftCount,
        queuedCount: queuedCount
      },
      kpis: {
        totalSent: totalSent,
        totalReplies: totalReplies,
        positiveReplies: positiveReplies,
        replyRate: replyRate,
        positiveReplyRate: posReplyRate,
        interviewsCount: interviewsCount
      },
      recentDispatches: recentDispatches,
      activeDeals: activeDeals
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * API 2: Contacts Data Table
 */
function apiGetContacts() {
  try {
    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    if (!ss) return { success: false, error: 'Database not initialized.' };

    const contactSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.CONTACTS);
    if (!contactSheet) return { success: false, error: 'Contacts sheet not found.' };

    const data = contactSheet.getDataRange().getValues().slice(1);
    const contacts = data.map(r => ({
      contactId: r[0],
      companyId: r[1],
      opportunityId: r[2],
      name: (r[3] || '') + ' ' + (r[4] || ''),
      email: r[5],
      verification: r[6],
      persona: r[7],
      role: r[8],
      leadScore: r[11],
      status: r[12]
    }));

    return { success: true, contacts: contacts };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * API 3: Outreach Drafts & Queue
 */
function apiGetOutreachQueue() {
  try {
    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    if (!ss) return { success: false, error: 'Database not initialized.' };

    const sheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OUTREACH_HISTORY);
    if (!sheet) return { success: false, error: 'Outreach_History not found.' };

    const data = sheet.getDataRange().getValues().slice(1);
    const drafts = data.map((r, idx) => ({
      rowIndex: idx + 2,
      outreachId: r[0],
      campaignId: r[1],
      contactId: r[2],
      oppId: r[3],
      projectId: r[4],
      subject: r[5],
      htmlBody: r[6],
      scheduledSendAt: r[7],
      actualSentAt: r[8],
      threadId: r[9],
      status: r[11]
    }));

    return { success: true, drafts: drafts };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * API 4: Action triggers from UI
 */
function apiApproveOutreachRow(outreachId) {
  try {
    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    const sheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OUTREACH_HISTORY);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === outreachId) {
        sheet.getRange(i + 1, 12).setValue('APPROVED_QUEUED');
        return { success: true, outreachId: outreachId, status: 'APPROVED_QUEUED' };
      }
    }
    return { success: false, error: 'Outreach ID not found' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiApproveAllDrafts() {
  try {
    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    const sheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OUTREACH_HISTORY);
    const data = sheet.getDataRange().getValues();
    let count = 0;

    for (let i = 1; i < data.length; i++) {
      if (data[i][11] === 'DRAFT_READY') {
        sheet.getRange(i + 1, 12).setValue('APPROVED_QUEUED');
        count++;
      }
    }
    return { success: true, approvedCount: count };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiRunLeadScoring() {
  try {
    processAndScoreNewContacts();
    return { success: true, message: 'Lead scoring and MX verification completed.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiGenerateDrafts(campaignId) {
  try {
    generateCampaignOutreachDrafts(campaignId || 'CMPG-01');
    return { success: true, message: 'Draft generation completed.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiRunDispatcher() {
  try {
    runQuotaSafeMailDispatcher();
    return { success: true, message: 'Mail dispatcher cycle completed.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiSyncReplies() {
  try {
    syncGmailResponsesAndKPIs();
    return { success: true, message: 'Reply sync and follow-up engine completed.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiBootstrapDb() {
  try {
    bootstrapAntigravityArchitecture();
    return { success: true, spreadsheetUrl: ANTIGRAVITY_CONFIG.getSpreadsheetUrl() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiSyncAnkitProfile() {
  try {
    syncAnkitProfileNow();
    return {
      success: true,
      message: "Ankit Choudhary's CV, projects, and career records synced to Sheet successfully!",
      spreadsheetUrl: ANTIGRAVITY_CONFIG.getSpreadsheetUrl()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiInstallTriggers() {
  try {
    installPlatformTriggers();
    return { success: true, message: 'Automation triggers installed successfully.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiUninstallTriggers() {
  try {
    uninstallPlatformTriggers();
    return { success: true, message: 'Automation triggers removed.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiSaveSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (settings.geminiApiKey) {
      props.setProperty('GEMINI_API_KEY', settings.geminiApiKey.trim());
    }
    if (settings.spreadsheetId) {
      props.setProperty('SPREADSHEET_ID', settings.spreadsheetId.trim());
    }
    return { success: true, message: 'Settings saved successfully.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiGetSettings() {
  try {
    const props = PropertiesService.getScriptProperties();
    const apiKey = props.getProperty('GEMINI_API_KEY') || '';
    const maskedKey = apiKey ? (apiKey.slice(0, 4) + '••••••••' + apiKey.slice(-4)) : '';
    const sheetId = props.getProperty('SPREADSHEET_ID') || '';
    return {
      success: true,
      hasApiKey: !!apiKey,
      maskedApiKey: maskedKey,
      spreadsheetId: sheetId,
      spreadsheetUrl: ANTIGRAVITY_CONFIG.getSpreadsheetUrl()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}
