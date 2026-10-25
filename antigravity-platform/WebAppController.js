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
 * Returns clean flat CRM data specifically tailored for the Simple CRM Pipeline:
 * To whom we have shared what and if we have received any response from them.
 */
function apiGetSimpleCrmData() {
  try {
    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    if (!ss) return { success: false, error: 'Database not initialized.' };

    const contactSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.CONTACTS);
    const compSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.TARGET_COMPANIES);
    const outreachSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OUTREACH_HISTORY);
    const respSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.RESPONSES);

    if (!contactSheet || !outreachSheet) {
      return { success: false, error: 'Required CRM sheets missing.' };
    }

    const contacts = contactSheet.getDataRange().getValues().slice(1);
    const companies = compSheet ? compSheet.getDataRange().getValues().slice(1) : [];
    const outreaches = outreachSheet.getDataRange().getValues().slice(1);
    const responses = respSheet ? respSheet.getDataRange().getValues().slice(1) : [];

    // Index companies by companyId
    const compMap = {};
    companies.forEach(c => {
      compMap[String(c[0])] = {
        name: c[1] || '',
        domain: c[2] || '',
        industry: c[3] || '',
        headcount: c[5] || c[4] || '500-1,000 workforce',
        location: c[6] || 'India'
      };
    });

    // Index contacts by contactId
    const contactMap = {};
    contacts.forEach(c => {
      contactMap[String(c[0])] = {
        contactId: c[0],
        companyId: c[1],
        name: ((c[3] || '') + ' ' + (c[4] || '')).trim() || 'Executive Leader',
        email: c[5] || '',
        role: c[8] || 'Decision Maker',
        persona: c[7] || 'FOUNDER',
        leadScore: c[11] || 85,
        status: c[12] || 'DISCOVERED'
      };
    });

    // Index responses by outreachId or contactId
    const respMap = {};
    responses.forEach(r => {
      const outId = String(r[1] || '');
      const cId = String(r[2] || '');
      const data = {
        responseId: r[0],
        receivedAt: r[4] ? new Date(r[4]).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }) : '',
        snippet: r[5] || '',
        sentiment: r[6] || 'AWAITING',
        hoursToReply: r[7]
      };
      if (outId) respMap[outId] = data;
      if (cId) respMap[cId] = data;
    });

    // Build Unified CRM items
    const crmRows = [];

    outreaches.forEach((o, idx) => {
      const outId = String(o[0]);
      const contactId = String(o[2]);
      const contact = contactMap[contactId] || { name: 'Executive Leader', role: 'Decision Maker', email: '' };
      const company = compMap[contact.companyId] || { name: 'Target Manufacturing', headcount: '500-1,000', location: 'India' };
      const resp = respMap[outId] || respMap[contactId] || null;

      // Determine template archetype from htmlBody
      const html = String(o[6] || '');
      let archetype = 'Classic Letterhead';
      if (html.includes('Operations & Workforce Briefing')) archetype = 'Operations Brief';
      else if (html.includes('Watch Video Case Study Walkthrough')) archetype = 'Video Spotlight';
      else if (!html.includes('fluid-container')) archetype = 'Ultra-Human Note';

      const status = String(o[11] || 'DRAFT_GENERATED').trim();
      const sentDate = o[8] ? new Date(o[8]).toLocaleDateString('en-IN') : (o[7] ? 'Scheduled: ' + new Date(o[7]).toLocaleDateString('en-IN') : 'Draft Ready');

      crmRows.push({
        outreachId: outId,
        contactId: contactId,
        name: contact.name,
        email: contact.email,
        role: contact.role,
        company: company.name,
        headcount: company.headcount,
        location: company.location,
        leadScore: contact.leadScore,
        templateStyle: archetype,
        subject: o[5] || 'Workforce & Plant Systems Briefing',
        status: status,
        sentDate: sentDate,
        cvAttached: true,
        cvFileName: 'Ankit_Choudhary_Executive_HR_CV.pdf',
        response: resp ? {
          sentiment: resp.sentiment,
          snippet: resp.snippet,
          receivedAt: resp.receivedAt
        } : null
      });
    });

    return {
      success: true,
      crmData: crmRows,
      totalContacts: contacts.length,
      totalOutreaches: outreaches.length,
      totalResponses: responses.length
    };
  } catch (err) {
    return { success: false, error: 'Failed getting CRM data: ' + err.message };
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

function apiGenerateDrafts(campaignId, templateArchetype) {
  try {
    const res = generateCampaignOutreachDrafts(campaignId || 'CMPG-01', templateArchetype);
    const count = (res && typeof res.count === 'number') ? res.count : 0;
    return {
      success: true,
      count: count,
      archetype: res.archetype || templateArchetype,
      message: count > 0 
        ? ('Synthesized ' + count + ' outreach draft(s) ready for review!')
        : 'All eligible contacts already have drafts synthesized in the queue.'
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Re-renders an existing outreach draft row using a newly selected design archetype.
 */
function apiReRenderDraftWithTemplate(outreachId, newArchetype) {
  try {
    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    if (!ss) return { success: false, error: 'Spreadsheet unavailable' };

    const outreachSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OUTREACH_HISTORY);
    const contactSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.CONTACTS);
    const profileSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.CANDIDATE_PROFILE);
    const compSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.TARGET_COMPANIES);
    const oppSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OPPORTUNITIES);
    const projSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.PROJECTS);
    const expSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.WORK_EXPERIENCE);

    const outreachData = outreachSheet.getDataRange().getValues();
    let rowIndex = -1;
    let targetRow = null;

    for (let i = 1; i < outreachData.length; i++) {
      if (String(outreachData[i][0]) === outreachId) {
        rowIndex = i + 1;
        targetRow = outreachData[i];
        break;
      }
    }

    if (!targetRow) return { success: false, error: 'Outreach row not found: ' + outreachId };

    const contactId = String(targetRow[2]);
    const contacts = contactSheet.getDataRange().getValues();
    let contactRow = null;
    for (let i = 1; i < contacts.length; i++) {
      if (String(contacts[i][0]) === contactId) {
        contactRow = contacts[i];
        break;
      }
    }
    if (!contactRow) return { success: false, error: 'Contact not found: ' + contactId };

    const profile = profileSheet.getDataRange().getValues()[1] || [];
    const companies = buildLookupMap_(compSheet.getDataRange().getValues(), 0);
    const opps = buildLookupMap_(oppSheet.getDataRange().getValues(), 0);
    const projects = buildLookupMap_(projSheet.getDataRange().getValues(), 0);
    const experiences = expSheet.getDataRange().getValues();
    const primaryExp = experiences.length > 1 ? experiences[1] : [];

    const comp = companies[contactRow[1]] || [];
    const opp = opps[contactRow[2]] || [];
    const projId = opp[8] || targetRow[4] || 'PRJ-001';
    const proj = projects[projId] || [];

    const tokenPayload = {
      first_name: contactRow[3] || 'Leader',
      last_name: contactRow[4] || '',
      company_name: comp[1] || 'your organization',
      funding_stage_headcount: comp[4] || '500-1000 employees',
      recent_trigger_event_short: comp[5] || 'expanding plant operations and workforce footprint',
      role_title_opp: opp[2] || 'Head - Plant Human Resources & Industrial Relations',
      primary_pain_point: opp[6] || 'shopfloor absenteeism, contract labour compliance, and unstandardized appraisal systems',
      required_skills_top3: opp[7] || 'Plant HR, Industrial Relations, Statutory Compliance',
      required_skill_1: (opp[7] || 'Plant HR').split(',')[0].trim(),
      required_skill_2: ((opp[7] || 'Plant HR, IR').split(',')[1] || 'Industrial Relations').trim(),
      project_title: proj[2] || 'Workforce Attendance & PMS Institutionalization',
      architecture_summary: proj[6] || 'architected an in-house integrated employee workforce platform combining Safety reporting, KRA/KPI performance appraisals, and TNI-driven training workflows',
      quantified_business_impact: proj[7] || 'rolled out PMS to 100 senior leaders in 6 months and scaled coverage to 200 managers across 3,850 personnel',
      github_repo_url: appendUtm_(proj[8], comp[1], contactId),
      case_study_url_utm: appendUtm_(proj[9], comp[1], contactId),
      exp_company_name: primaryExp[2] || 'Bihar Foundry & Castings Ltd (GMG Group)',
      exp_role_title: primaryExp[3] || 'Head - Human Resources & Industrial Relations',
      exp_roi_statement: primaryExp[7] || 'governed 3,850 plant personnel with 0-strike industrial peace and in-house digitized workforce platforms',
      custom_roi_hypothesis: opp[9] || 'applying my 3,850-worker workforce governance and digital Safety/PMS framework can eliminate overtime leakage and optimize shopfloor productivity',
      candidate_full_name: profile[1] || 'Ankit Choudhary',
      candidate_headline: profile[2] || 'Manufacturing HR Leader | Plant HR, Industrial Relations, Workforce Strategy & Performance Transformation',
      years_experience: profile[3] || 15,
      github_profile_url: profile[6] || 'https://linkedin.com/in/ankit-choudhary-human-resources',
      portfolio_base_url: profile[5] || 'https://linkedin.com/in/ankit-choudhary-human-resources',
      portfolio_video_url: (profile[5] && profile[5].length > 10) ? appendUtm_(profile[5], comp[1], contactId) : 'https://linkedin.com/in/ankit-choudhary-human-resources',
      resume_drive_url: (profile[8] && profile[8].length > 10 && !profile[8].includes('1AnkitChoudhary_Resume_DriveID'))
        ? ('https://drive.google.com/file/d/' + profile[8] + '/view')
        : (ANTIGRAVITY_CONFIG.DEFAULT_RESUME_URL || 'https://drive.google.com/file/d/1FCyxOqXRGDnxBHugiJzRcy1OnJjbbckb/view'),
      calendar_booking_url: (profile[9] && profile[9].length > 5 && !profile[9].includes('ankitchoudharymba/hr-leadership-discussion'))
        ? profile[9]
        : (ANTIGRAVITY_CONFIG.DEFAULT_CALENDLY_URL || 'mailto:ankitchoudharymba@gmail.com?subject=Plant%20HR%20Leadership%20Discussion%20-%20Ankit%20Choudhary')
    };

    const rendered = renderPersonaTemplate_(contactRow[7], tokenPayload, newArchetype);

    // Save newly re-rendered subject & HTML body to Google Sheet
    outreachSheet.getRange(rowIndex, 6).setValue(rendered.subject);
    outreachSheet.getRange(rowIndex, 7).setValue(rendered.htmlBody);

    return {
      success: true,
      outreachId: outreachId,
      archetype: newArchetype,
      subject: rendered.subject,
      htmlBody: rendered.htmlBody
    };
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
    if (settings.apolloApiKey) {
      props.setProperty('APOLLO_API_KEY', settings.apolloApiKey.trim());
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
    const apolloKey = props.getProperty('APOLLO_API_KEY') || '';
    const maskedApolloKey = apolloKey ? (apolloKey.slice(0, 4) + '••••••••' + apolloKey.slice(-4)) : '';
    const sheetId = props.getProperty('SPREADSHEET_ID') || '';
    return {
      success: true,
      hasApiKey: !!apiKey,
      maskedApiKey: maskedKey,
      hasApolloKey: !!apolloKey,
      maskedApolloKey: maskedApolloKey,
      spreadsheetId: sheetId,
      spreadsheetUrl: ANTIGRAVITY_CONFIG.getSpreadsheetUrl()
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Apollo.io Integration API Bridges
 */
function apiTestApolloConnection(customKey) {
  return ANTIGRAVITY_APOLLO.testConnection(customKey);
}

function apiSearchApollo(queryOptions) {
  return ANTIGRAVITY_APOLLO.searchPeople(queryOptions);
}

function apiImportApolloLeads(leads) {
  return ANTIGRAVITY_APOLLO.importLeadsToCrm(leads);
}

function apiImportApolloCsv(csvText) {
  return ANTIGRAVITY_APOLLO.importCsvText(csvText);
}

/**
 * Returns the defined targeting segments and presets.
 */
function apiGetTargetSegments() {
  return {
    success: true,
    segments: ANTIGRAVITY_APOLLO.TARGET_SEGMENTS
  };
}

/**
 * Draft Review Studio: Dispatches a single test email directly to candidate for review.
 */
function apiSendTestEmail(outreachId, recipientOverride) {
  try {
    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    if (!ss) return { success: false, error: 'Database not initialized.' };
    const sheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OUTREACH_HISTORY);
    if (!sheet) return { success: false, error: 'Outreach_History table not found.' };

    const data = sheet.getDataRange().getValues();
    let draftRow = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(outreachId)) {
        draftRow = data[i];
        break;
      }
    }

    if (!draftRow) return { success: false, error: 'Draft ID ' + outreachId + ' not found.' };

    const targetEmail = recipientOverride || ANTIGRAVITY_CONFIG.getCandidateEmail() || (typeof Session !== 'undefined' ? Session.getActiveUser().getEmail() : 'ankitchoudharymba@gmail.com');
    const testSubject = '[TEST PREVIEW] ' + draftRow[5];
    const testBody = draftRow[6] + 
      '<br><br><hr style="border:none;border-top:1px solid #e2e8f0;margin:24px 0;">' +
      '<p style="color:#64748b;font-size:12px;font-family:sans-serif;line-height:1.5;">' +
      '⚡ <strong>Project Antigravity Test Delivery</strong><br>' +
      'Outreach ID: <code>' + outreachId + '</code> • Campaign: <code>' + draftRow[1] + '</code><br>' +
      '<em>This preview email was sent directly from your Antigravity Draft Studio for visual review and quality assurance.</em></p>';

    const plainText = testBody.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();

    // Fetch and prepare physical CV PDF attachment from Google Drive
    let attachments = [];
    let attachedName = '';
    let resumeDriveId = ANTIGRAVITY_CONFIG.DEFAULT_RESUME_DRIVE_ID;
    const candSheet = ss.getSheetByName('Candidate_Profile');
    if (candSheet && candSheet.getLastRow() >= 2) {
      const p = candSheet.getRange(2, 1, 1, 10).getValues()[0];
      if (p[8] && p[8].length > 10 && !p[8].includes('1AnkitChoudhary_Resume_DriveID')) {
        resumeDriveId = p[8];
      }
    }

    if (resumeDriveId && typeof DriveApp !== 'undefined') {
      try {
        const cvFile = DriveApp.getFileById(resumeDriveId);
        const fileName = ANTIGRAVITY_CONFIG.DEFAULT_RESUME_FILE_NAME || 'Ankit_Choudhary_Executive_HR_CV.pdf';
        attachments.push(cvFile.getAs('application/pdf').setName(fileName));
        attachedName = fileName;
      } catch (attErr) {
        console.warn('Could not attach CV from Drive: ' + attErr.message);
      }
    }

    const emailOptions = {
      htmlBody: testBody,
      attachments: attachments
    };

    try {
      if (typeof GmailApp !== 'undefined' && GmailApp.sendEmail) {
        GmailApp.sendEmail(targetEmail, testSubject, plainText, emailOptions);
      } else {
        MailApp.sendEmail({
          to: targetEmail,
          subject: testSubject,
          htmlBody: testBody,
          body: plainText,
          attachments: attachments
        });
      }
    } catch (e1) {
      MailApp.sendEmail({
        to: targetEmail,
        subject: testSubject,
        htmlBody: testBody,
        body: plainText,
        attachments: attachments
      });
    }

    return {
      success: true,
      message: 'Test draft delivered to ' + targetEmail + (attachedName ? ' with ' + attachedName + ' attached!' : '!'),
      targetEmail: targetEmail,
      attached: attachments.length > 0,
      fileName: attachedName
    };
  } catch (e) {
    return { success: false, error: 'Failed sending test email: ' + e.message };
  }
}

/**
 * Returns user settings for Web App configuration modal.
 */
function apiGetCandidateSettings() {
  try {
    const props = (typeof PropertiesService !== 'undefined' && PropertiesService.getScriptProperties) 
      ? PropertiesService.getScriptProperties() : null;
    const apolloKey = props ? (props.getProperty('APOLLO_API_KEY') || '') : '';
    const maskedApolloKey = apolloKey ? (apolloKey.slice(0, 4) + '••••••••' + apolloKey.slice(-4)) : '';

    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    if (!ss) {
      return {
        success: true,
        settings: {
          fullName: 'Ankit Choudhary',
          headline: 'Manufacturing HR Leader',
          calendarBookingUrl: ANTIGRAVITY_CONFIG.DEFAULT_CALENDLY_URL,
          resumeDriveId: ANTIGRAVITY_CONFIG.DEFAULT_RESUME_DRIVE_ID,
          linkedinUrl: 'https://linkedin.com/in/ankit-choudhary-human-resources',
          hasApolloKey: !!apolloKey,
          maskedApolloKey: maskedApolloKey
        }
      };
    }
    const sheet = ss.getSheetByName('Candidate_Profile');
    if (!sheet || sheet.getLastRow() < 2) {
      return {
        success: true,
        settings: {
          fullName: 'Ankit Choudhary',
          headline: 'Manufacturing HR Leader | Plant HR, Industrial Relations, Workforce Strategy & Performance Transformation',
          calendarBookingUrl: ANTIGRAVITY_CONFIG.DEFAULT_CALENDLY_URL,
          resumeDriveId: ANTIGRAVITY_CONFIG.DEFAULT_RESUME_DRIVE_ID,
          linkedinUrl: 'https://linkedin.com/in/ankit-choudhary-human-resources',
          hasApolloKey: !!apolloKey,
          maskedApolloKey: maskedApolloKey
        }
      };
    }
    const row = sheet.getRange(2, 1, 1, 11).getValues()[0];
    const calUrl = (row[9] && row[9].length > 5 && !row[9].includes('ankitchoudharymba/hr-leadership-discussion'))
      ? row[9] : ANTIGRAVITY_CONFIG.DEFAULT_CALENDLY_URL;
    const resId = (row[8] && row[8].length > 10 && !row[8].includes('1AnkitChoudhary_Resume_DriveID'))
      ? row[8] : ANTIGRAVITY_CONFIG.DEFAULT_RESUME_DRIVE_ID;

    return {
      success: true,
      settings: {
        fullName: row[1] || 'Ankit Choudhary',
        headline: row[2] || 'Manufacturing HR Leader',
        calendarBookingUrl: calUrl,
        resumeDriveId: resId,
        linkedinUrl: row[5] || 'https://linkedin.com/in/ankit-choudhary-human-resources',
        hasApolloKey: !!apolloKey,
        maskedApolloKey: maskedApolloKey
      }
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Saves updated candidate settings from Web App modal to Google Sheet & Script Properties.
 */
function apiSaveCandidateSettings(settings) {
  try {
    if (settings.apolloApiKey && settings.apolloApiKey.trim().length > 0) {
      if (typeof PropertiesService !== 'undefined' && PropertiesService.getScriptProperties) {
        if (settings.apolloApiKey === '__CLEAR__') {
          PropertiesService.getScriptProperties().deleteProperty('APOLLO_API_KEY');
        } else {
          PropertiesService.getScriptProperties().setProperty('APOLLO_API_KEY', settings.apolloApiKey.trim());
        }
      }
    }

    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    if (ss) {
      const sheet = ss.getSheetByName('Candidate_Profile');
      if (sheet && sheet.getLastRow() >= 2) {
        if (settings.fullName) sheet.getRange(2, 2).setValue(settings.fullName);
        if (settings.headline) sheet.getRange(2, 3).setValue(settings.headline);
        if (settings.calendarBookingUrl) sheet.getRange(2, 10).setValue(settings.calendarBookingUrl);
        if (settings.resumeDriveId) sheet.getRange(2, 9).setValue(settings.resumeDriveId);
        if (settings.linkedinUrl) sheet.getRange(2, 6).setValue(settings.linkedinUrl);
      }
    }
    return { success: true, message: 'Settings and Apollo credentials saved successfully!' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Draft Review Studio: Updates subject and HTML body after in-place editing.
 */
function apiUpdateOutreachDraft(outreachId, updatedSubject, updatedHtmlBody) {
  try {
    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    if (!ss) return { success: false, error: 'Database not initialized.' };
    const sheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OUTREACH_HISTORY);
    if (!sheet) return { success: false, error: 'Outreach_History table not found.' };

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(outreachId)) {
        if (updatedSubject) sheet.getRange(i + 1, 6).setValue(updatedSubject.trim());
        if (updatedHtmlBody) sheet.getRange(i + 1, 7).setValue(updatedHtmlBody.trim());
        return { success: true, message: 'Draft ' + outreachId + ' updated in CRM spreadsheet.' };
      }
    }
    return { success: false, error: 'Draft ID ' + outreachId + ' not found.' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Sourcing Studio: Ingests a single decision-maker lead from LinkedIn or referral.
 */
function apiQuickAddSingleLead(leadData) {
  try {
    if (!leadData || !leadData.email) {
      return { success: false, error: 'Lead email address is required.' };
    }
    return ANTIGRAVITY_APOLLO.importLeadsToCrm([leadData]);
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// Node module export for offline testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    apiGetCandidateSettings,
    apiSaveCandidateSettings,
    apiTestApolloConnection,
    apiSearchApollo
  };
}
