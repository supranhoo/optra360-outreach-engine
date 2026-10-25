/**
 * PROJECT ANTIGRAVITY: END-TO-END JOB OUTREACH INTELLIGENCE PLATFORM
 * Module 2 & 3: Campaign Queue Builder & Dynamic Multi-Persona Personalizer
 * (02_03_CampaignAndPersonalizer.js)
 */

/**
 * Generates hyper-personalized outreach email drafts for all eligible contacts
 * enrolled in an active campaign cohort, and appends them to Outreach_History with DRAFT_READY status.
 */
function generateCampaignOutreachDrafts(campaignId) {
  const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
  if (!ss) return;
  const profileSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.CANDIDATE_PROFILE);
  const contactSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.CONTACTS);
  const compSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.TARGET_COMPANIES);
  const oppSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OPPORTUNITIES);
  const projSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.PROJECTS);
  const expSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.WORK_EXPERIENCE);
  const campSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.EMAIL_CAMPAIGNS);
  const outreachSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OUTREACH_HISTORY);

  if (!profileSheet || !contactSheet || !campSheet || !outreachSheet) {
    throw new Error('Required platform sheets missing. Run bootstrap first.');
  }

  const profile = profileSheet.getDataRange().getValues()[1] || [];
  const contacts = contactSheet.getDataRange().getValues();
  const companies = buildLookupMap_(compSheet.getDataRange().getValues(), 0);
  const opps = buildLookupMap_(oppSheet.getDataRange().getValues(), 0);
  const projects = buildLookupMap_(projSheet.getDataRange().getValues(), 0);
  const experiences = expSheet.getDataRange().getValues();
  const campaigns = buildLookupMap_(campSheet.getDataRange().getValues(), 0);

  const activeCamp = campaigns[campaignId];
  if (!activeCamp || String(activeCamp[9]).toUpperCase() !== 'ACTIVE') {
    throw new Error('Campaign inactive, completed, or not found: ' + campaignId);
  }

  const primaryExp = experiences.length > 1 ? experiences[1] : [];
  const newOutreachRows = [];
  const now = new Date();

  // Auto-verify & score any unhandled contacts first so newly added leads are immediately eligible
  if (typeof processAndScoreNewContacts === 'function') {
    try {
      processAndScoreNewContacts();
    } catch (e) {
      console.warn('Auto-scoring note: ' + e.message);
    }
  }

  // Campaign parameters
  const campPersona = String(activeCamp[2]).trim().toUpperCase();
  const sendStartHr = Number(activeCamp[5]) || ANTIGRAVITY_CONFIG.BUSINESS_HOURS_START;
  const sendEndHr = Number(activeCamp[6]) || ANTIGRAVITY_CONFIG.BUSINESS_HOURS_END;

  // Re-read contacts in case scoring modified them
  const currentContacts = contactSheet.getDataRange().getValues();

  // Track contacts already drafted in this batch or Outreach_History
  const existingOutreachContacts = new Set(
    outreachSheet.getDataRange().getValues().slice(1).map(r => String(r[2]))
  );

  for (let i = 1; i < currentContacts.length; i++) {
    const c = currentContacts[i];
    const contactId = String(c[0]);
    const outreachStatus = String(c[12] || '').trim();
    const verificationStatus = String(c[6] || '').trim();
    const contactPersona = String(c[7] || '').trim().toUpperCase();

    // Verification & Eligibility Gate:
    // Eligible if status is QUEUED, or if newly added with verified email
    const isEligible = (outreachStatus === 'QUEUED') || 
      (verificationStatus === 'VERIFIED_VALID' && outreachStatus !== 'BOUNCED' && outreachStatus !== 'OPT_OUT');
    if (!isEligible) continue;

    // Multi-Persona Campaign Filter: CMPG-01 is multi-persona; otherwise match campPersona
    const isMultiPersonaCampaign = !campPersona || campPersona === 'ALL' || campPersona === 'ANY' || campaignId === 'CMPG-01';
    if (!isMultiPersonaCampaign && contactPersona !== campPersona) continue;

    if (existingOutreachContacts.has(contactId)) continue;

    const comp = companies[c[1]] || [];
    const opp = opps[c[2]] || [];
    const projId = opp[8] || 'PRJ-001';
    const proj = projects[projId] || (Object.keys(projects).length > 0 ? projects[Object.keys(projects)[0]] : []);

    // Parse pain points safely
    let primaryPain = 'shopfloor absenteeism and plant workforce governance';
    try {
      if (opp[6]) {
        const parsed = JSON.parse(opp[6]);
        if (Array.isArray(parsed) && parsed.length > 0) primaryPain = parsed[0];
      }
    } catch (e) {
      primaryPain = String(opp[6] || 'shopfloor absenteeism and plant workforce governance');
    }

    const skillsArray = String(opp[7] || '').split(',').map(s => s.trim()).filter(Boolean);

    const tokenPayload = {
      first_name: c[3] || 'there',
      last_name: c[4] || '',
      company_name: comp[1] || 'your organization',
      funding_stage_headcount: comp[4] || 'growing manufacturing operations',
      recent_trigger_event_short: comp[5] || 'scaling plant operations and workforce capacity',
      role_title_opp: opp[2] || 'Plant HR & Industrial Relations Leadership',
      primary_pain_point: primaryPain,
      required_skills_top3: skillsArray.slice(0, 3).join(', ') || 'plant-hr, ir, statutory-compliance',
      required_skill_1: skillsArray[0] || 'plant HR governance',
      required_skill_2: skillsArray[1] || 'statutory compliance',
      project_title: proj[2] || 'Integrated Safety, PMS & LMS Digital Workforce Platform',
      architecture_summary: proj[6] || 'architected an in-house integrated employee workforce platform combining Safety reporting, KRA/KPI performance appraisals, and TNI-driven training workflows',
      quantified_business_impact: proj[7] || 'rolled out PMS to 100 senior leaders in 6 months and scaled coverage to 200 managers across 3,850 personnel',
      github_repo_url: appendUtm_(proj[8], comp[1], c[0]),
      case_study_url_utm: appendUtm_(proj[9], comp[1], c[0]),
      exp_company_name: primaryExp[2] || 'Bihar Foundry & Castings Ltd (GMG Group)',
      exp_role_title: primaryExp[3] || 'Head - Human Resources & Industrial Relations',
      exp_roi_statement: primaryExp[7] || 'governed 3,850 plant personnel with 0-strike industrial peace and in-house digitized workforce platforms',
      custom_roi_hypothesis: opp[9] || 'applying my 3,850-worker workforce governance and digital Safety/PMS framework can eliminate overtime leakage and optimize shopfloor productivity',
      candidate_full_name: profile[1] || 'Ankit Choudhary',
      candidate_headline: profile[2] || 'Manufacturing HR Leader | Plant HR, Industrial Relations, Workforce Strategy & Performance Transformation',
      years_experience: profile[3] || 15,
      github_profile_url: profile[6] || 'https://linkedin.com/in/ankit-choudhary-human-resources',
      portfolio_base_url: profile[5] || 'https://linkedin.com/in/ankit-choudhary-human-resources',
      resume_drive_url: profile[8] ? ('https://drive.google.com/file/d/' + profile[8] + '/view') : '',
      calendar_booking_url: profile[9] || 'https://calendly.com/ankitchoudharymba/hr-leadership-discussion'
    };

    const rendered = renderPersonaTemplate_(c[7], tokenPayload);
    const scheduledTime = computeNextSendSlot_(now, newOutreachRows.length, sendStartHr, sendEndHr);
    const outreachId = 'OUT-' + Utilities.getUuid().slice(0, 8).toUpperCase();

    newOutreachRows.push([
      outreachId,
      campaignId,
      c[0],
      c[2],
      projId,
      rendered.subject,
      rendered.htmlBody,
      scheduledTime.toISOString(),
      '', // actual_sent_at
      '', // gmail_thread_id
      '', // gmail_message_id
      'DRAFT_READY'
    ]);

    existingOutreachContacts.add(contactId);
  }

  if (newOutreachRows.length > 0) {
    outreachSheet.getRange(outreachSheet.getLastRow() + 1, 1, newOutreachRows.length, 12).setValues(newOutreachRows);
    try {
      if (SpreadsheetApp.getActiveSpreadsheet()) {
        SpreadsheetApp.getActiveSpreadsheet().toast('Generated ' + newOutreachRows.length + ' outreach drafts in DRAFT_READY status.', 'Success', 5);
      }
    } catch (e) {}
    console.log('Generated ' + newOutreachRows.length + ' outreach drafts.');
  } else {
    try {
      if (SpreadsheetApp.getActiveSpreadsheet()) {
        SpreadsheetApp.getActiveSpreadsheet().toast('No eligible QUEUED contacts found matching campaign criteria.', 'Info', 4);
      }
    } catch (e) {}
    console.log('No eligible QUEUED contacts found matching campaign criteria.');
  }

  return {
    success: true,
    count: newOutreachRows.length,
    message: newOutreachRows.length > 0 
      ? ('Synthesized ' + newOutreachRows.length + ' outreach drafts in DRAFT_READY status.')
      : 'No new drafts synthesized (all eligible contacts already have drafts).'
  };
}

/**
 * Renders persona-aligned subject line and HTML body across the 6 template families.
 */
function renderPersonaTemplate_(persona, t) {
  let subject = '';
  let body = '';
  const p = String(persona || 'HIRING_MGR').toUpperCase();

  switch (p) {
    case 'FOUNDER':
      subject = 'Building institutional Plant HR & workforce systems at ' + t.company_name;
      body = '<p>Hi ' + t.first_name + ',</p>' +
        '<p>As manufacturing organizations scale across ' + t.funding_stage_headcount + ', ad-hoc shopfloor management, uncontrolled contract-labour absenteeism, and manual time-office loopholes often start directly eroding plant EBITDA and creating promoter bottlenecks.</p>' +
        '<p>At Bihar Foundry & Castings Ltd (GMG Group), I engineered the complete workforce governance turnaround from scratch across a <strong>3,850-personnel plant workforce</strong> (2,350 on-roll + 1,500 contract workers):</p>' +
        '<ul>' +
        '<li><strong>Attendance & Overtime Turnaround:</strong> Restructured biometric attendance governance and introduced a production-linked increment policy, drastically curbing shopfloor absenteeism and overtime leakage (<a href="' + t.case_study_url_utm + '">' + t.project_title + '</a>).</li>' +
        '<li><strong>0-to-1 PMS Implementation:</strong> Rolled out structured KRA/KPI appraisals to 100 senior leaders in 6 months, expanding coverage to 200 managers in 9 months.</li>' +
        '<li><strong>In-House Digital Platform:</strong> Architected a custom full-stack Safety, Appraisal & LMS platform without expensive multi-crore enterprise ERP bloat.</li>' +
        '<li><strong>Statutory & IR De-risking:</strong> 100% labour compliance (Factory Act, CLRA, BOCW) and maintained zero-strike industrial peace.</li>' +
        '</ul>' +
        '<p><strong>Why this matters for ' + t.company_name + ':</strong> ' + t.custom_roi_hypothesis + '</p>' +
        '<p>Full background & career portfolio: <a href="' + t.case_study_url_utm + '">' + t.case_study_url_utm + '</a>.</p>' +
        '<p>Would you be open to a 15-minute introductory discussion this week to explore how this systems playbook can institutionalize ' + t.company_name + '\'s plant operations as you scale?</p>' +
        '<p>Best regards,<br><strong>' + t.candidate_full_name + '</strong><br>' + t.candidate_headline + '<br><a href="' + t.calendar_booking_url + '">Schedule a 15-min call</a></p>';
      break;

    case 'TECH_LEAD':
    case 'CTO':
      subject = 'Curtailed shopfloor absenteeism & overtime across 3,850 plant personnel (case study)';
      body = '<p>Hi ' + t.first_name + ',</p>' +
        '<p>Managing multi-shift plant operations and expanding production capacity always comes down to shopfloor discipline: shift deployment adherence, contractor compliance, and supervisory capability.</p>' +
        '<p>In my recent role heading Plant HR & IR across a <strong>3,850-personnel manufacturing workforce</strong>, I resolved these exact operational bottlenecks:</p>' +
        '<ul>' +
        '<li><strong>Shopfloor Discipline:</strong> Restructured time-office biometric governance across 3 production shifts, curbing absenteeism and eliminating overtime leakage (<a href="' + t.case_study_url_utm + '">' + t.project_title + '</a>).</li>' +
        '<li><strong>Digital Safety & Workforce Platform:</strong> Architected an in-house Safety incident reporting and KRA/KPI platform adopted plant-wide.</li>' +
        '<li><strong>Zero-Strike IR Governance:</strong> De-risked union friction, labour contractor compliance, and local administrative liaison.</li>' +
        '</ul>' +
        '<p>I am exploring my next manufacturing leadership role and would welcome the opportunity to bring this proven systems playbook to ' + t.company_name + '.</p>' +
        '<p>Open to a brief 12-minute operational discussion this Thursday to share what worked across our heavy manufacturing plants?</p>' +
        '<p>Best regards,<br><strong>' + t.candidate_full_name + '</strong><br>' + t.candidate_headline + '</p>';
      break;

    case 'HIRING_MGR':
      subject = '0-to-1 PMS & Digital Safety rollout across plant operations — Ankit Choudhary';
      body = '<p>Hi ' + t.first_name + ',</p>' +
        '<p>I reviewed ' + t.company_name + '\'s growth in manufacturing operations and wanted to connect regarding your HR leadership and workforce governance roadmap.</p>' +
        '<p>Over 15 years leading Plant HR & Industrial Relations, I specialize in taking growing manufacturing organizations from ad-hoc processes to institutionalized, tech-enabled systems:</p>' +
        '<ol>' +
        '<li><strong>0-to-1 PMS Architecture:</strong> Built and deployed structured KRA/KPI performance appraisals for 100 leaders in 6 months, scaling to 200 managers across 3,850 personnel.</li>' +
        '<li><strong>Attendance & Absenteeism Overhaul:</strong> Formulated production-linked increment policies and biometric time-office governance to curb shopfloor absenteeism (<a href="' + t.case_study_url_utm + '">' + t.project_title + '</a>).</li>' +
        '<li><strong>L&D Division Setup:</strong> Established training architecture driven by real-time Training Need Identification (TNI) matrices.</li>' +
        '</ol>' +
        '<p>Full portfolio & leadership credentials: <a href="' + t.case_study_url_utm + '">' + t.case_study_url_utm + '</a>' + (t.resume_drive_url ? ' | <a href="' + t.resume_drive_url + '">Resume PDF</a>' : '') + '.</p>' +
        '<p>Do you have 15 minutes early next week to discuss your workforce roadmap and how my systems-implementation background aligns?</p>' +
        '<p>Best regards,<br><strong>' + t.candidate_full_name + '</strong><br>' + t.candidate_headline + '</p>';
      break;

    case 'RECRUITER':
      subject = 'Candidate: ' + t.role_title_opp + ' — 15 yrs Manufacturing HR Leadership + Plant Systems Turnaround';
      body = '<p>Hi ' + t.first_name + ',</p>' +
        '<p>I noticed you are leading talent acquisition for ' + t.company_name + ' and wanted to share my background for leadership opportunities in Plant HR, Industrial Relations, and Workforce Strategy.</p>' +
        '<p><strong>Quick Fit Overview:</strong></p>' +
        '<ul>' +
        '<li><strong>Experience:</strong> 15 years leading Plant HR, IR, and workforce governance across heavy manufacturing, steel, and foundry plants.</li>' +
        '<li><strong>Workforce Scale:</strong> Governed 3,850 personnel (2,350 on-roll + 1,500 contract workers) across 3 production shifts.</li>' +
        '<li><strong>Systems Implementation:</strong> Built 0-to-1 PMS (200 managers), digital Safety/LMS platform, and production-linked attendance policy (<a href="' + t.case_study_url_utm + '">' + t.project_title + '</a>).</li>' +
        '<li><strong>Compliance & IR:</strong> 100% statutory compliance (Factory Act, CLRA, BOCW) and zero-strike industrial harmony.</li>' +
        '<li><strong>Portfolio & CV:</strong> <a href="' + t.case_study_url_utm + '">' + t.case_study_url_utm + '</a>' + (t.resume_drive_url ? ' | <a href="' + t.resume_drive_url + '">Resume PDF</a>' : '') + '.</li>' +
        '</ul>' +
        '<p>Are you open to a brief 10-minute screening call this week?</p>' +
        '<p>Best regards,<br><strong>' + t.candidate_full_name + '</strong><br>' + t.candidate_headline + '</p>';
      break;

    case 'PEER':
      subject = 'Quick question on ' + t.company_name + '\'s manufacturing plant culture';
      body = '<p>Hi ' + t.first_name + ',</p>' +
        '<p>Came across your profile while researching ' + t.company_name + '\'s manufacturing operations around ' + t.recent_trigger_event_short + '.</p>' +
        '<p>I have led Plant HR and workforce governance across large industrial setups (most recently 3,850 personnel at Bihar Foundry & Castings Ltd), implementing digital workforce systems and attendance turnarounds (<a href="' + t.case_study_url_utm + '">' + t.project_title + '</a>).</p>' +
        '<p>I am exploring senior HR leadership opportunities with growing manufacturing firms and wanted to ask: how have you found the operational leadership and growth culture at ' + t.company_name + '?</p>' +
        '<p>Best regards,<br><strong>' + t.candidate_full_name + '</strong></p>';
      break;

    default: // General Problem-First Systems Implementation Pitch
      subject = 'Building institutional Plant HR & PMS governance at ' + t.company_name;
      body = '<p>Hi ' + t.first_name + ',</p>' +
        '<p>As manufacturing organizations scale across ' + t.funding_stage_headcount + ', building structured HR processes, curbing shopfloor absenteeism, and de-risking contract-labour compliance are critical to sustaining operating margins.</p>' +
        '<p>Over 15 years, I have architected and deployed these exact systems in heavy manufacturing:</p>' +
        '<ul>' +
        '<li><strong><a href="' + t.case_study_url_utm + '">' + t.project_title + '</a>:</strong> ' + t.architecture_summary + '</li>' +
        '<li><strong>Verified Business Impact:</strong> ' + t.quantified_business_impact + '</li>' +
        '</ul>' +
        '<p>Full case study & credentials: <a href="' + t.case_study_url_utm + '">' + t.case_study_url_utm + '</a>. Open to a 15-minute discussion this week, or feel free to grab a slot at <a href="' + t.calendar_booking_url + '">my calendar link</a>.</p>' +
        '<p>Best regards,<br><strong>' + t.candidate_full_name + '</strong><br>' + t.candidate_headline + '</p>';
  }

  return { subject: subject, htmlBody: body };
}

/**
 * Appends standard UTM telemetry parameters to portfolio/GitHub links.
 */
function appendUtm_(baseUrl, companyName, contactId) {
  if (!baseUrl) return '';
  const sep = String(baseUrl).includes('?') ? '&' : '?';
  const slug = String(companyName || 'org').toLowerCase().replace(/[^a-z0-9]/g, '');
  return baseUrl + sep + 'utm_source=antigravity&utm_company=' + slug + '&utm_cid=' + (contactId || '');
}

/**
 * Staggers scheduled send slots by 7-minute increments during business hours (Mon-Fri).
 * Automatically shifts weekend dates to Monday morning.
 */
function computeNextSendSlot_(baseDate, indexOffset, startHr, endHr) {
  const staggerMs = (ANTIGRAVITY_CONFIG.STAGGER_INTERVAL_MINS || 7) * 60000;
  const slot = new Date(baseDate.getTime() + (indexOffset * staggerMs));
  const hr = slot.getHours();

  if (hr < startHr) {
    slot.setHours(startHr, 15, 0, 0);
  }
  if (hr >= endHr) {
    slot.setDate(slot.getDate() + 1);
    slot.setHours(startHr, 15, 0, 0);
  }

  // Skip Saturday (6) and Sunday (0)
  if (slot.getDay() === 6) slot.setDate(slot.getDate() + 2);
  if (slot.getDay() === 0) slot.setDate(slot.getDate() + 1);

  return slot;
}

// Node.js export compatibility for offline testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    renderPersonaTemplate_,
    appendUtm_,
    computeNextSendSlot_
  };
}
