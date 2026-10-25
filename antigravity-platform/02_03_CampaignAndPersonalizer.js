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

  // Campaign parameters
  const campPersona = String(activeCamp[2]).trim().toUpperCase();
  const sendStartHr = Number(activeCamp[5]) || ANTIGRAVITY_CONFIG.BUSINESS_HOURS_START;
  const sendEndHr = Number(activeCamp[6]) || ANTIGRAVITY_CONFIG.BUSINESS_HOURS_END;

  // Track contacts already drafted in this batch or Outreach_History
  const existingOutreachContacts = new Set(
    outreachSheet.getDataRange().getValues().slice(1).map(r => String(r[2]))
  );

  for (let i = 1; i < contacts.length; i++) {
    const c = contacts[i];
    const contactId = String(c[0]);
    const outreachStatus = String(c[12] || '').trim();
    const verificationStatus = String(c[6] || '').trim();
    const contactPersona = String(c[7] || '').trim().toUpperCase();

    // Verification & Persona Match Gate
    if (outreachStatus !== 'QUEUED' || verificationStatus !== 'VERIFIED_VALID') continue;
    if (campPersona && contactPersona !== campPersona) continue;
    if (existingOutreachContacts.has(contactId)) continue;

    const comp = companies[c[1]] || [];
    const opp = opps[c[2]] || [];
    const projId = opp[8] || 'PRJ-001';
    const proj = projects[projId] || (Object.keys(projects).length > 0 ? projects[Object.keys(projects)[0]] : []);

    // Parse pain points safely
    let primaryPain = 'platform scalability & throughput';
    try {
      if (opp[6]) {
        const parsed = JSON.parse(opp[6]);
        if (Array.isArray(parsed) && parsed.length > 0) primaryPain = parsed[0];
      }
    } catch (e) {
      primaryPain = String(opp[6] || 'infrastructure scaling');
    }

    const skillsArray = String(opp[7] || '').split(',').map(s => s.trim()).filter(Boolean);

    const tokenPayload = {
      first_name: c[3] || 'there',
      last_name: c[4] || '',
      company_name: comp[1] || 'your team',
      funding_stage_headcount: comp[4] || 'high-growth stage',
      recent_trigger_event_short: comp[5] || 'scaling your technical roadmap',
      role_title_opp: opp[2] || 'Engineering Leadership',
      primary_pain_point: primaryPain,
      required_skills_top3: skillsArray.slice(0, 3).join(', ') || 'distributed systems, cloud pipelines',
      required_skill_1: skillsArray[0] || 'high-throughput architectures',
      required_skill_2: skillsArray[1] || 'cloud infrastructure',
      project_title: proj[2] || 'Distributed Architecture Engine',
      architecture_summary: proj[6] || 'built a resilient distributed event caching pipeline',
      quantified_business_impact: proj[7] || 'shaved latency by 45% and reduced cloud infrastructure overhead',
      github_repo_url: appendUtm_(proj[8], comp[1], c[0]),
      case_study_url_utm: appendUtm_(proj[9], comp[1], c[0]),
      exp_company_name: primaryExp[2] || 'prior engineering leadership',
      exp_role_title: primaryExp[3] || 'Lead Engineer',
      exp_roi_statement: primaryExp[7] || 'scaled core pipeline to 99.99% availability',
      custom_roi_hypothesis: opp[9] || 'applying these patterns can de-risk Q1 roadmap milestones by 4-6 weeks',
      candidate_full_name: profile[1] || 'Aarav Mehta',
      candidate_headline: profile[2] || 'Staff Systems & AI Infrastructure Engineer',
      years_experience: profile[3] || 8,
      github_profile_url: profile[6] || 'https://github.com',
      portfolio_base_url: profile[5] || 'https://portfolio.dev',
      resume_drive_url: profile[8] ? ('https://drive.google.com/file/d/' + profile[8] + '/view') : '',
      calendar_booking_url: profile[9] || 'https://cal.com'
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
      subject = 'De-risking ' + t.company_name + '\'s ' + t.primary_pain_point + ' velocity in Q1';
      body = '<p>Hi ' + t.first_name + ',</p>' +
        '<p>Congrats on ' + t.recent_trigger_event_short + '. At ' + t.funding_stage_headcount + ', shipping fast without bloating infrastructure burn is usually the #1 bottleneck.</p>' +
        '<p>I recently architected and open-sourced <a href="' + t.github_repo_url + '">' + t.project_title + '</a>, which ' + t.quantified_business_impact + '. Previously at ' + t.exp_company_name + ', I led the architecture that ' + t.exp_roi_statement + '.</p>' +
        '<p><strong>Why this matters for ' + t.company_name + ':</strong> ' + t.custom_roi_hypothesis + '</p>' +
        '<p>Full case study & live benchmarks: <a href="' + t.case_study_url_utm + '">' + t.case_study_url_utm + '</a>.</p>' +
        '<p>Worth a quick 10-minute founder-to-builder chat this week to see if I can take this off your plate?</p>' +
        '<p>Best regards,<br>' + t.candidate_full_name + '<br>' + t.candidate_headline + '</p>';
      break;

    case 'TECH_LEAD':
    case 'CTO':
      subject = 'Architecture note on ' + t.company_name + '\'s ' + t.primary_pain_point + ' pipeline (GitHub inside)';
      body = '<p>Hi ' + t.first_name + ',</p>' +
        '<p>Followed ' + t.company_name + '\'s work on ' + t.recent_trigger_event_short + '. Noticed from your ' + t.role_title_opp + ' opening that your platform team is tackling ' + t.primary_pain_point + ' across ' + t.required_skills_top3 + '.</p>' +
        '<p>When tackling that same concurrency and latency tradeoff, I built <a href="' + t.github_repo_url + '">' + t.project_title + '</a>:</p>' +
        '<ul>' +
        '<li><strong>Technical Approach:</strong> ' + t.architecture_summary + '</li>' +
        '<li><strong>Production Benchmark:</strong> ' + t.quantified_business_impact + '</li>' +
        '<li><strong>System Design Breakdown:</strong> <a href="' + t.case_study_url_utm + '">' + t.case_study_url_utm + '</a></li>' +
        '</ul>' +
        '<p>I\'m exploring my next senior engineering move and would love to bring this architectural playbook to ' + t.company_name + '.</p>' +
        '<p>Would love your technical feedback on the repo design — and if it aligns with what you need, open to a 12-minute technical chat this Thursday?</p>' +
        '<p>Best regards,<br>' + t.candidate_full_name + '<br>' + t.candidate_headline + '</p>';
      break;

    case 'HIRING_MGR':
      subject = 'Your ' + t.role_title_opp + ' search — shipped this exact stack in production';
      body = '<p>Hi ' + t.first_name + ',</p>' +
        '<p>I reviewed your specification for the ' + t.role_title_opp + ' role — specifically the requirement around ' + t.primary_pain_point + ' and ' + t.required_skills_top3 + '.</p>' +
        '<p>Over the past ' + t.years_experience + ' years, I\'ve solved this exact bottleneck in production:</p>' +
        '<ol>' +
        '<li><a href="' + t.github_repo_url + '">' + t.project_title + '</a>: ' + t.architecture_summary + ' &rarr; ' + t.quantified_business_impact + '</li>' +
        '<li>' + t.exp_company_name + ' (' + t.exp_role_title + '): ' + t.exp_roi_statement + '</li>' +
        '</ol>' +
        '<p>Because I\'ve already built and battle-tested these modules, my ramp-up time on your team\'s Q1 backlog would be near zero. Full architecture case study: <a href="' + t.case_study_url_utm + '">' + t.case_study_url_utm + '</a>.</p>' +
        '<p>Do you have 15 minutes early next week to discuss how my prior work aligns with your roadmap?</p>' +
        '<p>Best regards,<br>' + t.candidate_full_name + '<br>' + t.candidate_headline + '</p>';
      break;

    case 'RECRUITER':
      subject = t.role_title_opp + ' candidate — ' + t.years_experience + ' yrs ' + t.required_skill_1 + ' + live production proof';
      body = '<p>Hi ' + t.first_name + ',</p>' +
        '<p>Saw you\'re leading talent acquisition for ' + t.company_name + '\'s engineering team and wanted to make evaluating my background for the ' + t.role_title_opp + ' role effortless.</p>' +
        '<p>Quick 30-second fit checklist against your requirements:</p>' +
        '<ul>' +
        '<li><strong>Core Stack (' + t.required_skills_top3 + '):</strong> ' + t.years_experience + ' years production experience at ' + t.exp_company_name + '.</li>' +
        '<li><strong>Verified Business Impact:</strong> ' + t.exp_roi_statement + '.</li>' +
        '<li><strong>Live Proof-of-Work:</strong> Built <a href="' + t.github_repo_url + '">' + t.project_title + '</a>, which ' + t.quantified_business_impact + '.</li>' +
        '<li><strong>Portfolio & Deep Dives:</strong> <a href="' + t.case_study_url_utm + '">' + t.case_study_url_utm + '</a>' + (t.resume_drive_url ? ' | <a href="' + t.resume_drive_url + '">Resume PDF</a>' : '') + '.</li>' +
        '</ul>' +
        '<p>I\'ve submitted my formal application via your portal — do you have 10 minutes this week for an initial screen?</p>' +
        '<p>Best regards,<br>' + t.candidate_full_name + '<br>' + t.candidate_headline + '</p>';
      break;

    case 'PEER':
      subject = 'Quick question on ' + t.company_name + '\'s engineering culture (+ ' + t.project_title + ' repo)';
      body = '<p>Hi ' + t.first_name + ',</p>' +
        '<p>Came across your profile while researching ' + t.company_name + '\'s work around ' + t.recent_trigger_event_short + '.</p>' +
        '<p>I\'m preparing to apply for the ' + t.role_title_opp + ' opening. Since my recent open-source work on <a href="' + t.github_repo_url + '">' + t.project_title + '</a> (' + t.quantified_business_impact + ') maps closely to what your team is building, I wanted to ask: how have you found the engineering autonomy and technical velocity at ' + t.company_name + '?</p>' +
        '<p>If my portfolio (<a href="' + t.case_study_url_utm + '">' + t.case_study_url_utm + '</a>) looks like a strong fit, would you be open to dropping a quick internal referral link? Either way, appreciate any insights!</p>' +
        '<p>Best regards,<br>' + t.candidate_full_name + '</p>';
      break;

    default: // General Problem-First Proof-of-Work Pitch
      subject = 'Reducing ' + t.company_name + '\'s ' + t.primary_pain_point + ' bottleneck (GitHub repo inside)';
      body = '<p>Hi ' + t.first_name + ',</p>' +
        '<p>Noticed your team at ' + t.company_name + ' is hiring for the <strong>' + t.role_title_opp + '</strong> role to tackle ' + t.primary_pain_point + ' across ' + t.required_skills_top3 + '.</p>' +
        '<p>Instead of just sending a static resume, I wanted to share a production system I built for this exact challenge:</p>' +
        '<ul>' +
        '<li><strong><a href="' + t.github_repo_url + '">' + t.project_title + '</a>:</strong> ' + t.architecture_summary + '</li>' +
        '<li><strong>Measured ROI:</strong> ' + t.quantified_business_impact + '</li>' +
        '</ul>' +
        '<p>Full architecture breakdown: <a href="' + t.case_study_url_utm + '">' + t.case_study_url_utm + '</a>. Open to a 12-minute technical chat this Thursday, or feel free to grab a slot at <a href="' + t.calendar_booking_url + '">my calendar link</a>?</p>' +
        '<p>Best regards,<br>' + t.candidate_full_name + '<br>' + t.candidate_headline + '</p>';
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
