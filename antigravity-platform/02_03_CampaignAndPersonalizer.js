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
      portfolio_video_url: (profile[5] && profile[5].length > 10) ? appendUtm_(profile[5], comp[1], c[0]) : 'https://linkedin.com/in/ankit-choudhary-human-resources',
      resume_drive_url: profile[8] ? ('https://drive.google.com/file/d/' + profile[8] + '/view') : 'https://linkedin.com/in/ankit-choudhary-human-resources',
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
 * Helper to build the responsive executive HTML email layout.
 * Includes header banner, video portfolio showcase card, 3-metric proof grid, and dual CTA buttons.
 */
function buildExecutiveHtmlEmail_(cfg, t) {
  const videoUrl = t.portfolio_video_url || t.case_study_url_utm || t.portfolio_base_url || 'https://linkedin.com/in/ankit-choudhary-human-resources';
  const cvUrl = t.resume_drive_url || t.portfolio_base_url || 'https://linkedin.com/in/ankit-choudhary-human-resources';
  const calendarUrl = t.calendar_booking_url || 'https://calendly.com/ankitchoudharymba/hr-leadership-discussion';

  return '<div style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif; max-width: 620px; margin: 0 auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; color: #1e293b; line-height: 1.6;">' +
    '<!-- Executive Header Banner -->' +
    '<div style="background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); padding: 24px 28px; border-bottom: 3px solid #38bdf8;">' +
      '<div style="font-size: 11px; font-weight: 700; color: #38bdf8; text-transform: uppercase; letter-spacing: 0.1em; margin-bottom: 4px;">Executive Leadership Dossier</div>' +
      '<div style="font-size: 22px; font-weight: 800; color: #ffffff; letter-spacing: -0.02em;">' + t.candidate_full_name + '</div>' +
      '<div style="font-size: 13px; color: #cbd5e1; margin-top: 3px;">' + t.candidate_headline + '</div>' +
    '</div>' +

    '<!-- Body Content Container -->' +
    '<div style="padding: 24px 28px;">' +
      '<p style="font-size: 15px; margin-top: 0; color: #0f172a;">Hi <strong>' + t.first_name + '</strong>,</p>' +
      '<p style="font-size: 14px; color: #334155; line-height: 1.6;">' + cfg.openingHook + '</p>' +

      '<!-- Featured Video Portfolio & Case Study Card -->' +
      '<div style="background: #0f172a; border-radius: 8px; padding: 20px 22px; margin: 20px 0; border: 1px solid #1e293b; color: #ffffff;">' +
        '<div style="display: inline-block; background: rgba(56, 189, 248, 0.15); border: 1px solid #38bdf8; color: #38bdf8; font-size: 10px; font-weight: 800; padding: 3px 8px; border-radius: 4px; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 10px;">' +
          '🎬 Featured Video Portfolio & Case Study' +
        '</div>' +
        '<div style="font-size: 15px; font-weight: 700; color: #f8fafc; line-height: 1.35; margin-bottom: 6px;">' +
          'Workforce Systems Turnaround: Scaling from Ad-Hoc Management to Institutional Governance' +
        '</div>' +
        '<div style="font-size: 12px; color: #94a3b8; line-height: 1.55; margin-bottom: 14px;">' +
          'A brief executive walkthrough detailing how we curbed shopfloor absenteeism, deployed 0-to-1 PMS, and maintained zero-strike IR peace across 3,850 personnel at Bihar Foundry & Castings Ltd.' +
        '</div>' +
        '<a href="' + videoUrl + '" target="_blank" style="display: inline-block; background: #38bdf8; color: #0f172a; font-weight: 800; font-size: 12px; padding: 9px 18px; border-radius: 5px; text-decoration: none; letter-spacing: 0.02em;">' +
          '▶ Watch Executive Walkthrough & Portfolio →' +
        '</a>' +
      '</div>' +

      '<!-- Verified Scale & Proof Metrics (3-Column Table Grid) -->' +
      '<div style="font-size: 11px; font-weight: 800; color: #64748b; text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 10px;">' +
        'Proven Scale & Verified Achievements across 3,850-personnel plant workforce:' +
      '</div>' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom: 22px;">' +
        '<tr>' +
          '<td width="31%" style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 12px 10px; text-align: center; vertical-align: top;">' +
            '<div style="font-size: 20px; font-weight: 800; color: #0284c7; line-height: 1.1;">3,850</div>' +
            '<div style="font-size: 10px; font-weight: 800; color: #0f172a; margin-top: 4px; text-transform: uppercase;">Plant Workforce</div>' +
            '<div style="font-size: 10px; color: #64748b; margin-top: 2px;">2,350 On-roll + 1,500 Contract</div>' +
          '</td>' +
          '<td width="3.5%"></td>' +
          '<td width="31%" style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 12px 10px; text-align: center; vertical-align: top;">' +
            '<div style="font-size: 20px; font-weight: 800; color: #10b981; line-height: 1.1;">100%</div>' +
            '<div style="font-size: 10px; font-weight: 800; color: #0f172a; margin-top: 4px; text-transform: uppercase;">Strike-Free IR</div>' +
            '<div style="font-size: 10px; color: #64748b; margin-top: 2px;">Zero Labour Stoppages</div>' +
          '</td>' +
          '<td width="3.5%"></td>' +
          '<td width="31%" style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 12px 10px; text-align: center; vertical-align: top;">' +
            '<div style="font-size: 20px; font-weight: 800; color: #8b5cf6; line-height: 1.1;">0-to-1</div>' +
            '<div style="font-size: 10px; font-weight: 800; color: #0f172a; margin-top: 4px; text-transform: uppercase;">PMS Rollout</div>' +
            '<div style="font-size: 10px; color: #64748b; margin-top: 2px;">200 Managers with KRAs</div>' +
          '</td>' +
        '</tr>' +
      '</table>' +

      '<!-- Targeted Execution Playbook -->' +
      '<div style="font-size: 11px; font-weight: 800; color: #64748b; text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 8px;">' +
        'How I Can Support ' + t.company_name + ':' +
      '</div>' +
      '<ul style="padding-left: 20px; margin: 0 0 20px 0; font-size: 13px; color: #334155; line-height: 1.6;">' +
        cfg.playbookBullets.map(function(b) { return '<li style="margin-bottom: 6px;">' + b + '</li>'; }).join('') +
      '</ul>' +

      '<!-- Dual Action Buttons: CV & Calendar -->' +
      '<table width="100%" cellpadding="0" cellspacing="0" style="margin: 24px 0 16px 0;">' +
        '<tr>' +
          '<td align="center">' +
            '<a href="' + cvUrl + '" target="_blank" style="display: inline-block; background: #0f172a; color: #ffffff; font-weight: 700; font-size: 12px; padding: 11px 20px; border-radius: 6px; text-decoration: none; margin-right: 8px; border: 1px solid #334155;">' +
              '📄 View Full Executive CV (PDF) →' +
            '</a>' +
            '<a href="' + calendarUrl + '" target="_blank" style="display: inline-block; background: #0284c7; color: #ffffff; font-weight: 700; font-size: 12px; padding: 11px 20px; border-radius: 6px; text-decoration: none;">' +
              '📅 Schedule 15-Min Discussion →' +
            '</a>' +
          '</td>' +
        '</tr>' +
      '</table>' +

      '<p style="font-size: 13px; color: #475569; margin-top: 20px; line-height: 1.5;">' +
        cfg.closingAsk +
      '</p>' +

      '<!-- Executive Signoff -->' +
      '<div style="border-top: 1px solid #e2e8f0; padding-top: 16px; margin-top: 22px;">' +
        '<div style="font-weight: 800; font-size: 14px; color: #0f172a;">' + t.candidate_full_name + '</div>' +
        '<div style="font-size: 12px; color: #64748b; margin-top: 2px;">' + t.candidate_headline + '</div>' +
        '<div style="font-size: 11px; color: #0284c7; margin-top: 4px;">' +
          '<a href="' + (t.portfolio_base_url || 'https://linkedin.com/in/ankit-choudhary-human-resources') + '" target="_blank" style="color: #0284c7; text-decoration: none; font-weight: 600;">LinkedIn Executive Profile ↗</a>' +
        '</div>' +
      '</div>' +
    '</div>' +
  '</div>';
}

/**
 * Renders persona-aligned subject line and HTML body across the 6 template families.
 */
function renderPersonaTemplate_(persona, t) {
  let subject = '';
  const p = String(persona || 'HIRING_MGR').toUpperCase();
  let cfg = { openingHook: '', playbookBullets: [], closingAsk: '' };

  switch (p) {
    case 'FOUNDER':
      subject = 'Plant HR & workforce systems — Ankit Choudhary (' + t.company_name + ')';
      cfg.openingHook = 'As ' + t.company_name + ' scales manufacturing operations across ' + t.funding_stage_headcount + ', ad-hoc shopfloor management, uncontrolled contract-labour absenteeism, and manual time-office loopholes often directly erode plant EBITDA and create persistent promoter bottlenecks.';
      cfg.playbookBullets = [
        '<strong>Attendance & Overtime Turnaround:</strong> Restructured biometric attendance governance and introduced production-linked increment policies to eliminate shopfloor absenteeism and overtime leakage (<a href="' + t.case_study_url_utm + '">' + t.project_title + '</a>).',
        '<strong>0-to-1 PMS Implementation:</strong> Rolled out structured KRA/KPI appraisals to 100 senior leaders in 6 months, expanding coverage to 200 managers across 3,850 plant personnel.',
        '<strong>In-House Digital Architecture:</strong> Deployed integrated Safety reporting, Appraisal, and LMS workflows without multi-crore enterprise ERP bloat.',
        '<strong>Statutory & IR De-risking:</strong> 100% compliance across Factory Act, CLRA & BOCW with zero strikes and zero labour stoppages at Bihar Foundry & Castings Ltd.'
      ];
      cfg.closingAsk = 'Would you be open to a 15-minute introductory discussion this week to explore how this systems playbook can institutionalize ' + t.company_name + '\'s plant operations as you scale?';
      break;

    case 'TECH_LEAD':
    case 'CTO':
      subject = 'Curtailed shopfloor absenteeism & overtime across 3,850 plant personnel (case study)';
      cfg.openingHook = 'Managing multi-shift plant operations and expanding production capacity always comes down to shopfloor discipline: shift deployment adherence, contractor compliance, and supervisory capability.';
      cfg.playbookBullets = [
        '<strong>Shopfloor Discipline:</strong> Overhauled biometric time-office governance across 3 production shifts, curbing absenteeism and eliminating overtime leakage in a 3,850-personnel manufacturing workforce (<a href="' + t.case_study_url_utm + '">' + t.project_title + '</a>).',
        '<strong>Digital Safety & Workforce Governance:</strong> Architected an in-house Safety incident reporting and KRA/KPI framework adopted plant-wide.',
        '<strong>Zero-Strike IR Governance:</strong> De-risked union friction, labour contractor compliance, and local administrative liaison.'
      ];
      cfg.closingAsk = 'I am exploring my next manufacturing leadership role and would welcome a brief 12-minute operational discussion to share what worked across our heavy manufacturing plants.';
      break;

    case 'HIRING_MGR':
      subject = '0-to-1 PMS & Digital Safety rollout across plant operations — Ankit Choudhary';
      cfg.openingHook = 'I reviewed ' + t.company_name + '\'s growth in manufacturing operations and wanted to connect regarding your HR leadership and workforce governance roadmap.';
      cfg.playbookBullets = [
        '<strong>0-to-1 PMS Architecture:</strong> Built and deployed structured KRA/KPI performance appraisals for 100 leaders in 6 months, scaling to 200 managers across 3,850 personnel.',
        '<strong>Attendance & Absenteeism Overhaul:</strong> Formulated production-linked increment policies and biometric time-office governance to curb shopfloor absenteeism (<a href="' + t.case_study_url_utm + '">' + t.project_title + '</a>).',
        '<strong>L&D Division Setup:</strong> Established training architecture driven by real-time Training Need Identification (TNI) matrices.'
      ];
      cfg.closingAsk = 'Do you have 15 minutes early next week to discuss your workforce roadmap and how my systems-implementation background aligns?';
      break;

    case 'RECRUITER':
      subject = 'Candidate: ' + t.role_title_opp + ' — 15 yrs Manufacturing HR Leadership + Plant Systems Turnaround';
      cfg.openingHook = 'I noticed you are leading talent acquisition for ' + t.company_name + ' and wanted to share my background for leadership opportunities in Plant HR, Industrial Relations, and Workforce Strategy.';
      cfg.playbookBullets = [
        '<strong>15 Years Leadership:</strong> Plant HR, IR, and workforce governance across heavy manufacturing, steel, and foundry plants.',
        '<strong>Workforce Scale:</strong> Governed 3,850 personnel (2,350 on-roll + 1,500 contract workers) across 3 production shifts.',
        '<strong>0-to-1 Systems:</strong> Built PMS (200 managers), digital Safety/LMS platform, and production-linked attendance policy (<a href="' + t.case_study_url_utm + '">' + t.project_title + '</a>).',
        '<strong>100% Compliance:</strong> Complete statutory governance (Factory Act, CLRA, BOCW) and zero-strike industrial harmony.'
      ];
      cfg.closingAsk = 'Are you open to a brief 10-minute screening call this week?';
      break;

    case 'PEER':
      subject = 'Quick question on ' + t.company_name + '\'s manufacturing plant culture';
      cfg.openingHook = 'Came across your profile while researching ' + t.company_name + '\'s manufacturing operations around ' + t.recent_trigger_event_short + '. I have led Plant HR and workforce governance across large industrial setups (most recently 3,850 personnel at Bihar Foundry & Castings Ltd), implementing digital workforce systems and attendance turnarounds (<a href="' + t.case_study_url_utm + '">' + t.project_title + '</a>).';
      cfg.playbookBullets = [
        'Curtailed shopfloor absenteeism and overtime via biometric governance.',
        'Architected 0-to-1 PMS and L&D frameworks across 200 managers.',
        'Maintained 100% industrial peace and zero labour stoppages.'
      ];
      cfg.closingAsk = 'I am exploring senior HR leadership opportunities with growing manufacturing firms and wanted to ask: how have you found the operational leadership and growth culture at ' + t.company_name + '?';
      break;

    default:
      subject = 'Building institutional Plant HR & PMS governance at ' + t.company_name;
      cfg.openingHook = 'As manufacturing organizations scale across ' + t.funding_stage_headcount + ', building structured HR processes, curbing shopfloor absenteeism, and de-risking contract-labour compliance are critical to sustaining operating margins.';
      cfg.playbookBullets = [
        '<strong>' + t.project_title + ':</strong> ' + t.architecture_summary,
        '<strong>Verified Business Impact:</strong> ' + t.quantified_business_impact,
        '<strong>Turnaround Scale:</strong> Governed 3,850 personnel with 0-strike industrial peace.'
      ];
      cfg.closingAsk = 'Full case study & credentials above. Open to a 15-minute discussion this week to explore how this playbook aligns with ' + t.company_name + '.';
  }

  const htmlBody = buildExecutiveHtmlEmail_(cfg, t);
  return { subject: subject, htmlBody: htmlBody };
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
