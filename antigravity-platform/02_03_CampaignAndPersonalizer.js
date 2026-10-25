/**
 * PROJECT ANTIGRAVITY: END-TO-END JOB OUTREACH INTELLIGENCE PLATFORM
 * Module 2 & 3: Campaign Queue Builder & Dynamic Multi-Persona Personalizer
 * (02_03_CampaignAndPersonalizer.js)
 */

/**
 * Generates hyper-personalized outreach email drafts for all eligible contacts
 * enrolled in an active campaign cohort, and appends them to Outreach_History with DRAFT_READY status.
 */
function generateCampaignOutreachDrafts(campaignId, templateArchetype) {
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
  const selectedArchetype = templateArchetype || (ANTIGRAVITY_CONFIG.DEFAULT_TEMPLATE_ARCHETYPE || 'CLASSIC_LETTER');

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
      resume_drive_url: (profile[8] && profile[8].length > 10 && !profile[8].includes('1AnkitChoudhary_Resume_DriveID'))
        ? ('https://drive.google.com/file/d/' + profile[8] + '/view')
        : (ANTIGRAVITY_CONFIG.DEFAULT_RESUME_URL || 'https://drive.google.com/file/d/1FCyxOqXRGDnxBHugiJzRcy1OnJjbbckb/view'),
      calendar_booking_url: (profile[9] && profile[9].length > 5 && !profile[9].includes('ankitchoudharymba/hr-leadership-discussion'))
        ? profile[9]
        : (ANTIGRAVITY_CONFIG.DEFAULT_CALENDLY_URL || 'mailto:ankitchoudharymba@gmail.com?subject=Plant%20HR%20Leadership%20Discussion%20-%20Ankit%20Choudhary')
    };

    const rendered = renderPersonaTemplate_(c[7], tokenPayload, selectedArchetype);
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
 * Returns embedded responsive CSS for mobile/desktop email clients.
 * Media queries scale down layouts on screens <= 600px, stack columns,
 * expand touch buttons to 100% full width, and adjust typography.
 */
function getResponsiveEmailStyles_() {
  return '<style type="text/css">\n' +
    '  @media only screen and (max-width: 600px) {\n' +
    '    .fluid-container { width: 100% !important; max-width: 100% !important; border-radius: 0 !important; border-left: none !important; border-right: none !important; }\n' +
    '    .fluid-padding { padding: 18px 14px !important; }\n' +
    '    .stack-col { display: block !important; width: 100% !important; max-width: 100% !important; margin: 8px 0 !important; box-sizing: border-box !important; text-align: left !important; padding: 14px 16px !important; }\n' +
    '    .mobile-full-btn { display: block !important; width: 100% !important; margin: 8px 0 !important; padding: 13px 16px !important; font-size: 13.5px !important; text-align: center !important; box-sizing: border-box !important; }\n' +
    '    .mobile-header-table td { display: block !important; width: 100% !important; text-align: left !important; }\n' +
    '    .mobile-header-table td:last-child { margin-top: 10px !important; }\n' +
    '    .mobile-metric-val { font-size: 22px !important; }\n' +
    '  }\n' +
    '</style>\n';
}

/**
 * ARCHETYPE 1: Classic Executive Letter
 * Understated executive letterhead, fluid-hybrid stackable 3-column proof cards,
 * clean video walkthrough callout, and full-width touch buttons on mobile.
 */
function buildClassicLetterEmail_(cfg, t) {
  const videoUrl = t.portfolio_video_url || t.case_study_url_utm || t.portfolio_base_url || 'https://linkedin.com/in/ankit-choudhary-human-resources';
  const cvUrl = t.resume_drive_url || (typeof ANTIGRAVITY_CONFIG !== 'undefined' ? ANTIGRAVITY_CONFIG.DEFAULT_RESUME_URL : 'https://drive.google.com/file/d/1FCyxOqXRGDnxBHugiJzRcy1OnJjbbckb/view');
  const defaultCalUrl = (typeof ANTIGRAVITY_CONFIG !== 'undefined' && ANTIGRAVITY_CONFIG.DEFAULT_GOOGLE_CALENDAR_URL)
    ? ANTIGRAVITY_CONFIG.DEFAULT_GOOGLE_CALENDAR_URL
    : 'https://calendar.google.com/calendar/render?action=TEMPLATE&text=HR+Transformation+%26+Operations+Discussion+-+Ankit+Choudhary&details=15-minute+discussion+on+manufacturing+workforce+strategy%2C+PMS+and+plant+operations.+Candidate%3A+Ankit+Choudhary+(ankitchoudharymba%40gmail.com)&add=ankitchoudharymba@gmail.com';
  const calendarUrl = (t.calendar_booking_url && !t.calendar_booking_url.includes('ankitchoudharymba/hr-leadership-discussion') && !t.calendar_booking_url.startsWith('mailto:'))
    ? t.calendar_booking_url
    : defaultCalUrl;

  return getResponsiveEmailStyles_() +
    '<div class="fluid-container" style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif; max-width: 620px; width: 100%; margin: 0 auto; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; color: #1e293b; line-height: 1.6; box-sizing: border-box;">' +
      '<!-- Executive Letterhead -->' +
      '<div class="fluid-padding" style="background: #f8fafc; padding: 18px 24px; border-bottom: 1px solid #e2e8f0;">' +
        '<table class="mobile-header-table" width="100%" cellpadding="0" cellspacing="0">' +
          '<tr>' +
            '<td>' +
              '<div style="font-size: 16px; font-weight: 700; color: #0f172a; letter-spacing: -0.01em;">' + t.candidate_full_name + '</div>' +
              '<div style="font-size: 12px; color: #64748b; margin-top: 2px;">' + t.candidate_headline + '</div>' +
            '</td>' +
            '<td align="right" style="vertical-align: middle;">' +
              '<a href="' + cvUrl + '" target="_blank" style="display: inline-block; background: #0f172a; color: #ffffff; font-size: 11px; font-weight: 600; padding: 6px 12px; border-radius: 4px; text-decoration: none;">📄 Attached CV (PDF)</a>' +
            '</td>' +
          '</tr>' +
        '</table>' +
      '</div>' +

      '<!-- Body Content Container -->' +
      '<div class="fluid-padding" style="padding: 24px 28px;">' +
        '<p style="font-size: 15px; margin-top: 0; color: #0f172a;">Hi <strong>' + t.first_name + '</strong>,</p>' +
        '<p style="font-size: 14px; color: #334155; line-height: 1.6; margin-bottom: 18px;">' + cfg.openingHook + '</p>' +

        '<!-- Video Portfolio & Systems Case Study -->' +
        '<div style="background: #f8fafc; border: 1px solid #e2e8f0; border-left: 4px solid #0284c7; border-radius: 6px; padding: 16px 18px; margin: 20px 0;">' +
          '<div style="font-size: 13px; font-weight: 700; color: #0f172a; margin-bottom: 4px;">' +
            '📹 Video Portfolio & Systems Case Study Walkthrough' +
          '</div>' +
          '<div style="font-size: 12px; color: #475569; line-height: 1.5; margin-bottom: 12px;">' +
            'A concise executive walkthrough showing how we curbed shopfloor absenteeism, deployed 0-to-1 PMS, and maintained zero-strike IR peace across 3,850 personnel at Bihar Foundry & Castings Ltd.' +
          '</div>' +
          '<a href="' + videoUrl + '" target="_blank" class="mobile-full-btn" style="display: inline-block; background: #0284c7; color: #ffffff; font-weight: 600; font-size: 12px; padding: 8px 16px; border-radius: 4px; text-decoration: none;">' +
            '▶ Watch Video Walkthrough & Case Study →' +
          '</a>' +
        '</div>' +

        '<!-- Proven Scale & Verified Metrics (Fluid-Hybrid Responsive Stackable) -->' +
        '<div style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 10px;">' +
          'Proven Scale & Verified Achievements across 3,850-personnel plant workforce:' +
        '</div>' +
        '<div style="text-align: center; font-size: 0; margin-bottom: 22px;">' +
          '<div class="stack-col" style="display: inline-block; width: 100%; max-width: 178px; vertical-align: top; margin: 4px; box-sizing: border-box; text-align: center; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 14px 10px;">' +
            '<div class="mobile-metric-val" style="font-size: 20px; font-weight: 800; color: #0284c7; line-height: 1.1;">3,850</div>' +
            '<div style="font-size: 10px; font-weight: 700; color: #0f172a; margin-top: 4px; text-transform: uppercase;">Plant Workforce</div>' +
            '<div style="font-size: 10px; color: #64748b; margin-top: 2px;">2,350 On-roll + 1,500 Contract</div>' +
          '</div>' +
          '<div class="stack-col" style="display: inline-block; width: 100%; max-width: 178px; vertical-align: top; margin: 4px; box-sizing: border-box; text-align: center; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 14px 10px;">' +
            '<div class="mobile-metric-val" style="font-size: 20px; font-weight: 800; color: #059669; line-height: 1.1;">100%</div>' +
            '<div style="font-size: 10px; font-weight: 700; color: #0f172a; margin-top: 4px; text-transform: uppercase;">Strike-Free IR</div>' +
            '<div style="font-size: 10px; color: #64748b; margin-top: 2px;">Zero Labour Stoppages</div>' +
          '</div>' +
          '<div class="stack-col" style="display: inline-block; width: 100%; max-width: 178px; vertical-align: top; margin: 4px; box-sizing: border-box; text-align: center; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 14px 10px;">' +
            '<div class="mobile-metric-val" style="font-size: 20px; font-weight: 800; color: #7c3aed; line-height: 1.1;">0-to-1</div>' +
            '<div style="font-size: 10px; font-weight: 700; color: #0f172a; margin-top: 4px; text-transform: uppercase;">PMS Rollout</div>' +
            '<div style="font-size: 10px; color: #64748b; margin-top: 2px;">200 Managers with KRAs</div>' +
          '</div>' +
        '</div>' +

        '<!-- Targeted Execution Playbook -->' +
        '<div style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 8px;">' +
          'How I Can Support ' + t.company_name + ':' +
        '</div>' +
        '<ul style="padding-left: 20px; margin: 0 0 16px 0; font-size: 13.5px; color: #334155; line-height: 1.6;">' +
          cfg.playbookBullets.map(function(b) { return '<li style="margin-bottom: 6px;">' + b + '</li>'; }).join('') +
        '</ul>' +

        '<!-- Attachment Callout -->' +
        '<div style="font-size: 12.5px; color: #334155; margin: 14px 0 16px 0; background: #f8fafc; border: 1px solid #e2e8f0; border-left: 3px solid #059669; padding: 10px 14px; border-radius: 4px;">' +
          '📎 <strong>Executive CV Attached:</strong> My complete 15-year credentials and plant workforce track record are attached as a PDF to this email.' +
        '</div>' +

        '<!-- Dual Action Buttons (Fluid-Hybrid Full Width on Mobile) -->' +
        '<div style="margin: 18px 0 16px 0; text-align: left;">' +
          '<a href="' + cvUrl + '" target="_blank" class="mobile-full-btn" style="display: inline-block; background: #0f172a; color: #ffffff; font-weight: 600; font-size: 12px; padding: 11px 18px; border-radius: 5px; text-decoration: none; margin: 4px 8px 4px 0; box-sizing: border-box;">' +
            '📄 View Full Executive CV (PDF) →' +
          '</a>' +
          '<a href="' + calendarUrl + '" target="_blank" class="mobile-full-btn" style="display: inline-block; background: #1a73e8; color: #ffffff; font-weight: 600; font-size: 12px; padding: 11px 18px; border-radius: 5px; text-decoration: none; margin: 4px 0; box-sizing: border-box;">' +
            '📅 Book on Google Calendar (15 Min) →' +
          '</a>' +
        '</div>' +

        '<p style="font-size: 13.5px; color: #475569; margin-top: 18px; line-height: 1.5;">' +
          cfg.closingAsk +
        '</p>' +

        '<!-- Executive Signoff -->' +
        '<div style="border-top: 1px solid #e2e8f0; padding-top: 16px; margin-top: 22px;">' +
          '<div style="font-weight: 700; font-size: 14px; color: #0f172a;">' + t.candidate_full_name + '</div>' +
          '<div style="font-size: 12px; color: #64748b; margin-top: 2px;">' + t.candidate_headline + '</div>' +
          '<div style="font-size: 12px; color: #64748b; margin-top: 2px;">Mobile: +91 91225 39111 &bull; Email: ankitchoudharymba@gmail.com</div>' +
          '<div style="font-size: 11px; margin-top: 6px;">' +
            '<a href="' + (t.portfolio_base_url || 'https://linkedin.com/in/ankit-choudhary-human-resources') + '" target="_blank" style="color: #0284c7; text-decoration: none; font-weight: 600;">LinkedIn Executive Profile ↗</a>' +
          '</div>' +
        '</div>' +
      '</div>' +
    '</div>';
}

/**
 * ARCHETYPE 2: Modern Operations Brief
 * Minimalist briefing memo with top accent rule, horizontal-to-vertical metrics strip,
 * and responsive action pills.
 */
function buildModernBriefEmail_(cfg, t) {
  const videoUrl = t.portfolio_video_url || t.case_study_url_utm || t.portfolio_base_url || 'https://linkedin.com/in/ankit-choudhary-human-resources';
  const cvUrl = t.resume_drive_url || (typeof ANTIGRAVITY_CONFIG !== 'undefined' ? ANTIGRAVITY_CONFIG.DEFAULT_RESUME_URL : 'https://drive.google.com/file/d/1FCyxOqXRGDnxBHugiJzRcy1OnJjbbckb/view');
  const defaultCalUrl = (typeof ANTIGRAVITY_CONFIG !== 'undefined' && ANTIGRAVITY_CONFIG.DEFAULT_GOOGLE_CALENDAR_URL)
    ? ANTIGRAVITY_CONFIG.DEFAULT_GOOGLE_CALENDAR_URL
    : 'https://calendar.google.com/calendar/render?action=TEMPLATE&text=HR+Transformation+%26+Operations+Discussion+-+Ankit+Choudhary&details=15-minute+discussion+on+manufacturing+workforce+strategy%2C+PMS+and+plant+operations.+Candidate%3A+Ankit+Choudhary+(ankitchoudharymba%40gmail.com)&add=ankitchoudharymba@gmail.com';
  const calendarUrl = (t.calendar_booking_url && !t.calendar_booking_url.includes('ankitchoudharymba/hr-leadership-discussion') && !t.calendar_booking_url.startsWith('mailto:'))
    ? t.calendar_booking_url
    : defaultCalUrl;

  return getResponsiveEmailStyles_() +
    '<div class="fluid-container" style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif; max-width: 620px; width: 100%; margin: 0 auto; background: #ffffff; color: #0f172a; line-height: 1.6; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; box-sizing: border-box;">' +
      '<div style="height: 4px; background: #0284c7; width: 100%;"></div>' +
      '<div class="fluid-padding" style="padding: 24px 24px;">' +
        '<table class="mobile-header-table" width="100%" cellpadding="0" cellspacing="0" style="margin-bottom: 20px; border-bottom: 1px solid #e2e8f0; padding-bottom: 14px;">' +
          '<tr>' +
            '<td>' +
              '<div style="font-size: 11px; font-weight: 700; color: #0284c7; text-transform: uppercase; letter-spacing: 0.08em;">Operations & Workforce Briefing</div>' +
              '<div style="font-size: 18px; font-weight: 800; color: #0f172a; margin-top: 2px;">Institutional Plant HR & Systems Rollout</div>' +
            '</td>' +
            '<td align="right" style="vertical-align: top;">' +
              '<span style="font-size: 11px; color: #64748b; background: #f1f5f9; padding: 4px 8px; border-radius: 4px; font-family: monospace;">CONFIDENTIAL</span>' +
            '</td>' +
          '</tr>' +
        '</table>' +

        '<p style="font-size: 15px; margin-top: 0; color: #0f172a;">Hi <strong>' + t.first_name + '</strong>,</p>' +
        '<p style="font-size: 14px; color: #334155; line-height: 1.6;">' + cfg.openingHook + '</p>' +

        '<!-- Responsive Horizontal/Stacked Metrics Strip -->' +
        '<div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 14px 12px; margin: 18px 0; text-align: center; font-size: 0;">' +
          '<div class="stack-col" style="display: inline-block; width: 100%; max-width: 180px; vertical-align: top; padding: 6px 10px; box-sizing: border-box; text-align: left;">' +
            '<div style="font-size: 10.5px; font-weight: 700; color: #64748b; text-transform: uppercase;">WORKFORCE SCALE</div>' +
            '<div class="mobile-metric-val" style="font-size: 16px; font-weight: 800; color: #0f172a; margin-top: 2px;">3,850 Personnel</div>' +
            '<div style="font-size: 11px; color: #64748b;">Steel & Foundry Plants</div>' +
          '</div>' +
          '<div class="stack-col" style="display: inline-block; width: 100%; max-width: 180px; vertical-align: top; padding: 6px 10px; box-sizing: border-box; text-align: left;">' +
            '<div style="font-size: 10.5px; font-weight: 700; color: #64748b; text-transform: uppercase;">SHOPFLOOR IR</div>' +
            '<div class="mobile-metric-val" style="font-size: 16px; font-weight: 800; color: #059669; margin-top: 2px;">100% Strike-Free</div>' +
            '<div style="font-size: 11px; color: #64748b;">Zero Labour Stoppages</div>' +
          '</div>' +
          '<div class="stack-col" style="display: inline-block; width: 100%; max-width: 180px; vertical-align: top; padding: 6px 10px; box-sizing: border-box; text-align: left;">' +
            '<div style="font-size: 10.5px; font-weight: 700; color: #64748b; text-transform: uppercase;">APPRAISAL / PMS</div>' +
            '<div class="mobile-metric-val" style="font-size: 16px; font-weight: 800; color: #0284c7; margin-top: 2px;">0-to-1 In-House</div>' +
            '<div style="font-size: 11px; color: #64748b;">200 Plant Managers</div>' +
          '</div>' +
        '</div>' +

        '<p style="font-size: 14px; color: #334155; line-height: 1.6;">' +
          'I summarized the operational playbook we used at Bihar Foundry & Castings Ltd into a concise 3-minute video portfolio and documented case study:' +
        '</p>' +

        '<!-- Attachment Notice -->' +
        '<div style="font-size: 12.5px; color: #334155; margin: 12px 0 16px 0; background: #f8fafc; border: 1px solid #e2e8f0; border-left: 3px solid #0284c7; padding: 10px 14px; border-radius: 4px;">' +
          '📎 <strong>Executive CV Attached:</strong> My full credentials and verified plant track record are attached as a PDF to this email.' +
        '</div>' +

        '<div style="margin: 14px 0 20px 0;">' +
          '<a href="' + videoUrl + '" target="_blank" class="mobile-full-btn" style="display: inline-block; background: #f1f5f9; border: 1px solid #cbd5e1; color: #0f172a; font-size: 12.5px; font-weight: 600; padding: 9px 15px; border-radius: 6px; text-decoration: none; margin: 4px 8px 4px 0; box-sizing: border-box;">' +
            '▶ Watch Video Walkthrough (3 Mins)' +
          '</a>' +
          '<a href="' + cvUrl + '" target="_blank" class="mobile-full-btn" style="display: inline-block; background: #0f172a; color: #ffffff; font-size: 12.5px; font-weight: 600; padding: 9px 15px; border-radius: 6px; text-decoration: none; margin: 4px 8px 4px 0; box-sizing: border-box;">' +
            '📄 Download CV (PDF)' +
          '</a>' +
          '<a href="' + calendarUrl + '" target="_blank" class="mobile-full-btn" style="display: inline-block; background: #1a73e8; color: #ffffff; font-size: 12.5px; font-weight: 600; padding: 9px 15px; border-radius: 6px; text-decoration: none; margin: 4px 0; box-sizing: border-box;">' +
            '📅 Book on Google Calendar (15 Min)' +
          '</a>' +
        '</div>' +

        '<p style="font-size: 14px; color: #334155; line-height: 1.6;">' +
          cfg.closingAsk +
        '</p>' +

        '<div style="margin-top: 24px; padding-top: 16px; border-top: 1px solid #e2e8f0;">' +
          '<div style="font-weight: 700; font-size: 14px; color: #0f172a;">' + t.candidate_full_name + '</div>' +
          '<div style="font-size: 12.5px; color: #64748b;">' + t.candidate_headline + '</div>' +
          '<div style="font-size: 12px; color: #64748b; margin-top: 4px;">' +
            'Mobile: +91 91225 39111 &bull; Email: ankitchoudharymba@gmail.com &bull; <a href="' + (t.portfolio_base_url || 'https://linkedin.com/in/ankit-choudhary-human-resources') + '" style="color: #0284c7; text-decoration: none;">LinkedIn</a>' +
          '</div>' +
        '</div>' +
      '</div>' +
    '</div>';
}

/**
 * ARCHETYPE 3: Video Spotlight Showcase
 * Media-forward layout featuring an embedded video simulation player card hero,
 * verified metric scorecard, and direct CV download.
 */
function buildVideoSpotlightEmail_(cfg, t) {
  const videoUrl = t.portfolio_video_url || t.case_study_url_utm || t.portfolio_base_url || 'https://linkedin.com/in/ankit-choudhary-human-resources';
  const cvUrl = t.resume_drive_url || (typeof ANTIGRAVITY_CONFIG !== 'undefined' ? ANTIGRAVITY_CONFIG.DEFAULT_RESUME_URL : 'https://drive.google.com/file/d/1FCyxOqXRGDnxBHugiJzRcy1OnJjbbckb/view');
  const defaultCalUrl = (typeof ANTIGRAVITY_CONFIG !== 'undefined' && ANTIGRAVITY_CONFIG.DEFAULT_GOOGLE_CALENDAR_URL)
    ? ANTIGRAVITY_CONFIG.DEFAULT_GOOGLE_CALENDAR_URL
    : 'https://calendar.google.com/calendar/render?action=TEMPLATE&text=HR+Transformation+%26+Operations+Discussion+-+Ankit+Choudhary&details=15-minute+discussion+on+manufacturing+workforce+strategy%2C+PMS+and+plant+operations.+Candidate%3A+Ankit+Choudhary+(ankitchoudharymba%40gmail.com)&add=ankitchoudharymba@gmail.com';
  const calendarUrl = (t.calendar_booking_url && !t.calendar_booking_url.includes('ankitchoudharymba/hr-leadership-discussion') && !t.calendar_booking_url.startsWith('mailto:'))
    ? t.calendar_booking_url
    : defaultCalUrl;

  return getResponsiveEmailStyles_() +
    '<div class="fluid-container" style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif; max-width: 620px; width: 100%; margin: 0 auto; background: #ffffff; color: #1e293b; line-height: 1.6; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; box-sizing: border-box;">' +
      '<div class="fluid-padding" style="padding: 24px 24px;">' +
        '<p style="font-size: 15px; margin-top: 0; color: #0f172a;">Hi <strong>' + t.first_name + '</strong>,</p>' +
        '<p style="font-size: 14px; color: #334155; line-height: 1.6;">' + cfg.openingHook + '</p>' +

        '<!-- Video Spotlight Player Card -->' +
        '<div style="background: #0f172a; border-radius: 8px; overflow: hidden; margin: 20px 0; border: 1px solid #1e293b; box-shadow: 0 4px 15px rgba(0,0,0,0.1);">' +
          '<div style="padding: 24px 20px; text-align: center; background: radial-gradient(circle at center, #1e293b 0%, #0f172a 100%);">' +
            '<a href="' + videoUrl + '" target="_blank" style="display: inline-block; width: 54px; height: 54px; line-height: 54px; border-radius: 50%; background: #0284c7; color: #ffffff; font-size: 20px; text-decoration: none; text-align: center; box-shadow: 0 0 20px rgba(2, 132, 199, 0.6);">' +
              '▶' +
            '</a>' +
            '<div style="font-size: 15px; font-weight: 700; color: #ffffff; margin-top: 14px;">' +
              'Plant Workforce Turnaround & Systems Architecture' +
            '</div>' +
            '<div style="font-size: 12px; color: #94a3b8; max-width: 440px; margin: 6px auto 14px auto;">' +
              'A 3-minute executive walkthrough on how we eliminated shopfloor absenteeism, rolled out 0-to-1 PMS, and kept 100% strike-free IR across 3,850 personnel at Bihar Foundry & Castings Ltd.' +
            '</div>' +
            '<a href="' + videoUrl + '" target="_blank" class="mobile-full-btn" style="display: inline-block; background: #0284c7; color: #ffffff; font-size: 12px; font-weight: 700; padding: 9px 18px; border-radius: 4px; text-decoration: none;">' +
              '▶ Watch Video Case Study Walkthrough' +
            '</a>' +
          '</div>' +
          '<div style="background: #1e293b; padding: 10px 16px; font-size: 11px; color: #cbd5e1; text-align: center;">' +
            'Scale: 3,850 Workforce (Steel & Foundry) &bull; 3 mins &bull; No Login Required' +
          '</div>' +
        '</div>' +

        '<!-- Verified Metrics Strip -->' +
        '<div style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 8px;">' +
          'Verified Operational Results across 3,850 workforce:' +
        '</div>' +
        '<div style="text-align: center; font-size: 0; margin-bottom: 20px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 8px 4px;">' +
          '<div class="stack-col" style="display: inline-block; width: 100%; max-width: 180px; padding: 8px; box-sizing: border-box; text-align: center;">' +
            '<div class="mobile-metric-val" style="font-size: 18px; font-weight: 800; color: #0284c7;">3,850</div>' +
            '<div style="font-size: 10.5px; color: #64748b; text-transform: uppercase; font-weight: 600;">Workforce Scope</div>' +
          '</div>' +
          '<div class="stack-col" style="display: inline-block; width: 100%; max-width: 180px; padding: 8px; box-sizing: border-box; text-align: center;">' +
            '<div class="mobile-metric-val" style="font-size: 18px; font-weight: 800; color: #059669;">Zero Strikes</div>' +
            '<div style="font-size: 10.5px; color: #64748b; text-transform: uppercase; font-weight: 600;">100% IR Harmony</div>' +
          '</div>' +
          '<div class="stack-col" style="display: inline-block; width: 100%; max-width: 180px; padding: 8px; box-sizing: border-box; text-align: center;">' +
            '<div class="mobile-metric-val" style="font-size: 18px; font-weight: 800; color: #7c3aed;">200 Mgrs</div>' +
            '<div style="font-size: 10.5px; color: #64748b; text-transform: uppercase; font-weight: 600;">0-to-1 PMS Live</div>' +
          '</div>' +
        '</div>' +

        '<!-- Attachment Callout -->' +
        '<div style="font-size: 12.5px; color: #334155; margin: 12px 0 16px 0; background: #f8fafc; border: 1px solid #e2e8f0; border-left: 3px solid #059669; padding: 10px 14px; border-radius: 4px;">' +
          '📎 <strong>Executive CV Attached:</strong> My full credentials and verified plant track record are attached as a PDF to this email.' +
        '</div>' +

        '<!-- Direct CV and Booking Action -->' +
        '<div style="margin: 16px 0 20px 0;">' +
          '<a href="' + cvUrl + '" target="_blank" class="mobile-full-btn" style="display: inline-block; background: #0f172a; color: #ffffff; font-weight: 600; font-size: 12px; padding: 10px 18px; border-radius: 5px; text-decoration: none; margin: 4px 8px 4px 0; box-sizing: border-box;">' +
            '📄 Open Attached CV (PDF) →' +
          '</a>' +
          '<a href="' + calendarUrl + '" target="_blank" class="mobile-full-btn" style="display: inline-block; background: #1a73e8; color: #ffffff; font-weight: 600; font-size: 12px; padding: 10px 18px; border-radius: 5px; text-decoration: none; margin: 4px 0; box-sizing: border-box;">' +
            '📅 Book on Google Calendar (15 Min)' +
          '</a>' +
        '</div>' +

        '<p style="font-size: 13.5px; color: #475569; margin-top: 18px; line-height: 1.5;">' +
          cfg.closingAsk +
        '</p>' +

        '<div style="border-top: 1px solid #e2e8f0; padding-top: 16px; margin-top: 22px;">' +
          '<div style="font-weight: 700; font-size: 14px; color: #0f172a;">' + t.candidate_full_name + '</div>' +
          '<div style="font-size: 12px; color: #64748b; margin-top: 2px;">' + t.candidate_headline + '</div>' +
          '<div style="font-size: 12px; color: #64748b; margin-top: 2px;">Mobile: +91 91225 39111 &bull; Email: ankitchoudharymba@gmail.com</div>' +
          '<div style="font-size: 11px; margin-top: 4px;">' +
            '<a href="' + (t.portfolio_base_url || 'https://linkedin.com/in/ankit-choudhary-human-resources') + '" target="_blank" style="color: #0284c7; text-decoration: none;">LinkedIn Executive Profile ↗</a>' +
          '</div>' +
        '</div>' +
      '</div>' +
    '</div>';
}

/**
 * ARCHETYPE 4: Ultra-Human 1-to-1 Note
 * Pure clean email text with zero outer cards or banners.
 * Looks 100% like a personal note typed from Gmail or Apple Mail with clean hyperlinks.
 */
function buildUltraHumanEmail_(cfg, t) {
  const videoUrl = t.portfolio_video_url || t.case_study_url_utm || t.portfolio_base_url || 'https://linkedin.com/in/ankit-choudhary-human-resources';
  const cvUrl = t.resume_drive_url || (typeof ANTIGRAVITY_CONFIG !== 'undefined' ? ANTIGRAVITY_CONFIG.DEFAULT_RESUME_URL : 'https://drive.google.com/file/d/1FCyxOqXRGDnxBHugiJzRcy1OnJjbbckb/view');
  const defaultCalUrl = (typeof ANTIGRAVITY_CONFIG !== 'undefined' && ANTIGRAVITY_CONFIG.DEFAULT_GOOGLE_CALENDAR_URL)
    ? ANTIGRAVITY_CONFIG.DEFAULT_GOOGLE_CALENDAR_URL
    : 'https://calendar.google.com/calendar/render?action=TEMPLATE&text=HR+Transformation+%26+Operations+Discussion+-+Ankit+Choudhary&details=15-minute+discussion+on+manufacturing+workforce+strategy%2C+PMS+and+plant+operations.+Candidate%3A+Ankit+Choudhary+(ankitchoudharymba%40gmail.com)&add=ankitchoudharymba@gmail.com';
  const calendarUrl = (t.calendar_booking_url && !t.calendar_booking_url.includes('ankitchoudharymba/hr-leadership-discussion') && !t.calendar_booking_url.startsWith('mailto:'))
    ? t.calendar_booking_url
    : defaultCalUrl;

  return '<div style="font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif; max-width: 600px; width: 100%; margin: 0 auto; background: #ffffff; color: #222222; line-height: 1.65; padding: 12px 6px; box-sizing: border-box;">' +
    '<p style="font-size: 15px; margin-top: 0;">Hi <strong>' + t.first_name + '</strong>,</p>' +
    '<p style="font-size: 14px; margin-bottom: 14px;">' + cfg.openingHook + '</p>' +
    '<p style="font-size: 14px; margin-bottom: 14px;">' +
      'To make this concrete, I put together a short <strong>3-minute video walkthrough</strong> breaking down how we structured shopfloor systems, curbed absenteeism, rolled out 0-to-1 PMS, and maintained 100% strike-free IR across 3,850 personnel at Bihar Foundry & Castings Ltd.' +
    '</p>' +
    '<p style="font-size: 14px; margin-bottom: 16px;">' +
      '👉 <strong><a href="' + videoUrl + '" style="color: #0066cc; text-decoration: underline;">Watch the 3-minute video portfolio & turnaround walkthrough here</a></strong>' +
    '</p>' +
    '<p style="font-size: 14px; margin-bottom: 14px;">' +
      'Key plant initiatives implemented:' +
    '</p>' +
    '<ul style="padding-left: 20px; font-size: 13.5px; line-height: 1.6; margin-bottom: 16px; color: #333333;">' +
      cfg.playbookBullets.map(function(b) { return '<li style="margin-bottom: 6px;">' + b + '</li>'; }).join('') +
    '</ul>' +
    '<p style="font-size: 14px; margin-bottom: 16px;">' +
      '📎 <strong>Executive CV Attached:</strong> I have attached my complete CV (PDF) to this email for your offline review (and you can also view it online here: 📄 <a href="' + cvUrl + '" style="color: #0066cc; text-decoration: underline;">View Full Executive CV (PDF)</a>).' +
    '</p>' +
    '<p style="font-size: 14px; margin-bottom: 18px;">' +
      cfg.closingAsk +
    '</p>' +
    '<p style="font-size: 14px; margin-bottom: 18px;">' +
      '📅 If you are open to a brief introductory conversation, feel free to reply with a time that works for you or book directly on Google Calendar here: <strong><a href="' + calendarUrl + '" style="color: #1a73e8; text-decoration: underline;">Book on Google Calendar (15 Min)</a></strong>' +
    '</p>' +
    '<p style="font-size: 14px; margin-bottom: 24px;">' +
      'Best regards,' +
    '</p>' +
    '<div style="font-size: 13.5px; line-height: 1.5; color: #333333;">' +
      '<strong>' + t.candidate_full_name + '</strong><br>' +
      'Head of HR & Industrial Relations<br>' +
      '15+ Years Leadership • Heavy Manufacturing & Plant Operations<br>' +
      'Mobile: <a href="tel:+919122539111" style="color: #333333; text-decoration: none;">+91 91225 39111</a> | Email: <a href="mailto:ankitchoudharymba@gmail.com" style="color: #0066cc; text-decoration: none;">ankitchoudharymba@gmail.com</a><br>' +
      'LinkedIn: <a href="' + (t.portfolio_base_url || 'https://linkedin.com/in/ankit-choudhary-human-resources') + '" style="color: #0066cc; text-decoration: underline;">LinkedIn Executive Profile</a>' +
    '</div>' +
  '</div>';
}

/**
 * Dispatcher helper to build the requested responsive executive HTML email layout.
 */
function buildExecutiveHtmlEmail_(cfg, t, archetype) {
  const arch = String(archetype || 'CLASSIC_LETTER').toUpperCase();
  switch (arch) {
    case 'MODERN_BRIEF':
      return buildModernBriefEmail_(cfg, t);
    case 'VIDEO_SPOTLIGHT':
      return buildVideoSpotlightEmail_(cfg, t);
    case 'ULTRA_HUMAN':
      return buildUltraHumanEmail_(cfg, t);
    case 'CLASSIC_LETTER':
    default:
      return buildClassicLetterEmail_(cfg, t);
  }
}

/**
 * Renders persona-aligned subject line and HTML body across the 6 template families and chosen archetype.
 */
function renderPersonaTemplate_(persona, t, archetype) {
  let subject = '';
  const p = String(persona || 'HIRING_MGR').toUpperCase();
  let cfg = { openingHook: '', playbookBullets: [], closingAsk: '' };

  switch (p) {
    case 'FOUNDER':
      subject = 'Plant HR & workforce systems — Ankit Choudhary (' + t.company_name + ')';
      cfg.openingHook = 'I\'ve been following ' + t.company_name + '\'s expansion in manufacturing operations. As growing manufacturing firms scale production capacity, setting up disciplined plant systems — biometric attendance governance, structured KRA/KPI appraisals, and contractor compliance — becomes essential to sustaining plant margins and eliminating promoter bottlenecks.';
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
      cfg.openingHook = 'As growing manufacturing organizations scale across ' + t.funding_stage_headcount + ', building structured HR processes, curbing shopfloor absenteeism, and de-risking contract-labour compliance are critical to sustaining operating margins.';
      cfg.playbookBullets = [
        '<strong>' + t.project_title + ':</strong> ' + t.architecture_summary,
        '<strong>Verified Business Impact:</strong> ' + t.quantified_business_impact,
        '<strong>Turnaround Scale:</strong> Governed 3,850 personnel with 0-strike industrial peace.'
      ];
      cfg.closingAsk = 'Full case study & credentials above. Open to a 15-minute discussion this week to explore how this playbook aligns with ' + t.company_name + '.';
  }

  const htmlBody = buildExecutiveHtmlEmail_(cfg, t, archetype);
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
    buildExecutiveHtmlEmail_,
    buildClassicLetterEmail_,
    buildModernBriefEmail_,
    buildVideoSpotlightEmail_,
    buildUltraHumanEmail_,
    appendUtm_,
    computeNextSendSlot_
  };
}
