/**
 * PROJECT ANTIGRAVITY: END-TO-END JOB OUTREACH INTELLIGENCE PLATFORM
 * Module 0.1: Spreadsheet Bootstrapper, Relational Data Validation & UI Menu
 */

/**
 * Installs the custom Antigravity UI menu on spreadsheet open.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Antigravity Platform')
    .addItem('🛠 1. Bootstrap 14-Table CRM & Command Center', 'bootstrapAntigravityArchitecture')
    .addSeparator()
    .addItem('🎯 2. Verify Emails & Score Leads', 'processAndScoreNewContacts')
    .addItem('✍️ 3. Generate Personalized Drafts', 'generateCampaignOutreachDraftsPrompt')
    .addItem('🚀 4. Run Quota-Safe Mail Dispatcher', 'runQuotaSafeMailDispatcher')
    .addItem('📥 5. Sync Inbound Replies & Bounces', 'syncGmailResponsesAndKPIs')
    .addItem('📊 6. Refresh Daily KPI Rollups', 'updateDailyKpiMetrics_')
    .addSeparator()
    .addItem('⏰ 7. Install 15-Minute Automation Triggers', 'installPlatformTriggers')
    .addItem('🛑 8. Uninstall Automation Triggers', 'uninstallPlatformTriggers')
    .addToUi();
}

/**
 * Master Bootstrapper: Programmatically constructs the complete 14-table relational
 * schema plus the Command Center (Tab 0) with professional styling, column validations,
 * formulas, and blueprint seed records.
 */
function bootstrapAntigravityArchitecture() {
  const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
  if (!ss) {
    throw new Error('Unable to initialize master CRM spreadsheet.');
  }
  const schemas = ANTIGRAVITY_CONFIG.SCHEMAS;
  const enums = ANTIGRAVITY_CONFIG.ENUMS;

  try {
    if (SpreadsheetApp.getActiveSpreadsheet()) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Bootstrapping 14-table relational architecture...', 'Antigravity Platform', 5);
    }
  } catch (e) {}

  // 1. Build Tab 0: Command_Center
  setupCommandCenterTab_(ss);

  // 2. Build 14 Normalized Relational Tables
  const tableOrder = [
    'Candidate_Profile', 'Projects', 'Work_Experience', 'Skills',
    'Target_Companies', 'Opportunities', 'Contacts', 'Email_Campaigns',
    'Outreach_History', 'Follow_Ups', 'Responses', 'Interview_Pipeline',
    'Notes', 'KPI_Metrics'
  ];

  tableOrder.forEach(tableName => {
    let sheet = ss.getSheetByName(tableName);
    if (!sheet) {
      sheet = ss.insertSheet(tableName);
    }

    const headers = schemas[tableName];
    if (headers && headers.length > 0) {
      // Set Row 1 Headers
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      headerRange.setFontWeight('bold');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setBackground('#0F172A'); // Professional deep slate navy
      headerRange.setHorizontalAlignment('center');
      sheet.setFrozenRows(1);

      // Auto-fit or clean default width
      for (let c = 1; c <= headers.length; c++) {
        sheet.setColumnWidth(c, 160);
      }
    }
  });

  // 3. Apply Data Validation Dropdowns for ENUM Columns
  applyDataValidations_(ss, enums);

  // 4. Seed Blueprint Initial Records if Tables are Fresh
  seedBlueprintSampleData_(ss);

  try {
    if (SpreadsheetApp.getActiveSpreadsheet()) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Architecture bootstrap complete! 14 Relational Tabs + Command Center ready.', 'Success', 6);
    }
  } catch (e) {}
  console.log('Architecture bootstrap complete for spreadsheet: ' + ss.getUrl());
}

/**
 * Configures Tab 0: Command Center with KPI Scorecards, Formulas & Wireframes
 */
function setupCommandCenterTab_(ss) {
  let cc = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.COMMAND_CENTER);
  if (!cc) {
    cc = ss.insertSheet(ANTIGRAVITY_CONFIG.SHEETS.COMMAND_CENTER, 0);
  } else {
    ss.setActiveSheet(cc);
    ss.moveActiveSheet(1);
  }

  cc.clear();
  cc.setTabColor('#0F172A');
  cc.getRange('A1:L50').setFontFamily('Inter, sans-serif');

  // Title Banner
  cc.getRange('A1:L1').merge()
    .setValue('⚡ PROJECT ANTIGRAVITY: JOB OUTREACH INTELLIGENCE COMMAND CENTER')
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#FFFFFF')
    .setBackground('#0F172A')
    .setHorizontalAlignment('left');

  // Section 1: KPI Scorecards Header
  cc.getRange('A3:L3').merge()
    .setValue('EXECUTIVE TELEMETRY & CONVERSION SCORECARDS (LIVE FORMULAS)')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#64748B')
    .setBackground('#F8FAFC');

  // Scorecard Cards Grid (Row 4-5)
  const cards = [
    { col: 'A', name: "Today's Quota Used", formula: '=IFERROR(INDEX(KPI_Metrics!D2:D, COUNTA(KPI_Metrics!A2:A)), 0) & " / 100"' },
    { col: 'C', name: 'Total Contacts Emailed', formula: '=IFERROR(COUNTIF(Outreach_History!L2:L, "SENT"), 0)' },
    { col: 'E', name: 'Overall Reply Rate', formula: '=IFERROR(TEXT(COUNTA(UNIQUE(Responses!C2:C)) / COUNTIF(Outreach_History!L2:L, "SENT"), "0.0%"), "0.0%")' },
    { col: 'G', name: 'Positive Reply Rate', formula: '=IFERROR(TEXT(COUNTIFS(Responses!G2:G, "POSITIVE_INTERVIEW") / COUNTIF(Outreach_History!L2:L, "SENT"), "0.0%"), "0.0%")' },
    { col: 'I', name: 'Interviews Booked', formula: '=IFERROR(COUNTA(Interview_Pipeline!A2:A)-1, 0)' },
    { col: 'K', name: 'Active Offers (Stage 5)', formula: '=IFERROR(COUNTIF(Interview_Pipeline!E2:E, "5_OFFER"), 0)' }
  ];

  cards.forEach(card => {
    const colIdx = card.col.charCodeAt(0) - 64;
    cc.getRange(4, colIdx, 1, 2).merge()
      .setValue(card.name)
      .setFontSize(9)
      .setFontColor('#475569')
      .setBackground('#F1F5F9')
      .setHorizontalAlignment('center');

    cc.getRange(5, colIdx, 1, 2).merge()
      .setValue(card.formula)
      .setFontWeight('bold')
      .setFontSize(16)
      .setFontColor('#0F172A')
      .setBackground('#FFFFFF')
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
  });

  // Section 2: 30-Day Velocity & Proof-of-Work ROI (Row 7)
  cc.getRange('A7:F7').merge()
    .setValue('30-DAY SEND & REPLY VELOCITY (SPARKLINE)')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#64748B')
    .setBackground('#F8FAFC');

  cc.getRange('G7:L7').merge()
    .setValue('PROOF-OF-WORK ASSET ATTRIBUTION (BY PROJECT ID)')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#64748B')
    .setBackground('#F8FAFC');

  // Sparkline formula & project query
  cc.getRange('A8:F12').merge()
    .setValue('=SPARKLINE(KPI_Metrics!G2:G31, {"charttype", "line"; "color", "#2563EB"; "linewidth", 2})')
    .setBorder(true, true, true, true, false, false, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);

  cc.getRange('G8:L12').merge()
    .setValue('=IFERROR(QUERY(Outreach_History!A2:L, "SELECT E, COUNT(A) WHERE L=\'SENT\' GROUP BY E LABEL E \'Matched Project\', COUNT(A) \'Emails Sent\'", 0), "No outbound dispatches yet")')
    .setFontSize(10)
    .setBorder(true, true, true, true, false, false, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);

  // Section 3: Actionable Interview Pipeline Panel (Row 14)
  cc.getRange('A14:L14').merge()
    .setValue('🔥 ACTIONABLE INTERVIEW PIPELINE & ACTIVE OPPORTUNITIES (WARM LEADS)')
    .setFontWeight('bold')
    .setFontSize(10)
    .setFontColor('#FFFFFF')
    .setBackground('#1E293B');

  const pipeHeaders = ['Pipeline ID', 'Company ID', 'Opportunity ID', 'Contact ID', 'Current Stage', 'Next Meeting', 'Prep Dossier URL', 'Comp Target', 'Win Prob %'];
  cc.getRange(15, 1, 1, pipeHeaders.length).setValues([pipeHeaders])
    .setFontWeight('bold')
    .setFontSize(9)
    .setBackground('#E2E8F0')
    .setHorizontalAlignment('center');

  // Mirror first 10 active pipeline rows dynamically
  cc.getRange(16, 1).setValue('=IFERROR(FILTER(Interview_Pipeline!A2:I, Interview_Pipeline!A2:A<>""), "No active pipeline stages yet")');

  cc.setFrozenRows(2);
}

/**
 * Applies strict Google Sheets Data Validation rules for ENUM columns
 */
function applyDataValidations_(ss, enums) {
  const setDropdown = (sheetName, colLetter, values) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(colLetter + '2:' + colLetter + '1000').setDataValidation(rule);
  };

  setDropdown('Projects', 'D', enums.PROJECT_TYPE);
  setDropdown('Projects', 'K', enums.PERSONA_TYPE);
  setDropdown('Skills', 'C', enums.SKILL_CATEGORY);
  setDropdown('Opportunities', 'D', enums.OPPORTUNITY_TYPE);
  setDropdown('Contacts', 'G', enums.VERIFICATION_STATUS);
  setDropdown('Contacts', 'H', enums.PERSONA_TYPE);
  setDropdown('Contacts', 'M', enums.OUTREACH_STATUS);
  setDropdown('Email_Campaigns', 'C', enums.PERSONA_TYPE);
  setDropdown('Email_Campaigns', 'J', enums.CAMPAIGN_STATUS);
  setDropdown('Outreach_History', 'L', enums.DISPATCH_STATUS);
  setDropdown('Follow_Ups', 'F', enums.FOLLOWUP_ANGLE);
  setDropdown('Follow_Ups', 'J', enums.FOLLOWUP_STATUS);
  setDropdown('Responses', 'G', enums.SENTIMENT_CATEGORY);
  setDropdown('Interview_Pipeline', 'E', enums.PIPELINE_STAGE);
  setDropdown('Notes', 'B', enums.ENTITY_TYPE);
  setDropdown('Notes', 'E', enums.NOTE_AUTHOR);
}

/**
 * Seeds the platform with blueprint-compliant master records if tables are empty.
 */
function seedBlueprintSampleData_(ss) {
  // Table 1: Candidate_Profile (Ankit Choudhary - Manufacturing HR Leader)
  const candSheet = ss.getSheetByName('Candidate_Profile');
  if (candSheet && candSheet.getLastRow() <= 1) {
    candSheet.appendRow([
      'CAND-01',
      'Ankit Choudhary',
      'Manufacturing HR Leader | Plant HR, Industrial Relations, Workforce Strategy & Performance Transformation',
      15,
      '["Plant HR & Industrial Relations Governance", "Performance Management (PMS & KRA/KPI) Architecture", "Integrated HR Technology & Greenfield Expansion"]',
      'https://linkedin.com/in/ankit-choudhary-human-resources',
      'https://linkedin.com/in/ankit-choudhary-human-resources',
      'https://linkedin.com/in/ankit-choudhary-human-resources',
      '1AnkitChoudhary_Resume_DriveID',
      'https://calendly.com/ankitchoudharymba/hr-leadership-discussion',
      100
    ]);
  }

  // Table 2: Projects (Proof-of-Work, System Implementations & Turnarounds)
  const projSheet = ss.getSheetByName('Projects');
  if (projSheet && projSheet.getLastRow() <= 1) {
    projSheet.appendRow([
      'PRJ-001',
      'CAND-01',
      'Integrated Safety, PMS & LMS Digital Workforce Platform',
      'LIVE_SAAS',
      'pms-rollout, kra-kpi, lms, safety-reporting, plant-digitization, workforce-analytics',
      'React, TypeScript, Supabase, Google Workspace, Power BI, Biometric ESSL',
      'Architected an in-house integrated employee workforce platform combining Safety reporting, KRA/KPI performance appraisals, and TNI-driven training workflows with employee-wide access across 3,850 plant personnel.',
      'Rolled out PMS to 100 senior leaders in 6 months and scaled coverage to 200 managers in 9 months; established L&D division driven by real-time TNI matrices.',
      'https://linkedin.com/in/ankit-choudhary-human-resources',
      'https://linkedin.com/in/ankit-choudhary-human-resources',
      'HIRING_MGR'
    ]);
    projSheet.appendRow([
      'PRJ-002',
      'CAND-01',
      '3,850-Worker Plant Attendance, Increment & Absenteeism Governance System',
      'CASE_STUDY',
      'plant-hr, absenteeism-control, time-office, biometric-essl, industrial-relations, shift-deployment',
      'Biometric ESSL, Time Office MIS, Excel Power Query, SAP B1, Factory Compliance',
      'Restructured plant attendance governance and formulated a production-linked increment policy with automated time-office biometric verification across 2,350 on-roll and 1,500 contract workers.',
      'Curbed chronic manufacturing absenteeism, optimized overtime spend across 3 production shifts, and improved shopfloor schedule adherence.',
      'https://linkedin.com/in/ankit-choudhary-human-resources',
      'https://linkedin.com/in/ankit-choudhary-human-resources',
      'TECH_LEAD'
    ]);
    projSheet.appendRow([
      'PRJ-003',
      'CAND-01',
      'Manufacturing Employer Brand & Workforce Experience Turnaround',
      'CASE_STUDY',
      'employer-branding, employee-relations, attrition-reduction, ambitionbox, culture-transformation',
      'HR Analytics, Grievance Resolution, AmbitionBox, Google Reviews, Engagement Surveys',
      'Executed a structured employee experience and employer-brand turnaround program addressing worker welfare, transparent appraisals, and rapid grievance resolution.',
      'Elevated AmbitionBox rating from 2.5 to 4.2 and Google rating from 3.1 to 4.6 within 6 months, drastically lowering talent acquisition friction.',
      'https://linkedin.com/in/ankit-choudhary-human-resources',
      'https://linkedin.com/in/ankit-choudhary-human-resources',
      'FOUNDER'
    ]);
    projSheet.appendRow([
      'PRJ-004',
      'CAND-01',
      'Multi-Entity Greenfield Plant Expansion & Statutory Clearance Coordination',
      'CASE_STUDY',
      'greenfield-expansion, plant-infrastructure, statutory-clearance, vendor-finalization, factory-licensing',
      'Regulatory Compliance, Industry Development Authority, Vendor RFP, Civil Contracts',
      'Directed greenfield manufacturing expansion from land identification and regulatory clearance liaison through contractor selection and plant-infrastructure vendor finalization.',
      'Delivered compliant industrial infrastructure and workforce ramp-up on schedule across 8 operating group entities.',
      'https://linkedin.com/in/ankit-choudhary-human-resources',
      'https://linkedin.com/in/ankit-choudhary-human-resources',
      'FOUNDER'
    ]);
  }

  // Table 3: Work_Experience (Enterprise Credibility & Quantified Track Record)
  const expSheet = ss.getSheetByName('Work_Experience');
  if (expSheet && expSheet.getLastRow() <= 1) {
    expSheet.appendRow([
      'EXP-001',
      'CAND-01',
      'Bihar Foundry & Castings Limited',
      'Senior Manager - Human Resources',
      'Steel / Foundry Manufacturing (3,850 Workforce Scope)',
      'Managing a complex 3,850-person industrial workforce (2,350 on-roll + 1,500 contract) across Plant HR, Time Office, industrial relations, statutory compliance, and production shift availability.',
      'Led 27-member HR/IR/admin function; built in-house Safety, PMS & LMS platform; instituted revised increment policy and biometric attendance governance; scaled PMS to 200+ managers.',
      'Elevated AmbitionBox rating to 4.2/5, covered 100% eligible managerial population under structured PMS, and stabilized multi-shift workforce compliance with zero production work-stoppages.',
      'https://linkedin.com/in/ankit-choudhary-human-resources'
    ]);
    expSheet.appendRow([
      'EXP-002',
      'CAND-01',
      'GMG Group',
      'HR Manager - Group HR Responsibility',
      'Automotive, Marine & Precision Engineering (1,700 Workforce Scope)',
      'Multi-entity HR governance across 8 operating brands (Multisteer, Multiflex, Navigator, Glydus, etc.) with fragmented attendance, disparate pay grades, and greenfield plant expansion demands.',
      'Implemented HR One HRIS with biometric ESSL integration across 1,700 employees; developed group role libraries, KRAs, KPIs, and salary grade structures; managed labour-office matters and dispute closures; coordinated greenfield plant expansion.',
      'Unified group-wide workforce compliance across 8 entities and delivered greenfield plant licensing, contractor liaison, and staffing on schedule.',
      'https://linkedin.com/in/ankit-choudhary-human-resources'
    ]);
    expSheet.appendRow([
      'EXP-003',
      'CAND-01',
      'Channelplay Limited (Pan-India Samsung Visibility Project)',
      'Manager - Human Resources',
      'Agency Manpower & Retail Visibility (2,000+ Field-Force Scope)',
      'Managing Pan-India distributed workforce operations, rapid hiring pipelines, and performance appraisals for 2,000+ off-roll personnel.',
      'Spearheaded national HR governance, appraisal cycles, POSH compliance, grievance investigations, and attrition analytics across all Indian zones.',
      'Maintained sub-15-day requisition fulfillment and standard SLA compliance across 2,000+ national retail deployments.',
      'https://linkedin.com/in/ankit-choudhary-human-resources'
    ]);
  }

  // Table 4: Skills
  const sklSheet = ss.getSheetByName('Skills');
  if (sklSheet && sklSheet.getLastRow() <= 1) {
    sklSheet.appendRow([
      'SKL-001',
      'Plant HR Operations & Industrial Relations (IR)',
      'INFRA',
      10,
      'PRJ-002',
      'EXP-001',
      'plant-hr, ir, industrial-relations, factory-act, statutory-compliance, contract-labour, union-grievance, labour-office'
    ]);
    sklSheet.appendRow([
      'SKL-002',
      'Performance Management Systems (PMS & KRA/KPI)',
      'PRODUCT',
      10,
      'PRJ-001',
      'EXP-001',
      'pms, kra-kpi, appraisals, competency-matrix, tni, l-and-d, talent-management, succession-planning'
    ]);
    sklSheet.appendRow([
      'SKL-003',
      'Integrated HR Technology & Workforce Analytics',
      'AI_ML',
      10,
      'PRJ-001',
      'EXP-001',
      'hris, hr-tech, react, supabase, power-bi, biometric-essl, keka, hr-one, attendance-analytics, attrition-modelling'
    ]);
    sklSheet.appendRow([
      'SKL-004',
      'Greenfield Plant Expansion & Contractor Liaison',
      'INFRA',
      9,
      'PRJ-004',
      'EXP-002',
      'greenfield, industrial-expansion, plant-infrastructure, statutory-liaison, contractor-finalization, civil-vendor'
    ]);
    sklSheet.appendRow([
      'SKL-005',
      'Workforce Strategy & Production Manpower Planning',
      'PRODUCT',
      10,
      'PRJ-002',
      'EXP-001',
      'manpower-planning, headcount-budgeting, shift-deployment, overtime-optimization, absenteeism-control'
    ]);
  }

  // Table 5: Target_Companies (Target Manufacturing, Steel, Metals & Mining Enterprises)
  const compSheet = ss.getSheetByName('Target_Companies');
  if (compSheet && compSheet.getLastRow() <= 1) {
    compSheet.appendRow([
      'CMP-101',
      'Jindal Steel & Power Ltd. (JSPL)',
      'jindalsteelpower.com',
      'Steel & Heavy Manufacturing',
      'Enterprise (15,000+ Employees) | Raigarh & Angul Plants',
      'Expanding integrated steel capacity to 12 MTPA; scaling multi-shift manufacturing workforce and technical leadership.',
      'plant-hr, ir-governance, contract-labour, pms-scale, absenteeism-control',
      95,
      'https://jindalsteelpower.com/careers'
    ]);
    compSheet.appendRow([
      'CMP-102',
      'Vedanta Aluminium & Power',
      'vedanta.co.in',
      'Metals, Mining & Power',
      'Enterprise (20,000+ Employees) | Jharsuguda Smelter',
      'Operationalizing mega smelter expansion and modernizing workforce digital governance across Odisha/Chhattisgarh.',
      'industrial-relations, workforce-planning, greenfield-expansion, hr-tech',
      92,
      'https://vedanta.co.in/careers'
    ]);
  }

  // Table 6: Opportunities
  const oppSheet = ss.getSheetByName('Opportunities');
  if (oppSheet && oppSheet.getLastRow() <= 1) {
    oppSheet.appendRow([
      'OPP-201',
      'CMP-101',
      'Head of Plant HR / General Manager - HR & IR',
      'POSTED_JD',
      'https://jindalsteelpower.com/careers/head-plant-hr-raigarh',
      'We require a seasoned Manufacturing HR Leader to lead Plant HR operations, statutory compliance, industrial relations, and workforce productivity across our integrated steel complex...',
      '["Contract-labour compliance and union grievance de-risking", "Shopfloor absenteeism and shift deployment productivity", "Digital PMS adoption and supervisory competency building"]',
      'plant-hr, ir, statutory-compliance, pms, contract-labour, manpower-planning',
      'PRJ-001',
      'Applying my 3,850-worker Plant HR governance and in-house digital PMS/Safety framework can directly de-risk IR bottlenecks and optimize shopfloor productivity across multi-shift operations.'
    ]);
  }

  // Table 7: Contacts (Real Decision Makers)
  const cntSheet = ss.getSheetByName('Contacts');
  if (cntSheet && cntSheet.getLastRow() <= 1) {
    cntSheet.appendRow([
      'CNT-501',
      'CMP-101',
      'OPP-201',
      'Rajesh',
      'Sharma',
      'rajesh.sharma@jindalsteelpower.com',
      'VERIFIED_VALID',
      'TECH_LEAD',
      'Executive Vice President - Plant Operations & Manufacturing',
      'https://linkedin.com/in/rajesh-sharma-steel',
      'Asia/Kolkata',
      92,
      'QUEUED'
    ]);
    cntSheet.appendRow([
      'CNT-502',
      'CMP-101',
      'OPP-201',
      'Alok',
      'Verma',
      'alok.verma@jindalsteelpower.com',
      'VERIFIED_VALID',
      'HIRING_MGR',
      'Chief Human Resources Officer (CHRO)',
      'https://linkedin.com/in/alok-verma-chro',
      'Asia/Kolkata',
      94,
      'QUEUED'
    ]);
  }

  // Table 8: Email_Campaigns
  const campSheet = ss.getSheetByName('Email_Campaigns');
  if (campSheet && campSheet.getLastRow() <= 1) {
    campSheet.appendRow([
      'CMPG-01',
      'Q4 Enterprise Steel & Manufacturing Plant HR Leadership Cohort',
      'TECH_LEAD',
      'TPL-TECH-LEAD-01',
      40,
      8,
      17,
      3,
      7,
      'ACTIVE'
    ]);
  }

  // Table 14: KPI_Metrics initial seed row
  const kpiSheet = ss.getSheetByName('KPI_Metrics');
  if (kpiSheet && kpiSheet.getLastRow() <= 1) {
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    kpiSheet.appendRow([
      todayStr, 0, 0, 0, 0, 0, 0, 0, '0.0%', 'PRJ-001'
    ]);
  }
}

/**
 * Interactive prompt helper for campaign draft generation
 */
function generateCampaignOutreachDraftsPrompt() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Generate Campaign Drafts', 'Enter Campaign ID to generate drafts for (e.g., CMPG-01):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() === ui.Button.OK) {
    const campId = resp.getResponseText().trim();
    if (campId) {
      generateCampaignOutreachDrafts(campId);
    }
  }
}
