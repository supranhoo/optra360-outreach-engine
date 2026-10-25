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
  // Table 1: Candidate_Profile
  const candSheet = ss.getSheetByName('Candidate_Profile');
  if (candSheet && candSheet.getLastRow() <= 1) {
    candSheet.appendRow([
      'CAND-01',
      'Aarav Mehta',
      'Staff Distributed Systems & AI Infrastructure Engineer',
      8.5,
      '["LLM Latency Reduction", "High-Throughput Event Pipelines", "Cloud Cost FinOps"]',
      'https://aaravmehta.dev',
      'https://github.com/aarav-systems',
      'https://linkedin.com/in/aaravmehta-eng',
      '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms',
      'https://cal.com/aaravmehta/15min-arch-chat',
      100
    ]);
  }

  // Table 2: Projects
  const projSheet = ss.getSheetByName('Projects');
  if (projSheet && projSheet.getLastRow() <= 1) {
    projSheet.appendRow([
      'PRJ-001',
      'CAND-01',
      'Sub-50ms Semantic RAG Caching Engine',
      'GITHUB_REPO',
      'llm-latency, rag, vector-db, redis, cost-optimization',
      'Python, FastAPI, Qdrant, Redis, Docker, Kubernetes',
      'Built a semantic vector cache layer in Rust/Python that intercepts repetitive LLM queries before hitting upstream foundation models.',
      'Cut OpenAI API spend by 64% ($14,200/mo savings) and reduced P99 latency from 2.4s to 38ms.',
      'https://github.com/aarav-systems/semantic-rag-cache',
      'https://aaravmehta.dev/case-studies/rag-cache-roi',
      'CTO'
    ]);
    projSheet.appendRow([
      'PRJ-002',
      'CAND-01',
      'Idempotent Event Replay & Dead-Letter Orchestrator',
      'CASE_STUDY',
      'event-streaming, kafka, high-concurrency, resilience',
      'Go, Apache Kafka, gRPC, PostgreSQL, Prometheus',
      'Engineered an enterprise transactional outbox and zero-loss consumer retry topology for high-throughput payment settlement.',
      'Maintained 99.996% uptime across 120M daily events with sub-second dead-letter replay.',
      'https://github.com/aarav-systems/kafka-deadletter-replay',
      'https://aaravmehta.dev/case-studies/kafka-settlement-scale',
      'TECH_LEAD'
    ]);
    projSheet.appendRow([
      'PRJ-003',
      'CAND-01',
      'Automated Multi-Cloud FinOps Kubernetes Engine',
      'LIVE_SAAS',
      'cloud-cost, finops, kubernetes, spot-orchestration',
      'Go, Kubernetes Operator, AWS Spot, GCP Preemptible, Datadog',
      'Engineered predictive autoscaler dynamically scheduling stateless ML inference across spot instances.',
      'Reduced monthly cloud infrastructure compute bills by $48,000 without single SLA breach.',
      'https://github.com/aarav-systems/k8s-spot-finops',
      'https://aaravmehta.dev/case-studies/k8s-finops',
      'FOUNDER'
    ]);
  }

  // Table 3: Work_Experience
  const expSheet = ss.getSheetByName('Work_Experience');
  if (expSheet && expSheet.getLastRow() <= 1) {
    expSheet.appendRow([
      'EXP-001',
      'CAND-01',
      'FinScale Payments Ltd.',
      'Lead Backend & Platform Engineer',
      'Fintech / High-Frequency Payments',
      'Legacy monolithic settlement service suffered 4.2% timeout rates during peak UPI traffic spikes.',
      'Re-architected settlement pipeline into idempotent Kafka event consumers with automated dead-letter replay.',
      'Processed $1.8B annualized volume at 99.996% reliability, recovering $420K in previously dropped annual transaction fees.',
      'https://finscale.io/blog/kafka-settlement-scale'
    ]);
  }

  // Table 4: Skills
  const sklSheet = ss.getSheetByName('Skills');
  if (sklSheet && sklSheet.getLastRow() <= 1) {
    sklSheet.appendRow([
      'SKL-001',
      'Distributed Event Streaming (Apache Kafka)',
      'BACKEND',
      10,
      'PRJ-002',
      'EXP-001',
      'kafka, redpanda, event-driven, pubsub, kinesis, streaming'
    ]);
    sklSheet.appendRow([
      'SKL-002',
      'Vector Search & RAG Latency Optimization',
      'AI_ML',
      10,
      'PRJ-001',
      'EXP-001',
      'rag, qdrant, vector-db, embeddings, redis, llm-latency, semantic-cache'
    ]);
  }

  // Table 5: Target_Companies
  const compSheet = ss.getSheetByName('Target_Companies');
  if (compSheet && compSheet.getLastRow() <= 1) {
    compSheet.appendRow([
      'CMP-101',
      'VectorFlow AI',
      'vectorflow.ai',
      'Enterprise B2B AI / Developer Tools',
      'Series B ($32M raised) | 85 Employees',
      'Launched real-time enterprise agent orchestration v2.0 last week; scaling engineering team by 40%.',
      'llm-latency, multi-tenant-isolation, vector-db, cost-optimization',
      92,
      'https://vectorflow.ai/careers'
    ]);
  }

  // Table 6: Opportunities
  const oppSheet = ss.getSheetByName('Opportunities');
  if (oppSheet && oppSheet.getLastRow() <= 1) {
    oppSheet.appendRow([
      'OPP-201',
      'CMP-101',
      'Senior Staff AI Infrastructure Engineer',
      'POSTED_JD',
      'https://boards.greenhouse.io/vectorflow/jobs/49102',
      'We need a Staff Engineer to reduce inference tail latency across our Qdrant/Redis RAG pipelines...',
      '["High P99 inference latency", "Rising foundation model token costs", "Vector index sharding"]',
      'python, fastapi, qdrant, redis, llm-latency, kubernetes',
      'PRJ-001',
      "Applying my Semantic RAG Caching pattern can cut VectorFlow's upstream LLM token spend by ~50%+ while bringing P99 query latency under 50ms."
    ]);
  }

  // Table 7: Contacts
  const cntSheet = ss.getSheetByName('Contacts');
  if (cntSheet && cntSheet.getLastRow() <= 1) {
    cntSheet.appendRow([
      'CNT-501',
      'CMP-101',
      'OPP-201',
      'Elena',
      'Rostova',
      'elena@vectorflow.ai',
      'VERIFIED_VALID',
      'TECH_LEAD',
      'VP of Engineering & AI Platform',
      'https://linkedin.com/in/elenarostova-ai',
      'America/Los_Angeles',
      88,
      'QUEUED'
    ]);
  }

  // Table 8: Email_Campaigns
  const campSheet = ss.getSheetByName('Email_Campaigns');
  if (campSheet && campSheet.getLastRow() <= 1) {
    campSheet.appendRow([
      'CMPG-01',
      'Q4 Series-B AI Infra CTO & VP Eng Cohort',
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
