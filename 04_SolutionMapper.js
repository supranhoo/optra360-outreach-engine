/**
 * Optra360 & Zoptro Automated Outreach & Opportunity Engine
 * 
 * 04_SolutionMapper.js: Gemini AI diagnosis, 6-Stage delivery mapping & demo matching.
 */

/**
 * Maps problem statements and AI diagnoses for pending entries in Problem_Solution tab.
 * @return {Object}
 */
function mapProblemSolutions() {
  var problems = getSheetDataAsObjects(OPTRA_CONFIG.TABS.PROBLEM_SOLUTION);
  var props = PropertiesService.getScriptProperties();
  var geminiKey = props.getProperty(OPTRA_CONFIG.PROPERTY_KEYS.GEMINI_KEY);
  var activeProfile = getActiveProfile();

  var updatedCount = 0;

  for (var i = 0; i < problems.length; i++) {
    var p = problems[i];
    if (p.Diagnosis_Status !== 'Pending AI Analysis' && p.Diagnosis_Status !== '') continue;

    var diagnosis = {};

    if (geminiKey) {
      diagnosis = analyzeWithGemini_(p.Company_Name, p.Identified_Trigger, p.Raw_Problem_Signal, activeProfile, geminiKey);
    } else {
      // Deterministic rule-based fallback based on Optra360 brand strategy guidelines
      diagnosis = generateRuleBasedDiagnosis_(p.Company_Name, p.Identified_Trigger, activeProfile);
    }

    // Update Problem Solution Tab
    updateSheetRow(OPTRA_CONFIG.TABS.PROBLEM_SOLUTION, 'Company_ID', p.Company_ID, {
      Formulated_Pain_Statement: diagnosis.painStatement,
      Recommended_Solution: diagnosis.solution,
      Mapped_Delivery_Stage: diagnosis.deliveryStage,
      Recommended_Engagement: diagnosis.engagementModel,
      Matched_Demo_ID: diagnosis.demoId,
      Expected_Transformation_Impact: diagnosis.expectedImpact,
      Diagnosis_Status: 'Completed'
    });

    updatedCount++;
    if (updatedCount >= 10) break;
  }

  return {
    status: 'success',
    diagnosesCompleted: updatedCount
  };
}

/**
 * Invokes Google Gemini API to structure the diagnosis and solution mapping.
 * @private
 */
function analyzeWithGemini_(companyName, trigger, rawSignal, activeProfile, apiKey) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;

  var systemInstruction = 
    "You are an enterprise business transformation strategist for Optra360 and Zoptro. " +
    "Optra360 solves enterprise operational crises using a 6-Stage Delivery Framework: Diagnose -> Design -> Digitize -> Deploy -> Stabilize -> Govern. " +
    "Active Profile: " + activeProfile.Name + " (" + activeProfile.Mode + "). " +
    "Return a strict JSON object with these keys: " +
    "painStatement (1 concise sentence describing the root operational bottleneck), " +
    "solution (1 sentence describing how Optra360 / Zoptro transforms this), " +
    "deliveryStage (One of: Diagnose, Design, Digitize, Deploy, Stabilize, Govern), " +
    "engagementModel ('₹3–7.5L Paid Diagnostic Audit' or '₹5–12L Controlled Pilot'), " +
    "demoId (One of: DEMO-01, DEMO-02, DEMO-03, DEMO-04), " +
    "expectedImpact (e.g. '25% reduction in cycle time and zero compliance leakage within 90 days').";

  var promptText = 
    "Company: " + companyName + "\n" +
    "Identified Trigger: " + trigger + "\n" +
    "Signal: " + rawSignal + "\n" +
    "Analyze this prospect and generate the structured JSON mapping.";

  var payload = {
    contents: [
      {
        role: "user",
        parts: [{ text: systemInstruction + "\n\n" + promptText }]
      }
    ],
    generationConfig: {
      responseMimeType: "application/json"
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      var text = data.candidates[0].content.parts[0].text;
      return JSON.parse(text);
    }
  } catch (e) {
    console.warn("Gemini API call failed, falling back to rule-based: " + e.message);
  }

  return generateRuleBasedDiagnosis_(companyName, trigger, activeProfile);
}

/**
 * Rule-based fallback aligning with Optra360 Brand Strategy Deck.
 * @private
 */
function generateRuleBasedDiagnosis_(companyName, trigger, activeProfile) {
  var mappings = {
    'Process Breakdown': {
      painStatement: companyName + ' faces operational friction across production handoffs, resulting in delayed turnaround and rework.',
      solution: 'Deploy Optra360 Rapid Workflow Diagnostic followed by digitized SOP controls to enforce process predictability.',
      deliveryStage: 'Diagnose',
      engagementModel: '₹3–7.5L Paid Diagnostic Audit',
      demoId: 'DEMO-04',
      expectedImpact: 'Elimination of process handoff bottlenecks and 20% faster execution cycles.'
    },
    'Performance Stagnation': {
      painStatement: companyName + '\'s legacy appraisal and output metrics lack real-time accountability, leading to stalled plant productivity.',
      solution: 'Implement Zoptro Performance Management Framework with cascading OKRs and Point-Factor job contribution scoring.',
      deliveryStage: 'Design',
      engagementModel: '₹5–12L Controlled Pilot',
      demoId: 'DEMO-02',
      expectedImpact: '15-25% productivity uplift and objective line-of-sight performance tracking.'
    },
    'Leadership Disconnect': {
      painStatement: 'Executive operational mandates at ' + companyName + ' are diluted before reaching shop-floor and site management.',
      solution: 'Establish an enterprise governance cadence and transparent KPI cascade connecting C-suite to front-line supervisors.',
      deliveryStage: 'Govern',
      engagementModel: '₹3–7.5L Paid Diagnostic Audit',
      demoId: 'DEMO-05',
      expectedImpact: 'Real-time daily operational visibility and zero-variance directive adherence.'
    },
    'System Fragmentation': {
      painStatement: 'Disparate spreadsheets and disconnected legacy software at ' + companyName + ' create data silos and reconciliation lag.',
      solution: 'Digitize core workflows through Zoptro Integrated HRMS/ERP cloud architecture with centralized SQL reporting.',
      deliveryStage: 'Digitize',
      engagementModel: '₹5–12L Controlled Pilot',
      demoId: 'DEMO-01',
      expectedImpact: 'Consolidation of 5+ disparate sheets into one single source of truth at ~2,500 employee scale.'
    },
    'Workforce Misalignment': {
      painStatement: companyName + ' experiences recurring staffing shortages, manual attendance leakage, and wage parity disputes.',
      solution: 'Roll out automated Biometric-SQL attendance verification and scientific Point Factor salary standardization.',
      deliveryStage: 'Deploy',
      engagementModel: '₹3–7.5L Paid Diagnostic Audit',
      demoId: 'DEMO-03',
      expectedImpact: 'Zero payroll leakage, 100% statutory compliance, and transparent grade bands.'
    }
  };

  return mappings[trigger] || mappings['Process Breakdown'];
}
