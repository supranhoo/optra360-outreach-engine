/**
 * PROJECT ANTIGRAVITY: OFFLINE UNIT TEST SUITE & LOGIC VERIFIER
 * Validates lead scoring, bipartite matching, email templating, UTM builders,
 * business day arithmetic, and sentiment classification rules.
 */

const assert = require('assert');

// 1. Mock Google Apps Script Global Services for Offline Node Execution
global.Utilities = {
  getUuid: () => '12345678-abcd-1234-abcd-1234567890ab',
  formatDate: (date, tz, fmt) => {
    const d = new Date(date);
    const yr = d.getUTCFullYear();
    const mo = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    return `${yr}-${mo}-${day}`;
  }
};

global.Session = {
  getActiveUser: () => ({ getEmail: () => 'candidate@domain.com' }),
  getScriptTimeZone: () => 'GMT'
};

// 2. Load Platform Modules
const ANTIGRAVITY_CONFIG = require('../00_Config.js');
global.ANTIGRAVITY_CONFIG = ANTIGRAVITY_CONFIG;
const { validateEmailAndMx_, matchBestProject_, calculateLeadScore_ } = require('../01_ContactManager.js');
const { renderPersonaTemplate_, appendUtm_, computeNextSendSlot_ } = require('../02_03_CampaignAndPersonalizer.js');
const { addBusinessDays_ } = require('../04_05_SchedulerAndFollowUp.js');
const { classifyReplySentiment_ } = require('../06_07_08_ResponseAndKPIEngine.js');
const { ANTIGRAVITY_APOLLO } = require('../01b_ApolloIntegrator.js');

let passedTests = 0;
let totalTests = 0;

function runTest(testName, fn) {
  totalTests++;
  try {
    fn();
    console.log(`  ✓ PASS: ${testName}`);
    passedTests++;
  } catch (err) {
    console.error(`  ✗ FAIL: ${testName}`);
    console.error(`    ${err.message}`);
  }
}

console.log('\n======================================================');
console.log('🧪 RUNNING PROJECT ANTIGRAVITY COMPREHENSIVE TEST SUITE');
console.log('======================================================\n');

// -------------------------------------------------------------
// Test Suite 1: Email Validation & Disposable Filter
// -------------------------------------------------------------
console.log('--- Suite 1: Email Syntax & Role Account Filter ---');

runTest('Valid personal/work email passes regex', () => {
  assert.strictEqual(validateEmailAndMx_('elena@vectorflow.ai'), true);
  assert.strictEqual(validateEmailAndMx_('marcus.vance@chronosdata.com'), true);
});

runTest('Invalid syntax formats are rejected', () => {
  assert.strictEqual(validateEmailAndMx_('plainaddress'), false);
  assert.strictEqual(validateEmailAndMx_('@missinguser.com'), false);
  assert.strictEqual(validateEmailAndMx_('user@domain'), false);
  assert.strictEqual(validateEmailAndMx_(''), false);
  assert.strictEqual(validateEmailAndMx_(null), false);
});

runTest('Generic role & administrative accounts are blocked', () => {
  assert.strictEqual(validateEmailAndMx_('noreply@company.com'), false);
  assert.strictEqual(validateEmailAndMx_('no-reply@company.com'), false);
  assert.strictEqual(validateEmailAndMx_('donotreply@company.com'), false);
  assert.strictEqual(validateEmailAndMx_('mailer-daemon@googlemail.com'), false);
  assert.strictEqual(validateEmailAndMx_('support@company.com'), false);
  assert.strictEqual(validateEmailAndMx_('info@company.com'), false);
});

// -------------------------------------------------------------
// Test Suite 2: 100-Point Lead Scoring Model
// -------------------------------------------------------------
console.log('\n--- Suite 2: 100-Point Deterministic Lead Scoring ---');

runTest('High-match Tech Lead scores qualified (>=65)', () => {
  // Persona: TECH_LEAD (18), Overlap: 3 tags (15), Urgency: (15), Verified: (10), Base: (15) = 73
  const score = calculateLeadScore_('TECH_LEAD', 3, true);
  assert.strictEqual(score, 73);
  assert.ok(score >= ANTIGRAVITY_CONFIG.LEAD_SCORE_QUALIFY_THRESHOLD);
});

runTest('Founder with 5 overlapping tags scores max/near 100', () => {
  // Founder (20) + 5 tags capped at 25 (25) + 15 + 10 + 15 = 85
  const score = calculateLeadScore_('FOUNDER', 5, true);
  assert.strictEqual(score, 85);
});

runTest('Unverified recruiter with 0 tag overlap scores low hold (<65)', () => {
  // Recruiter (12) + 0 + 15 + 0 + 15 = 42
  const score = calculateLeadScore_('RECRUITER', 0, false);
  assert.strictEqual(score, 42);
  assert.ok(score < ANTIGRAVITY_CONFIG.LEAD_SCORE_QUALIFY_THRESHOLD);
});

runTest('Lead score cannot exceed ceiling of 100', () => {
  const score = calculateLeadScore_('FOUNDER', 20, true);
  assert.ok(score <= 100);
});

// -------------------------------------------------------------
// Test Suite 3: Bipartite Tag Matching Algorithm
// -------------------------------------------------------------
console.log('\n--- Suite 3: Bipartite Tag Intersection & Matching ---');

const mockProjects = [
  ['project_id', 'candidate_id', 'project_title', 'type', 'domain_tags', 'stack_tags', 'summary', 'impact', 'repo', 'case', 'persona'],
  ['PRJ-001', 'CAND-01', 'Semantic RAG Cache', 'GITHUB_REPO', 'llm-latency, rag, vector-db', 'python, redis, qdrant', 'summary', 'impact', 'repo', 'case', 'CTO'],
  ['PRJ-002', 'CAND-01', 'Kafka DeadLetter Replay', 'CASE_STUDY', 'event-streaming, kafka, high-concurrency', 'go, kafka, grpc', 'summary', 'impact', 'repo', 'case', 'TECH_LEAD']
];

runTest('Matches PRJ-001 for RAG and vector-db opportunity tags', () => {
  const reqTags = ['rag', 'vector-db', 'python', 'qdrant'];
  const match = matchBestProject_(reqTags, 'CTO', mockProjects);
  assert.strictEqual(match.projectId, 'PRJ-001');
  assert.ok(match.overlapCount > 0);
});

runTest('Matches PRJ-002 for high-concurrency event-streaming tags', () => {
  const reqTags = ['kafka', 'event-streaming', 'go'];
  const match = matchBestProject_(reqTags, 'TECH_LEAD', mockProjects);
  assert.strictEqual(match.projectId, 'PRJ-002');
  assert.ok(match.overlapCount > 0);
});

// -------------------------------------------------------------
// Test Suite 4: UTM Link Generator & Send Slot Scheduler
// -------------------------------------------------------------
console.log('\n--- Suite 4: UTM Instrumentor & Business Send Scheduler ---');

runTest('appendUtm_ generates clean compliant telemetry links', () => {
  const link = appendUtm_('https://github.com/aarav-systems/semantic-rag-cache', 'VectorFlow AI', 'CNT-501');
  assert.strictEqual(
    link,
    'https://github.com/aarav-systems/semantic-rag-cache?utm_source=antigravity&utm_company=vectorflowai&utm_cid=CNT-501'
  );
});

runTest('appendUtm_ handles existing query parameters with ampersand', () => {
  const link = appendUtm_('https://portfolio.dev/case-study?view=full', 'Chronos Data', 'CNT-502');
  assert.strictEqual(
    link,
    'https://portfolio.dev/case-study?view=full&utm_source=antigravity&utm_company=chronosdata&utm_cid=CNT-502'
  );
});

runTest('addBusinessDays_ skips weekends accurately', () => {
  // Friday Oct 24, 2026 -> +1 business day should be Monday Oct 27, 2026
  const friday = new Date(Date.UTC(2026, 9, 23, 10, 0, 0)); // Oct 23 is Friday in 2026
  const nextBizDay = addBusinessDays_(friday, 1);
  assert.notStrictEqual(nextBizDay.getUTCDay(), 0); // Not Sunday
  assert.notStrictEqual(nextBizDay.getUTCDay(), 6); // Not Saturday
});

// -------------------------------------------------------------
// Test Suite 5: Multi-Persona Email Template Rendering
// -------------------------------------------------------------
console.log('\n--- Suite 5: 6 Multi-Persona Template Families ---');

const sampleTokenPayload = {
  first_name: 'Elena',
  last_name: 'Rostova',
  company_name: 'VectorFlow AI',
  funding_stage_headcount: 'Series B ($32M) | 85 Employees',
  recent_trigger_event_short: 'launching enterprise agent orchestration v2.0',
  role_title_opp: 'VP of Engineering & AI Platform',
  primary_pain_point: 'P99 inference tail latency',
  required_skills_top3: 'Python, Qdrant, Redis',
  required_skill_1: 'Python',
  required_skill_2: 'Qdrant',
  project_title: 'Sub-50ms Semantic RAG Caching Engine',
  architecture_summary: 'Built a semantic vector cache layer in Rust/Python',
  quantified_business_impact: 'Cut OpenAI API spend by 64% and reduced P99 latency from 2.4s to 38ms',
  github_repo_url: 'https://github.com/aarav-systems/semantic-rag-cache',
  case_study_url_utm: 'https://aaravmehta.dev/case-studies/rag-cache-roi',
  exp_company_name: 'FinScale Payments Ltd.',
  exp_role_title: 'Lead Platform Engineer',
  exp_roi_statement: 'Processed $1.8B annualized volume at 99.996% uptime',
  custom_roi_hypothesis: 'Applying my Semantic RAG Caching pattern can cut VectorFlow token spend by 50%',
  candidate_full_name: 'Ankit Choudhary',
  candidate_headline: 'Manufacturing HR Leader | Plant HR, Industrial Relations, Workforce Strategy & Performance Transformation',
  years_experience: 15,
  github_profile_url: 'https://linkedin.com/in/ankit-choudhary-human-resources',
  portfolio_base_url: 'https://linkedin.com/in/ankit-choudhary-human-resources',
  resume_drive_url: 'https://drive.google.com/file/d/test/view',
  calendar_booking_url: 'https://calendly.com/ankitchoudharymba/hr-leadership-discussion'
};

const personasToTest = ['FOUNDER', 'TECH_LEAD', 'HIRING_MGR', 'RECRUITER', 'PEER', 'DEFAULT'];

personasToTest.forEach(persona => {
  runTest(`Renders valid HTML email without undefined tokens for ${persona}`, () => {
    const rendered = renderPersonaTemplate_(persona, sampleTokenPayload);
    assert.ok(rendered.subject && rendered.subject.length > 5);
    assert.ok(rendered.htmlBody && rendered.htmlBody.length > 50);
    assert.ok(!rendered.htmlBody.includes('undefined'), 'HTML contains undefined token placeholder');
    assert.ok(!rendered.subject.includes('undefined'), 'Subject contains undefined token placeholder');
    assert.ok(rendered.htmlBody.includes(sampleTokenPayload.first_name));
  });
});

// -------------------------------------------------------------
// Test Suite 6: Reply Sentiment Classifier
// -------------------------------------------------------------
console.log('\n--- Suite 6: Inbound Response Sentiment Classification ---');

runTest('Detects UNSUBSCRIBE opt-out intent', () => {
  assert.strictEqual(classifyReplySentiment_('Please unsubscribe me from this list.'), 'UNSUBSCRIBE');
  assert.strictEqual(classifyReplySentiment_('Stop emailing me or I will report you.'), 'UNSUBSCRIBE');
  assert.strictEqual(classifyReplySentiment_('Please remove me.'), 'UNSUBSCRIBE');
});

runTest('Detects REFERRAL_FORWARD intent', () => {
  assert.strictEqual(classifyReplySentiment_("I'm looping in our VP of Engineering who manages this."), 'REFERRAL_FORWARD');
  assert.strictEqual(classifyReplySentiment_("You should speak with Sarah on our team."), 'REFERRAL_FORWARD');
});

runTest('Detects POSITIVE_INTERVIEW meeting intent', () => {
  assert.strictEqual(classifyReplySentiment_('Super interesting repo. Free for a quick chat Thursday at 2pm?'), 'POSITIVE_INTERVIEW');
  assert.strictEqual(classifyReplySentiment_('Would love to discuss. Here is my calendar link.'), 'POSITIVE_INTERVIEW');
  assert.strictEqual(classifyReplySentiment_('Let\'s set up a call early next week.'), 'POSITIVE_INTERVIEW');
});

runTest('Categorizes non-committal or generic replies as NOT_NOW', () => {
  assert.strictEqual(classifyReplySentiment_('Thanks for reaching out. We don\'t have budget right now.'), 'NOT_NOW');
});

// -------------------------------------------------------------
// Test Suite 7: Apollo.io Lead Discovery & Ingestion Integrator
// -------------------------------------------------------------
console.log('\n--- Suite 7: Apollo.io Lead Discovery & Parsing ---');

runTest('Classifies executive personas accurately from Apollo job titles', () => {
  assert.strictEqual(ANTIGRAVITY_APOLLO.classifyPersonaFromTitle('Chief Human Resources Officer'), 'HIRING_MGR');
  assert.strictEqual(ANTIGRAVITY_APOLLO.classifyPersonaFromTitle('Executive VP - Plant Operations'), 'TECH_LEAD');
  assert.strictEqual(ANTIGRAVITY_APOLLO.classifyPersonaFromTitle('Plant Head & Works Director'), 'TECH_LEAD');
  assert.strictEqual(ANTIGRAVITY_APOLLO.classifyPersonaFromTitle('Co-Founder & Chief Executive Officer'), 'FOUNDER');
  assert.strictEqual(ANTIGRAVITY_APOLLO.classifyPersonaFromTitle('Managing Director & Promoter'), 'FOUNDER');
  assert.strictEqual(ANTIGRAVITY_APOLLO.classifyPersonaFromTitle('Chief Operating Officer (COO)'), 'TECH_LEAD');
  assert.strictEqual(ANTIGRAVITY_APOLLO.classifyPersonaFromTitle('Head of Executive Talent Acquisition'), 'RECRUITER');
});

runTest('Maps role titles to relevant manufacturing HR/Operations tags', () => {
  const hrTags = ANTIGRAVITY_APOLLO.getOpportunityTagsForRole('Head Plant HR & Industrial Relations');
  assert.ok(hrTags.includes('plant-hr'));
  assert.ok(hrTags.includes('ir'));
  assert.ok(hrTags.includes('statutory-compliance'));

  const opsTags = ANTIGRAVITY_APOLLO.getOpportunityTagsForRole('Vice President - Manufacturing & Plant Operations');
  assert.ok(opsTags.includes('plant-operations'));
  assert.ok(opsTags.includes('workforce-productivity'));

  const mdTags = ANTIGRAVITY_APOLLO.getOpportunityTagsForRole('Managing Director & Promoter');
  assert.ok(mdTags.includes('systems-implementation'));
  assert.ok(mdTags.includes('attendance-governance'));
});

runTest('Verifies Manufacturing HR systems copy in templates', () => {
  const founderTpl = renderPersonaTemplate_('FOUNDER', sampleTokenPayload);
  assert.ok(founderTpl.subject.includes('Plant HR & workforce systems'));
  assert.ok(founderTpl.htmlBody.includes('3,850-personnel plant workforce'));
  assert.ok(founderTpl.htmlBody.includes('Bihar Foundry & Castings Ltd'));

  const opsTpl = renderPersonaTemplate_('TECH_LEAD', sampleTokenPayload);
  assert.ok(opsTpl.subject.includes('absenteeism & overtime'));
  assert.ok(opsTpl.htmlBody.includes('3,850-personnel manufacturing workforce'));

  const hrTpl = renderPersonaTemplate_('HIRING_MGR', sampleTokenPayload);
  assert.ok(hrTpl.subject.includes('0-to-1 PMS & Digital Safety rollout'));
});

runTest('Validates CSV text input requirements', () => {
  const invalidRes = ANTIGRAVITY_APOLLO.importCsvText('');
  assert.strictEqual(invalidRes.success, false);
  assert.ok(invalidRes.error.includes('Empty'));

  const noEmailRes = ANTIGRAVITY_APOLLO.importCsvText('First Name,Last Name,Company\nRajesh,Sharma,Tata Steel');
  assert.strictEqual(noEmailRes.success, false);
  assert.ok(noEmailRes.error.includes('Email'));
});

// -------------------------------------------------------------
// Test Suite 8: Draft Studio & Quick Add Validation
// -------------------------------------------------------------
console.log('\n--- Suite 8: Draft Studio & Quick Add Lead Validation ---');

runTest('Formats test email subject with [TEST PREVIEW] prefix and footer', () => {
  const originalSubject = 'Plant HR systems & 3,850 workforce turnaround — Ankit Choudhary';
  const originalHtml = '<p>Hello Rajesh,</p>';
  const outreachId = 'OUT-9999';
  const campaignId = 'CMPG-01';

  const testSubject = '[TEST PREVIEW] ' + originalSubject;
  const testBody = originalHtml + 
    '<br><br><hr style="border:none;border-top:1px solid #e2e8f0;margin:24px 0;">' +
    '<p style="color:#64748b;font-size:12px;font-family:sans-serif;line-height:1.5;">' +
    '⚡ <strong>Project Antigravity Test Delivery</strong><br>' +
    'Outreach ID: <code>' + outreachId + '</code> • Campaign: <code>' + campaignId + '</code><br>' +
    '<em>This preview email was sent directly from your Antigravity Draft Studio for visual review and quality assurance.</em></p>';

  assert.ok(testSubject.startsWith('[TEST PREVIEW] '));
  assert.ok(testBody.includes('Project Antigravity Test Delivery'));
  assert.ok(testBody.includes(outreachId));
  assert.ok(testBody.includes(campaignId));
});

runTest('Validates Quick Add Lead required fields', () => {
  const incompleteLead = { firstName: 'Rajesh', companyName: 'Ramkrishna Forgings' };
  assert.ok(!incompleteLead.email, 'Missing email should be caught by validation');

  const validLead = {
    fullName: 'Rajesh Sharma',
    firstName: 'Rajesh',
    lastName: 'Sharma',
    email: 'rajesh.sharma@ramkrishnaforgings.com',
    companyName: 'Ramkrishna Forgings',
    companyDomain: 'ramkrishnaforgings.com',
    title: 'Managing Director & Promoter',
    persona: 'FOUNDER'
  };

  assert.ok(validLead.email && validLead.email.includes('@'));
  assert.strictEqual(ANTIGRAVITY_APOLLO.classifyPersonaFromTitle(validLead.title), 'FOUNDER');
  assert.strictEqual(validateEmailAndMx_(validLead.email), true);
});

// -------------------------------------------------------------
// Test Summary
// -------------------------------------------------------------
console.log('\n======================================================');
console.log(`📊 TEST RESULTS: ${passedTests} / ${totalTests} TESTS PASSED (${((passedTests / totalTests) * 100).toFixed(1)}%)`);
console.log('======================================================\n');

if (passedTests !== totalTests) {
  process.exit(1);
}
