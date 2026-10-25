/**
 * PROJECT ANTIGRAVITY: END-TO-END JOB OUTREACH INTELLIGENCE PLATFORM
 * Module 1: Contact Management, DNS MX Verification & 100-Point Lead Scoring
 * (01_ContactManager.js)
 */

/**
 * Main batch runner: Processes all contacts in 'NEW' status, validates MX records,
 * scores fit against candidate projects/skills, and queues high-scoring leads.
 */
function processAndScoreNewContacts() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.warn('Process locked by concurrent execution. Skipping run.');
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contactSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.CONTACTS);
    const oppSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OPPORTUNITIES);
    const projSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.PROJECTS);

    if (!contactSheet || !oppSheet || !projSheet) {
      console.error('Required CRM sheets not found. Please run Bootstrap first.');
      return;
    }

    const contacts = contactSheet.getDataRange().getValues();
    const opps = oppSheet.getDataRange().getValues();
    const projects = projSheet.getDataRange().getValues();

    if (contacts.length <= 1) return;

    const oppMap = buildLookupMap_(opps, 0); // Keyed by opportunity_id (Col A)
    const seenEmails = new Set();

    let processedCount = 0;

    for (let i = 1; i < contacts.length; i++) {
      const email = String(contacts[i][5] || '').trim().toLowerCase();
      const currentStatus = String(contacts[i][12] || '').trim();

      // Email deduplication check
      if (!email || seenEmails.has(email)) {
        if (currentStatus === 'NEW') contacts[i][12] = 'DUPLICATE_SKIPPED';
        continue;
      }
      seenEmails.add(email);

      // Only evaluate contacts in NEW state
      if (currentStatus !== 'NEW') continue;

      // 1. Email Syntax & DNS MX Verification Gate
      const isValidEmail = validateEmailAndMx_(email);
      contacts[i][6] = isValidEmail ? 'VERIFIED_VALID' : 'INVALID';

      if (!isValidEmail) {
        contacts[i][12] = 'BOUNCED';
        contacts[i][11] = 0;
        processedCount++;
        continue;
      }

      // 2. Extract Opportunity Tags & Bipartite Project Matching
      const oppRow = oppMap[contacts[i][2]] || [];
      const reqTags = String(oppRow[7] || '').toLowerCase().split(',').map(s => s.trim());
      const persona = String(contacts[i][7] || 'HIRING_MGR');

      const bestMatch = matchBestProject_(reqTags, persona, projects);
      const score = calculateLeadScore_(persona, bestMatch.overlapCount, isValidEmail);

      contacts[i][11] = score; // Column L: lead_score
      contacts[i][12] = score >= ANTIGRAVITY_CONFIG.LEAD_SCORE_QUALIFY_THRESHOLD ? 'QUEUED' : 'LOW_SCORE_HOLD';
      processedCount++;
    }

    // Single atomic batch write back to sheet
    contactSheet.getRange(1, 1, contacts.length, contacts[0].length).setValues(contacts);
    SpreadsheetApp.getActiveSpreadsheet().toast('Processed and scored ' + processedCount + ' contacts.', 'Antigravity Lead Scoring', 4);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Validates RFC 5322 email syntax, blocks disposable/system roles, and performs
 * zero-cost DNS MX record verification via Google Public DNS-over-HTTPS.
 */
function validateEmailAndMx_(email) {
  if (!email || typeof email !== 'string') return false;
  
  // RFC Syntax check
  const syntaxRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  if (!syntaxRegex.test(email)) return false;

  const parts = email.split('@');
  if (parts.length !== 2) return false;
  const user = parts[0].toLowerCase();
  const domain = parts[1].toLowerCase();

  // Blocked generic and administrative prefixes
  const blockedPrefixes = ['noreply', 'no-reply', 'donotreply', 'mailer-daemon', 'postmaster', 'support', 'info', 'billing'];
  if (blockedPrefixes.some(p => user === p || user.startsWith(p + '+'))) {
    return false;
  }

  // Google Public DNS-over-HTTPS lookup for MX records
  try {
    if (typeof UrlFetchApp !== 'undefined' && UrlFetchApp.fetch) {
      const resp = UrlFetchApp.fetch('https://dns.google/resolve?name=' + encodeURIComponent(domain) + '&type=MX', {
        muteHttpExceptions: true
      });
      const json = JSON.parse(resp.getContentText());
      return json.Status === 0 && Array.isArray(json.Answer) && json.Answer.length > 0;
    }
    // Fallback if UrlFetchApp is unavailable (e.g. offline testing)
    return true;
  } catch (e) {
    console.warn('DNS lookup exception for domain ' + domain + ': ' + e.message);
    return false;
  }
}

/**
 * Computes highest-overlap candidate project matching the opportunity required skills tags
 * using bipartite tag intersection plus persona target bonus.
 */
function matchBestProject_(reqTags, persona, projects) {
  let bestProjId = projects.length > 1 ? projects[1][0] : 'PRJ-001';
  let maxOverlap = -1;

  for (let p = 1; p < projects.length; p++) {
    const domainTags = String(projects[p][4] || '').toLowerCase().split(',').map(s => s.trim());
    const stackTags = String(projects[p][5] || '').toLowerCase().split(',').map(s => s.trim());
    const allProjTags = domainTags.concat(stackTags);

    const overlap = reqTags.filter(t => t && allProjTags.includes(t)).length;
    const personaBonus = String(projects[p][10]) === persona ? 1 : 0;

    const totalWeight = overlap * 1.5 + personaBonus;
    if (totalWeight > maxOverlap) {
      maxOverlap = totalWeight;
      bestProjId = projects[p][0];
    }
  }

  return { projectId: bestProjId, overlapCount: Math.max(0, Math.floor(maxOverlap)) };
}

/**
 * 100-Point Deterministic Lead Scoring Composite Model:
 * 1. Stack & Project Overlap (Max 25 pts)
 * 2. Decision Maker Authority (Max 20 pts)
 * 3. Company Trigger Urgency (Max 15 pts)
 * 4. Verification Confidence (Max 10 pts)
 * 5. Base Dynamic Fit (15 pts)
 * Total ceiling: 100 pts.
 */
function calculateLeadScore_(persona, overlapCount, isVerified) {
  const pWeights = ANTIGRAVITY_CONFIG.SCORING.PERSONA_WEIGHTS;
  const pScore = pWeights[persona] || 12;

  const stackScore = Math.min(
    ANTIGRAVITY_CONFIG.SCORING.STACK_MAX_POINTS,
    (overlapCount || 0) * ANTIGRAVITY_CONFIG.SCORING.STACK_POINTS_PER_TAG
  );

  const urgencyScore = ANTIGRAVITY_CONFIG.SCORING.URGENCY_MAX_POINTS; // Active opportunity
  const verScore = isVerified ? ANTIGRAVITY_CONFIG.SCORING.VERIFICATION_POINTS.VERIFIED_VALID : 0;
  const baseFit = 15;

  return Math.min(100, pScore + stackScore + urgencyScore + verScore + baseFit);
}

/**
 * Utility helper to build an in-memory hash map indexed by a specific column.
 */
function buildLookupMap_(dataArray, keyColIdx) {
  const map = {};
  if (!Array.isArray(dataArray)) return map;
  for (let i = 1; i < dataArray.length; i++) {
    const key = String(dataArray[i][keyColIdx]);
    if (key) map[key] = dataArray[i];
  }
  return map;
}

// Node.js export compatibility for offline testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    validateEmailAndMx_,
    matchBestProject_,
    calculateLeadScore_,
    buildLookupMap_
  };
}
