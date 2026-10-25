/**
 * PROJECT ANTIGRAVITY: END-TO-END JOB OUTREACH INTELLIGENCE PLATFORM
 * Module 6, 7 & 8: Bidirectional Response Scanner, Bounce Sensor, Pipeline Promoter & KPI Engine
 * (06_07_08_ResponseAndKPIEngine.js)
 */

/**
 * Master sync runner: Scans Gmail threads for inbound replies, bounces, and OOO notices,
 * cancels pending follow-ups, promotes positive leads to Interview_Pipeline,
 * dispatches due follow-ups, and refreshes the daily KPI metrics table.
 */
function syncGmailResponsesAndKPIs() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.warn('Sync runner locked by concurrent execution. Skipping.');
    return;
  }

  try {
    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    if (!ss) return;
    const outreachSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OUTREACH_HISTORY);
    const contactSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.CONTACTS);
    const fupSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.FOLLOW_UPS);
    const respSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.RESPONSES);
    const pipeSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.INTERVIEW_PIPELINE);
    const campSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.EMAIL_CAMPAIGNS);

    if (!outreachSheet || !contactSheet || !fupSheet || !respSheet || !pipeSheet) {
      console.error('Required CRM sheets missing. Run bootstrap first.');
      return;
    }

    const outreachData = outreachSheet.getDataRange().getValues();
    const contactsData = contactSheet.getDataRange().getValues();
    const fupData = fupSheet.getDataRange().getValues();
    const existingRespThreads = new Set(
      respSheet.getDataRange().getValues().slice(1).map(r => String(r[3]))
    );

    const contactRowMap = {};
    for (let c = 1; c < contactsData.length; c++) {
      contactRowMap[String(contactsData[c][0])] = c;
    }

    const candidateEmail = ANTIGRAVITY_CONFIG.getCandidateEmail().toLowerCase();
    let dailyBounceCount = 0;

    // PART A: Scan Active Outbound Threads in Gmail
    for (let i = 1; i < outreachData.length; i++) {
      if (String(outreachData[i][11]) !== 'SENT') continue;
      const threadId = String(outreachData[i][9] || '');
      if (!threadId || existingRespThreads.has(threadId)) continue;

      let thread;
      try {
        thread = GmailApp.getThreadById(threadId);
      } catch (e) {
        continue;
      }
      if (!thread) continue;

      const messages = thread.getMessages();
      if (messages.length <= 1) continue; // Only our initial sent message exists

      const lastMsg = messages[messages.length - 1];
      const sender = lastMsg.getFrom().toLowerCase();

      // Skip if last message was sent by the candidate
      if (sender.includes(candidateEmail)) continue;

      const bodyPlain = lastMsg.getPlainBody().slice(0, 1000);
      const subj = lastMsg.getSubject().toLowerCase();
      const contactId = String(outreachData[i][2]);
      const cIdx = contactRowMap[contactId];

      // 1. Detect Hard & Soft NDR Bounces
      if (sender.includes('mailer-daemon') || sender.includes('postmaster') || subj.includes('delivery status') || subj.includes('undelivered')) {
        outreachSheet.getRange(i + 1, 12).setValue('BOUNCED');
        if (cIdx !== undefined) contactSheet.getRange(cIdx + 1, 13).setValue('BOUNCED');
        cancelPendingFollowUpsForThread_(fupSheet, fupData, threadId, 'CANCELLED_BOUNCED');
        dailyBounceCount++;
        continue;
      }

      // 2. Detect Out-of-Office Auto-Responders
      if (subj.includes('out of office') || subj.includes('automatic reply') || bodyPlain.toLowerCase().includes('away from my desk') || bodyPlain.toLowerCase().includes('on annual leave')) {
        delayPendingFollowUps_(fupSheet, fupData, threadId, 5);
        continue;
      }

      // 3. Classify Genuine Human Reply Sentiment
      const sentiment = classifyReplySentiment_(bodyPlain);
      const sentTime = new Date(outreachData[i][8]);
      const replyTime = lastMsg.getDate();
      const hoursToReply = !isNaN(sentTime.getTime()) 
        ? Math.round(((replyTime - sentTime) / 3600000) * 10) / 10 
        : 0;

      const isPositive = (sentiment === 'POSITIVE_INTERVIEW' || sentiment === 'REFERRAL_FORWARD');

      respSheet.appendRow([
        'RSP-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
        String(outreachData[i][0]),
        contactId,
        threadId,
        replyTime.toISOString(),
        bodyPlain.slice(0, 500).replace(/\n/g, ' '),
        sentiment,
        hoursToReply,
        isPositive
      ]);

      existingRespThreads.add(threadId);

      // Cancel all remaining scheduled follow-ups immediately
      cancelPendingFollowUpsForThread_(fupSheet, fupData, threadId, 'CANCELLED_REPLIED');

      if (cIdx !== undefined) {
        const newStatus = (sentiment === 'UNSUBSCRIBE') ? 'OPT_OUT' : 'REPLIED';
        const currentScore = Number(contactsData[cIdx][11] || 70);
        const newScore = (sentiment === 'POSITIVE_INTERVIEW') ? Math.min(100, currentScore + 30) : currentScore;
        contactSheet.getRange(cIdx + 1, 12, 1, 2).setValues([[newScore, newStatus]]);

        // Promote positive reply to active interview pipeline
        if (sentiment === 'POSITIVE_INTERVIEW') {
          const pipeId = 'PIP-' + Utilities.getUuid().slice(0, 6).toUpperCase();
          const compId = contactsData[cIdx][1];
          const oppId = outreachData[i][3];

          pipeSheet.appendRow([
            pipeId,
            compId,
            oppId,
            contactId,
            '1_INTRO_CALL',
            '', // next_meeting_datetime
            '', // prep_doc_drive_url (generated by AI Copilot)
            '$210K - $250K Base + Equity',
            40 // win_probability_pct
          ]);

          // Trigger AI Interview Dossier generation if available
          try {
            if (typeof generateInterviewPrepDossier === 'function') {
              generateInterviewPrepDossier(pipeId, compId, oppId, contactId);
            }
          } catch (e) {
            console.warn('AI Dossier generation deferred: ' + e.message);
          }
        }
      }
    }

    // Circuit breaker: Pause all campaigns if bounce count exceeds safety limit
    if (dailyBounceCount >= ANTIGRAVITY_CONFIG.BOUNCE_CIRCUIT_BREAKER_LIMIT && campSheet) {
      console.warn('Circuit breaker triggered! Bounces detected: ' + dailyBounceCount + '. Pausing active campaigns.');
      const camps = campSheet.getDataRange().getValues();
      for (let k = 1; k < camps.length; k++) {
        if (camps[k][9] === 'ACTIVE') campSheet.getRange(k + 1, 10).setValue('PAUSED');
      }
    }

    // PART B: Execute Due Follow-Ups (Only runs AFTER all replies and bounces have been processed)
    executePendingFollowUps();

    // PART C: Refresh Daily KPI Metrics Snapshot
    updateDailyKpiMetrics_();

  } finally {
    lock.releaseLock();
  }
}

/**
 * Classifies email reply body into 5 distinct sentiment categories.
 */
function classifyReplySentiment_(text) {
  if (!text || typeof text !== 'string') return 'NOT_NOW';
  const lower = text.toLowerCase();

  // 1. Opt-Out / Unsubscribe
  if (lower.includes('unsubscribe') || lower.includes('remove me') || lower.includes('stop emailing') || lower.includes('take me off')) {
    return 'UNSUBSCRIBE';
  }

  // 2. Forward / Internal Referral
  if (lower.includes("cc'ing") || lower.includes('looping in') || lower.includes('speak with') || lower.includes('forwarding to') || lower.includes('connect you with')) {
    return 'REFERRAL_FORWARD';
  }

  // 3. Positive Meeting / Interview Interest
  const positiveKeywords = ['calendar', 'chat', 'call', 'interview', 'tuesday', 'thursday', 'wednesday', 'monday', 'friday', 'discuss', 'free at', 'time to connect', 'let\'s talk'];
  if (positiveKeywords.some(kw => lower.includes(kw))) {
    return 'POSITIVE_INTERVIEW';
  }

  return 'NOT_NOW';
}

/**
 * Cancels all pending follow-up rows for a specific Gmail thread.
 */
function cancelPendingFollowUpsForThread_(fupSheet, fupData, threadId, cancelStatus) {
  for (let f = 1; f < fupData.length; f++) {
    if (String(fupData[f][3]) === threadId && String(fupData[f][9]) === 'PENDING') {
      fupSheet.getRange(f + 1, 10).setValue(cancelStatus);
      fupData[f][9] = cancelStatus;
    }
  }
}

/**
 * Defers pending follow-ups for a thread by N business days (used for Out-of-Office responses).
 */
function delayPendingFollowUps_(fupSheet, fupData, threadId, extraDays) {
  for (let f = 1; f < fupData.length; f++) {
    if (String(fupData[f][3]) === threadId && String(fupData[f][9]) === 'PENDING') {
      const currentDate = new Date(fupData[f][7]);
      const pushed = addBusinessDays_(isNaN(currentDate.getTime()) ? new Date() : currentDate, extraDays);
      fupSheet.getRange(f + 1, 8).setValue(pushed.toISOString());
      fupData[f][7] = pushed.toISOString();
    }
  }
}

/**
 * Aggregates daily telemetry counts and upserts today's row in KPI_Metrics.
 */
function updateDailyKpiMetrics_() {
  const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
  if (!ss) return;
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const outreachSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OUTREACH_HISTORY);
  const fupSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.FOLLOW_UPS);
  const respSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.RESPONSES);
  const pipeSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.INTERVIEW_PIPELINE);
  const kpiSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.KPI_METRICS);

  if (!outreachSheet || !fupSheet || !respSheet || !pipeSheet || !kpiSheet) return;

  const outreach = outreachSheet.getDataRange().getValues().slice(1);
  const fups = fupSheet.getDataRange().getValues().slice(1);
  const resps = respSheet.getDataRange().getValues().slice(1);
  const pipes = pipeSheet.getDataRange().getValues().slice(1);

  const initSent = outreach.filter(r => String(r[8]).startsWith(todayStr) && r[11] === 'SENT').length;
  const fupSent = fups.filter(r => String(r[8]).startsWith(todayStr) && r[9] === 'SENT').length;
  const totalSent = initSent + fupSent;
  const bounces = outreach.filter(r => String(r[8]).startsWith(todayStr) && r[11] === 'BOUNCED').length;
  const repliesToday = resps.filter(r => String(r[4]).startsWith(todayStr)).length;
  const posToday = resps.filter(r => String(r[4]).startsWith(todayStr) && (r[6] === 'POSITIVE_INTERVIEW' || r[6] === 'REFERRAL_FORWARD')).length;
  const interviewsBooked = pipes.length;

  const totalSentAllTime = Math.max(1, outreach.filter(r => r[11] === 'SENT').length);
  const rollingReplyRate = ((resps.length / totalSentAllTime) * 100).toFixed(1) + '%';

  const kpiData = kpiSheet.getDataRange().getValues();
  let updated = false;

  for (let k = 1; k < kpiData.length; k++) {
    const rowDate = Utilities.formatDate(new Date(kpiData[k][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (rowDate === todayStr) {
      kpiSheet.getRange(k + 1, 2, 1, 9).setValues([[
        initSent,
        fupSent,
        totalSent,
        bounces,
        repliesToday,
        posToday,
        interviewsBooked,
        rollingReplyRate,
        'PRJ-001'
      ]]);
      updated = true;
      break;
    }
  }

  if (!updated) {
    kpiSheet.appendRow([
      todayStr,
      initSent,
      fupSent,
      totalSent,
      bounces,
      repliesToday,
      posToday,
      interviewsBooked,
      rollingReplyRate,
      'PRJ-001'
    ]);
  }
}

// Node.js export compatibility for offline testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    classifyReplySentiment_
  };
}
