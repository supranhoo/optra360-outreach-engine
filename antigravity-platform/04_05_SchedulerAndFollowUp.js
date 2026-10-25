/**
 * PROJECT ANTIGRAVITY: END-TO-END JOB OUTREACH INTELLIGENCE PLATFORM
 * Module 4 & 5: Quota-Safe Mail Dispatcher (100/day) & Thread-Locked Follow-Up Engine
 * (04_05_SchedulerAndFollowUp.js)
 */

/**
 * Main Outbound Dispatcher:
 * Enforces strict daily 100-email ceiling, reserves safety quota buffer,
 * executes atomic row locks (SENDING_LOCKED), and dispatches up to 8 emails per 15-min cycle.
 */
function runQuotaSafeMailDispatcher() {
  const startTime = Date.now();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    console.warn('Dispatcher locked by another execution cycle. Exiting.');
    return;
  }

  try {
    const props = PropertiesService.getScriptProperties();
    const todayStr = Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd');
    const todayKey = 'SENT_COUNT_' + todayStr;
    let sentToday = Number(props.getProperty(todayKey) || 0);

    // Remaining daily quota check
    let remainingQuota = 1500;
    try {
      remainingQuota = MailApp.getRemainingDailyQuota();
    } catch (e) {
      console.warn('MailApp.getRemainingDailyQuota error: ' + e.message);
    }

    if (sentToday >= ANTIGRAVITY_CONFIG.DAILY_HARD_LIMIT || remainingQuota <= ANTIGRAVITY_CONFIG.SAFETY_QUOTA_BUFFER) {
      console.log('Daily send cap or safety quota reached. Sent today: ' + sentToday + ', Remaining: ' + remainingQuota);
      return;
    }

    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    if (!ss) return;
    const outreachSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OUTREACH_HISTORY);
    const contactSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.CONTACTS);
    const followupSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.FOLLOW_UPS);
    const campSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.EMAIL_CAMPAIGNS);

    if (!outreachSheet || !contactSheet || !followupSheet || !campSheet) {
      console.error('Required CRM sheets missing. Run bootstrap first.');
      return;
    }

    const campaigns = buildLookupMap_(campSheet.getDataRange().getValues(), 0);
    const outreachData = outreachSheet.getDataRange().getValues();
    const contactsData = contactSheet.getDataRange().getValues();

    const contactRowIdxMap = {};
    for (let c = 1; c < contactsData.length; c++) {
      contactRowIdxMap[String(contactsData[c][0])] = c;
    }

    const now = new Date();
    let batchCount = 0;

    // Fetch and prepare physical CV PDF attachment for campaign dispatches
    let attachments = [];
    let resumeDriveId = ANTIGRAVITY_CONFIG.DEFAULT_RESUME_DRIVE_ID;
    const candSheet = ss.getSheetByName('Candidate_Profile');
    if (candSheet && candSheet.getLastRow() >= 2) {
      const p = candSheet.getRange(2, 1, 1, 10).getValues()[0];
      if (p[8] && p[8].length > 10 && !p[8].includes('1AnkitChoudhary_Resume_DriveID')) {
        resumeDriveId = p[8];
      }
    }
    if (resumeDriveId && typeof DriveApp !== 'undefined') {
      try {
        const cvFile = DriveApp.getFileById(resumeDriveId);
        attachments.push(cvFile.getAs('application/pdf').setName(ANTIGRAVITY_CONFIG.DEFAULT_RESUME_FILE_NAME || 'Ankit_Choudhary_Executive_HR_CV.pdf'));
      } catch (attErr) {
        console.warn('Could not load CV attachment for campaign dispatch: ' + attErr.message);
      }
    }

    for (let i = 1; i < outreachData.length; i++) {
      // Execution ceiling & batch throttle
      if (Date.now() - startTime > ANTIGRAVITY_CONFIG.MAX_EXECUTION_MS) break;
      if (batchCount >= ANTIGRAVITY_CONFIG.BATCH_DISPATCH_LIMIT) break;
      if (sentToday >= ANTIGRAVITY_CONFIG.DAILY_HARD_LIMIT) break;

      const status = String(outreachData[i][11] || '').trim();
      const scheduledAt = new Date(outreachData[i][7]);

      if (status !== 'APPROVED_QUEUED' || isNaN(scheduledAt.getTime()) || scheduledAt > now) {
        continue;
      }

      const contactId = String(outreachData[i][2]);
      const cIdx = contactRowIdxMap[contactId];
      if (cIdx === undefined) continue;

      const recipientEmail = String(contactsData[cIdx][5] || '').trim();
      const firstName = String(contactsData[cIdx][3] || 'there');
      if (!recipientEmail) continue;

      // 1. Atomic row lock before calling Gmail MTA
      outreachSheet.getRange(i + 1, 12).setValue('SENDING_LOCKED');
      SpreadsheetApp.flush();

      const subject = String(outreachData[i][5]);
      const htmlBody = String(outreachData[i][6]);
      const plainBody = htmlBody.replace(/<[^>]+>/g, '');

      try {
        GmailApp.sendEmail(recipientEmail, subject, plainBody, { htmlBody: htmlBody, attachments: attachments });
      } catch (err) {
        console.error('Failed to send email to ' + recipientEmail + ': ' + err.message);
        outreachSheet.getRange(i + 1, 12).setValue('FAILED');
        continue;
      }

      Utilities.sleep(2500); // Allow Gmail Sent index to register thread

      // 2. Query sent thread to extract threadId and RFC 2822 messageId
      let threadId = '';
      let msgId = '';
      try {
        const query = 'to:' + recipientEmail + ' subject:"' + subject.replace(/"/g, '') + '"';
        const threads = GmailApp.search(query, 0, 1);
        if (threads.length > 0) {
          threadId = threads[0].getId();
          const msgs = threads[0].getMessages();
          if (msgs.length > 0) {
            msgId = msgs[msgs.length - 1].getId();
          }
        }
      } catch (e) {
        console.warn('Sent thread query exception: ' + e.message);
      }

      const sentIso = new Date().toISOString();
      outreachSheet.getRange(i + 1, 9, 1, 4).setValues([[sentIso, threadId, msgId, 'SENT']]);
      contactSheet.getRange(cIdx + 1, 13).setValue('IN_SEQUENCE');

      // 3. Pre-schedule Stage 1 & Stage 2 follow-ups
      const camp = campaigns[String(outreachData[i][1])] || [];
      const delay1 = Number(camp[7]) || 3;
      const delay2 = Number(camp[8]) || 7;

      if (threadId) {
        scheduleFollowUpRows_(followupSheet, String(outreachData[i][0]), contactId, threadId, firstName, delay1, delay2);
      }

      sentToday++;
      batchCount++;
      props.setProperty(todayKey, String(sentToday));

      // 4. Human randomized jitter delay
      const jitterMs = Math.floor(Math.random() * (ANTIGRAVITY_CONFIG.JITTER_MAX_MS - ANTIGRAVITY_CONFIG.JITTER_MIN_MS)) + ANTIGRAVITY_CONFIG.JITTER_MIN_MS;
      Utilities.sleep(jitterMs);
    }

    if (batchCount > 0) {
      console.log('Batch dispatch complete: ' + batchCount + ' emails sent. Total today: ' + sentToday);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Pre-schedules Stage 1 (architecture benchmark) and Stage 2 (soft breakup)
 * multi-touch follow-ups into the Follow_Ups table.
 */
function scheduleFollowUpRows_(followupSheet, outreachId, contactId, threadId, firstName, delay1Days, delay2Days) {
  if (!threadId) return;

  const now = new Date();
  const date1 = addBusinessDays_(now, delay1Days);
  const date2 = addBusinessDays_(now, delay2Days);

  const fup1Body = '<p>Hi ' + firstName + ' — bumping this to the top of your inbox in case it got buried. Happy to share a 1-page architecture benchmark showing how we achieved those latency numbers if useful for your roadmap.</p>';
  const fup2Body = '<p>Hi ' + firstName + ' — last note from my side so I don\'t clutter your inbox. If scaling this area becomes a priority next quarter, feel free to reply anytime or let me know if you\'d prefer I not follow up.</p>';

  const fup1Id = 'FUP-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  const fup2Id = 'FUP-' + Utilities.getUuid().slice(0, 8).toUpperCase();

  followupSheet.appendRow([
    fup1Id, outreachId, contactId, threadId, 1, 'ARCHITECTURE_DIAGRAM', fup1Body,
    date1.toISOString(), '', 'PENDING'
  ]);
  followupSheet.appendRow([
    fup2Id, outreachId, contactId, threadId, 2, 'SOFT_BREAKUP', fup2Body,
    date2.toISOString(), '', 'PENDING'
  ]);
}

/**
 * Executes due multi-touch follow-up emails, strictly replying inside the original Gmail thread.
 */
function executePendingFollowUps() {
  const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
  if (!ss) return;
  const fupSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.FOLLOW_UPS);
  const contactSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.CONTACTS);
  if (!fupSheet || !contactSheet) return;

  const contacts = buildLookupMap_(contactSheet.getDataRange().getValues(), 0);
  const fupData = fupSheet.getDataRange().getValues();
  const now = new Date();

  const props = PropertiesService.getScriptProperties();
  const todayStr = Utilities.formatDate(now, 'GMT', 'yyyy-MM-dd');
  const todayKey = 'SENT_COUNT_' + todayStr;
  let sentToday = Number(props.getProperty(todayKey) || 0);

  for (let i = 1; i < fupData.length; i++) {
    if (sentToday >= ANTIGRAVITY_CONFIG.DAILY_HARD_LIMIT) break;
    if (String(fupData[i][9]) !== 'PENDING') continue;

    const schedDate = new Date(fupData[i][7]);
    if (isNaN(schedDate.getTime()) || schedDate > now) continue;

    const contactId = String(fupData[i][2]);
    const contact = contacts[contactId];

    // Double safety check: Verify contact is STILL IN_SEQUENCE
    if (!contact || String(contact[12]) !== 'IN_SEQUENCE') {
      fupSheet.getRange(i + 1, 10).setValue('CANCELLED_REPLIED');
      continue;
    }

    const threadId = String(fupData[i][3]);
    if (!threadId) continue;

    try {
      const thread = GmailApp.getThreadById(threadId);
      if (!thread) {
        console.warn('Thread not found for ID: ' + threadId);
        continue;
      }

      // Inline reply preserves RFC 2822 In-Reply-To and References headers
      thread.reply('', { htmlBody: String(fupData[i][6]) });
      fupSheet.getRange(i + 1, 9, 1, 2).setValues([[new Date().toISOString(), 'SENT']]);

      sentToday++;
      props.setProperty(todayKey, String(sentToday));
      Utilities.sleep(ANTIGRAVITY_CONFIG.JITTER_MIN_MS);
    } catch (e) {
      console.error('Error dispatching follow-up on thread ' + threadId + ': ' + e.message);
    }
  }
}

/**
 * Adds N business days to a given date, skipping weekends (Saturday & Sunday).
 */
function addBusinessDays_(startDate, days) {
  const d = new Date(startDate.getTime());
  let added = 0;
  while (added < days) {
    d.setDate(d.getDate() + 1);
    if (d.getDay() !== 0 && d.getDay() !== 6) {
      added++;
    }
  }
  return d;
}

// Node.js export compatibility for offline testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    addBusinessDays_,
    scheduleFollowUpRows_
  };
}
