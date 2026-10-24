/**
 * Optra360 & Zoptro Automated Outreach & Opportunity Engine
 * 
 * 06_OutreachDispatcher.js: Branded HTML email creation, Gmail Drafts bundler & LinkedIn queue.
 */

/**
 * Creates Gmail drafts with attached pitch deck PDFs for all contacts in 'Decks Ready' status.
 * @return {Object}
 */
function processOutreachEmailBatch() {
  var contacts = getSheetDataAsObjects(OPTRA_CONFIG.TABS.CONTACTS);
  var problems = getSheetDataAsObjects(OPTRA_CONFIG.TABS.PROBLEM_SOLUTION);
  var demos = getSheetDataAsObjects(OPTRA_CONFIG.TABS.DEMOS);
  var activeProfile = getActiveProfile();

  var probMap = {};
  for (var p = 0; p < problems.length; p++) {
    probMap[problems[p].Company_ID] = problems[p];
  }

  var demoMap = {};
  for (var d = 0; d < demos.length; d++) {
    demoMap[demos[d].Demo_ID] = demos[d];
  }

  var draftsCreated = 0;

  for (var i = 0; i < contacts.length; i++) {
    var c = contacts[i];
    if (c.Outreach_Status !== 'Decks Ready') continue;

    var prob = probMap[c.Company_ID] || {};
    var matchedDemo = demoMap[prob.Matched_Demo_ID] || (demos.length > 0 ? demos[0] : null);

    // 1. Generate 300-char LinkedIn Connection Note
    var linkedInNote = generateLinkedInConnectionNote_(c.Prospect_Name, c.Company_Name, c.Identified_Trigger, activeProfile);

    // 2. Build Branded HTML Email
    var emailSubject = 'Operational Transformation & ' + (c.Identified_Trigger || 'Process') + ' Brief — ' + c.Company_Name;
    var htmlBody = buildBrandedHtmlEmail_(c, prob, matchedDemo, activeProfile);

    // 3. Collect PDF Attachments
    var attachments = [];
    if (c.Deck_5_Slide_ID) {
      try { attachments.push(DriveApp.getFileById(c.Deck_5_Slide_ID).getAs('application/pdf')); } catch (e) {}
    }
    if (c.Deck_12_Slide_ID) {
      try { attachments.push(DriveApp.getFileById(c.Deck_12_Slide_ID).getAs('application/pdf')); } catch (e) {}
    }

    // 4. Create Gmail Draft
    try {
      GmailApp.createDraft(c.Email, emailSubject, '', {
        htmlBody: htmlBody,
        attachments: attachments,
        name: activeProfile.Name + ' | Optra360'
      });

      // Update CRM Status
      updateSheetRow(OPTRA_CONFIG.TABS.CONTACTS, 'Contact_ID', c.Contact_ID, {
        Outreach_Status: 'Drafted',
        LinkedIn_Note: linkedInNote,
        Draft_Date: Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss')
      });

      draftsCreated++;
    } catch (e) {
      console.warn('Failed to create draft for ' + c.Email + ': ' + e.message);
    }

    if (draftsCreated >= 10) break; // Daily limit pacing
  }

  return {
    status: 'success',
    draftsCreated: draftsCreated
  };
}

/**
 * Builds responsive, brand-compliant HTML email.
 * @private
 */
function buildBrandedHtmlEmail_(contact, problem, demo, profile) {
  var c = OPTRA_CONFIG.BRAND.COLORS;
  var f = OPTRA_CONFIG.BRAND.FONTS;

  var firstName = (contact.Prospect_Name || 'there').split(' ')[0];
  var demoSection = '';

  if (demo && demo.Live_URL) {
    demoSection = 
      '<div style="background-color: ' + c.CLARITY_IVORY + '; border-left: 4px solid ' + c.PROGRESS_ORANGE + '; padding: 15px 20px; margin: 25px 0;">' +
      '<p style="margin: 0 0 8px 0; font-family: ' + f.HEADING + '; font-weight: 600; color: ' + c.AUTHORITY_BLUE + '; font-size: 15px;">Interactive Proof / Live Demo:</p>' +
      '<p style="margin: 0 0 10px 0; font-size: 14px; color: ' + c.AUTHORITY_BLUE + ';">' + (demo.Description || 'Operational & HR Tech Architecture at ~2,500 employee scale.') + '</p>' +
      '<a href="' + demo.Live_URL + '" style="display: inline-block; background-color: ' + c.AUTHORITY_BLUE + '; color: ' + c.WHITE + '; text-decoration: none; padding: 10px 18px; font-weight: 600; font-size: 13px; border-radius: 0px;">Explore Live System &rarr;</a>' +
      '</div>';
  }

  var emailSignature = buildOptraSignatureHtml_(profile);

  return (
    '<div style="font-family: ' + f.BODY + '; color: ' + c.AUTHORITY_BLUE + '; line-height: 1.6; max-width: 620px; margin: 0 auto; padding: 20px 0;">' +
      '<p style="font-size: 16px;">Dear ' + firstName + ',</p>' +
      '<p style="font-size: 15px;">I noticed ' + contact.Company_Name + '\'s operational momentum and wanted to share a targeted observation regarding <strong>' + (contact.Identified_Trigger || 'process predictability') + '</strong>.</p>' +
      '<p style="font-size: 15px;">' + (problem.Formulated_Pain_Statement || 'Many manufacturing and industrial operations struggle with recurring workflow friction, departmental disconnects, and manual reconciliation overhead.') + '</p>' +
      '<p style="font-size: 15px;">At <strong>Optra360</strong>, our core operating philosophy is <em>"Prove before you promise."</em> We address these friction points through our <strong>6-Stage Delivery Framework (Diagnose &rarr; Design &rarr; Digitize &rarr; Deploy &rarr; Stabilize &rarr; Govern)</strong>.</p>' +
      demoSection +
      '<p style="font-size: 15px;">I have attached two customized briefs for your team:</p>' +
      '<ul style="font-size: 14px; padding-left: 20px;">' +
        '<li><strong>5-Slide Executive Teaser:</strong> Synthesizing ' + contact.Company_Name + '\'s current trigger and projected 90-day impact.</li>' +
        '<li><strong>12-Slide Transformation &amp; Proof Deck:</strong> In-depth delivery methodology, client case studies, and phased engagement options (Paid Diagnostic &amp; Controlled Pilot).</li>' +
      '</ul>' +
      '<p style="font-size: 15px;">Would you have 15 minutes next week for a brief walkthrough of how we executed this at a similar scale?</p>' +
      '<p style="font-size: 15px; margin-bottom: 25px;">Best regards,</p>' +
      emailSignature +
    '</div>'
  );
}

/**
 * Builds standard Optra360 email signature adhering to brand expression rules.
 * @private
 */
function buildOptraSignatureHtml_(profile) {
  var c = OPTRA_CONFIG.BRAND.COLORS;
  var f = OPTRA_CONFIG.BRAND.FONTS;

  return (
    '<table cellpadding="0" cellspacing="0" border="0" style="font-family: ' + f.BODY + '; border-top: 2px solid ' + c.PROGRESS_ORANGE + '; padding-top: 15px; margin-top: 15px;">' +
      '<tr>' +
        '<td>' +
          '<div style="font-family: ' + f.HEADING + '; font-size: 17px; font-weight: 700; color: ' + c.AUTHORITY_BLUE + ';">' + profile.Name + '</div>' +
          '<div style="font-size: 13px; color: ' + c.GROUNDED_OLIVE + '; font-weight: 600; margin-top: 2px;">' + profile.Title + '</div>' +
          '<div style="font-size: 13px; color: ' + c.AUTHORITY_BLUE + '; margin-top: 6px;">' +
            '<strong>Optra360</strong> | <em>Prove Before You Promise</em>' +
          '</div>' +
          '<div style="font-size: 12px; color: ' + c.MUTED_GRAY + '; margin-top: 6px;">' +
            'Email: <a href="mailto:' + profile.Email + '" style="color: ' + c.AUTHORITY_BLUE + '; text-decoration: none;">' + profile.Email + '</a> | ' +
            'LinkedIn: <a href="' + (profile.LinkedIn_URL || 'https://linkedin.com') + '" style="color: ' + c.PROGRESS_ORANGE + '; text-decoration: none;">Profile</a>' +
          '</div>' +
        '</td>' +
      '</tr>' +
    '</table>'
  );
}

/**
 * Generates concise 300-character LinkedIn connection note.
 * @private
 */
function generateLinkedInConnectionNote_(prospectName, companyName, trigger, profile) {
  var firstName = (prospectName || 'there').split(' ')[0];
  var note = 'Hi ' + firstName + ', noticed ' + companyName + '\'s operational footprint. At Optra360, we specialize in resolving ' + (trigger || 'process friction') + ' via our 6-stage transformation framework and Zoptro HRMS. Would love to connect and share our brief on your sector.';

  if (note.length > 300) {
    note = note.substring(0, 297) + '...';
  }
  return note;
}
