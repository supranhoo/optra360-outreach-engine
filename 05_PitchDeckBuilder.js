/**
 * Optra360 & Zoptro Automated Outreach & Opportunity Engine
 * 
 * 05_PitchDeckBuilder.js: Two-tier dynamic Google Slides cloning, tag merging & PDF export.
 */

/**
 * Builds and exports both 5-Slide Teaser and 12-Slide Detailed pitch decks for pending contacts.
 * @return {Object}
 */
function buildAllPendingPitchDecks() {
  var contacts = getSheetDataAsObjects(OPTRA_CONFIG.TABS.CONTACTS);
  var problems = getSheetDataAsObjects(OPTRA_CONFIG.TABS.PROBLEM_SOLUTION);
  var activeProfile = getActiveProfile();

  var probMap = {};
  for (var p = 0; p < problems.length; p++) {
    probMap[problems[p].Company_ID] = problems[p];
  }

  var builtCount = 0;

  for (var i = 0; i < contacts.length; i++) {
    var c = contacts[i];
    if (c.Outreach_Status !== 'Pending Pitch Deck' && c.Outreach_Status !== 'New Lead') continue;

    var prob = probMap[c.Company_ID] || {};
    var deckPair = generateDeckPairForContact_(c, prob, activeProfile);

    if (deckPair.teaserPdfUrl && deckPair.detailedPdfUrl) {
      updateSheetRow(OPTRA_CONFIG.TABS.CONTACTS, 'Contact_ID', c.Contact_ID, {
        Deck_5_Slide_PDF: deckPair.teaserPdfUrl,
        Deck_12_Slide_PDF: deckPair.detailedPdfUrl,
        Deck_5_Slide_ID: deckPair.teaserPdfId,
        Deck_12_Slide_ID: deckPair.detailedPdfId,
        Outreach_Status: 'Decks Ready'
      });
      builtCount++;
    }

    if (builtCount >= 5) break; // Batch pacing
  }

  return {
    status: 'success',
    deckPairsGenerated: builtCount
  };
}

/**
 * Clones, personalizes, and exports a 5-Slide and 12-Slide deck pair.
 * @private
 */
function generateDeckPairForContact_(contact, problem, profile) {
  var folder = getOrCreateOutreachFolder_();

  var mergeTags = {
    '{{Company_Name}}': contact.Company_Name || 'Enterprise Partner',
    '{{Prospect_Name}}': contact.Prospect_Name || 'Executive Leader',
    '{{Designation}}': contact.Designation || 'Department Head',
    '{{Identified_Trigger}}': contact.Identified_Trigger || 'Process Breakdown',
    '{{Specific_Pain}}': problem.Formulated_Pain_Statement || 'Operational friction across core business workflows.',
    '{{Recommended_Solution}}': problem.Recommended_Solution || 'Optra360 6-Stage Delivery Framework.',
    '{{Expected_Impact}}': problem.Expected_Transformation_Impact || '20-30% efficiency gain and automated control.',
    '{{Engagement_Model}}': problem.Recommended_Engagement || '₹3–7.5L Paid Diagnostic Audit',
    '{{Sender_Name}}': profile.Name || 'Ankit Choudhary',
    '{{Sender_Title}}': profile.Title || 'Co-founder & CTO',
    '{{Sender_Email}}': profile.Email || 'ankit.choudhary@optra360.com',
    '{{Sender_Phone}}': profile.Phone || '+91 99999 99999',
    '{{Sender_LinkedIn}}': profile.LinkedIn_URL || 'https://linkedin.com/in/ankitchoudhary',
    '{{Date}}': Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMMM yyyy')
  };

  var teaserName = 'Optra360_Executive_Teaser_' + sanitizeFileName_(contact.Company_Name);
  var detailedName = 'Optra360_Detailed_Transformation_' + sanitizeFileName_(contact.Company_Name);

  var teaserPdf = cloneAndMergeSlidesDeck_(OPTRA_CONFIG.SLIDES_TEASER_ID, teaserName, mergeTags, folder);
  var detailedPdf = cloneAndMergeSlidesDeck_(OPTRA_CONFIG.SLIDES_DETAILED_ID, detailedName, mergeTags, folder);

  return {
    teaserPdfUrl: teaserPdf.url,
    teaserPdfId: teaserPdf.id,
    detailedPdfUrl: detailedPdf.url,
    detailedPdfId: detailedPdf.id
  };
}

/**
 * Clones master Google Slides presentation, substitutes merge tags, and converts to PDF.
 * @private
 */
function cloneAndMergeSlidesDeck_(templateId, newTitle, mergeTags, targetFolder) {
  try {
    var templateFile = DriveApp.getFileById(templateId);
    var clonedFile = templateFile.makeCopy(newTitle, targetFolder);
    var presentation = SlidesApp.openById(clonedFile.getId());

    // Merge tags across all slides
    for (var tag in mergeTags) {
      presentation.replaceAllText(tag, mergeTags[tag]);
    }
    presentation.saveAndClose();

    // Export as PDF
    var pdfBlob = clonedFile.getAs('application/pdf');
    pdfBlob.setName(newTitle + '.pdf');
    var pdfFile = targetFolder.createFile(pdfBlob);

    // Set permission to viewable with link for outreach
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Trash temporary Google Slides file
    clonedFile.setTrashed(true);

    return {
      id: pdfFile.getId(),
      url: pdfFile.getUrl(),
      blob: pdfBlob
    };
  } catch (e) {
    console.warn('Could not clone template ' + templateId + ': ' + e.message + '. Generating branded fallback PDF.');
    return generateBrandedFallbackPdf_(newTitle, mergeTags, targetFolder);
  }
}

/**
 * Fallback branded PDF generator if master slides template is restricted.
 * @private
 */
function generateBrandedFallbackPdf_(title, tags, targetFolder) {
  var html = 
    '<html><body style="font-family: \'Figtree\', sans-serif; background-color: #F4F2EA; color: #09293D; padding: 40px;">' +
    '<div style="background-color: #09293D; color: #F4F2EA; padding: 25px; border-radius: 0px;">' +
    '<h1 style="font-family: \'Outfit\', sans-serif; margin: 0; font-size: 28px; text-transform: uppercase; letter-spacing: 1px;">Optra360 &amp; Zoptro</h1>' +
    '<p style="margin: 5px 0 0 0; color: #F65800; font-weight: bold;">' + tags['{{Identified_Trigger}}'] + ' Transformation Brief</p>' +
    '</div>' +
    '<div style="background: #FFFFFF; padding: 30px; margin-top: 20px; border-left: 5px solid #F65800;">' +
    '<h2 style="font-family: \'Outfit\', sans-serif; color: #09293D;">Executive Problem Diagnosis for ' + tags['{{Company_Name}}'] + '</h2>' +
    '<p style="font-size: 16px; line-height: 1.6;">' + tags['{{Specific_Pain}}'] + '</p>' +
    '<h3 style="font-family: \'Outfit\', sans-serif; color: #666D57; margin-top: 25px;">Recommended Transformation Architecture</h3>' +
    '<p style="font-size: 16px; line-height: 1.6;">' + tags['{{Recommended_Solution}}'] + '</p>' +
    '<div style="background-color: #F4F2EA; padding: 15px; margin-top: 20px;">' +
    '<strong>Expected Business Impact:</strong> ' + tags['{{Expected_Impact}}'] + '<br/>' +
    '<strong>Commercial Model:</strong> ' + tags['{{Engagement_Model}}'] +
    '</div>' +
    '</div>' +
    '<div style="margin-top: 30px; font-size: 13px; color: #666D57;">' +
    'Prepared by ' + tags['{{Sender_Name}}'] + ' (' + tags['{{Sender_Title}}'] + ') | ' + tags['{{Sender_Email}}'] + '<br/>' +
    'Optra360 — Prove Before You Promise.' +
    '</div>' +
    '</body></html>';

  var blob = Utilities.newBlob(html, 'text/html', title + '.html').getAs('application/pdf');
  blob.setName(title + '.pdf');
  var file = targetFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    id: file.getId(),
    url: file.getUrl(),
    blob: blob
  };
}

/**
 * Gets or creates the Google Drive folder for generated decks.
 * @private
 */
function getOrCreateOutreachFolder_() {
  var folderName = 'Optra360 Generated Pitch Decks';
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

/**
 * Sanitizes strings for filenames.
 * @private
 */
function sanitizeFileName_(name) {
  return String(name || 'Lead').replace(/[^a-zA-Z0-9_-]/g, '_');
}
