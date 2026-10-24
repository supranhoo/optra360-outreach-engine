/**
 * Optra360 & Zoptro Automated Outreach & Opportunity Engine
 * 
 * 03_ContactEnrichment.js: Decision-maker discovery, permutation generator & DNS MX verification.
 */

/**
 * Enriches contacts for all companies currently in 'New Lead' status.
 * @return {Object}
 */
function enrichProspectContacts() {
  var companies = getSheetDataAsObjects(OPTRA_CONFIG.TABS.TARGET_COMPANIES);
  var contacts = getSheetDataAsObjects(OPTRA_CONFIG.TABS.CONTACTS);
  var existingContactEmails = contacts.map(function(c) { return String(c.Email || '').toLowerCase().trim(); });

  var props = PropertiesService.getScriptProperties();
  var hunterKey = props.getProperty(OPTRA_CONFIG.PROPERTY_KEYS.HUNTER_KEY);
  var apolloKey = props.getProperty(OPTRA_CONFIG.PROPERTY_KEYS.APOLLO_KEY);

  var enrichedCount = 0;

  for (var i = 0; i < companies.length; i++) {
    var comp = companies[i];
    if (comp.Status !== 'New Lead' && comp.Status !== 'Active') continue;

    var domain = cleanDomain_(comp.Domain_Website || '');
    if (!domain) continue;

    // Verify MX records for the domain first
    var hasMx = verifyDomainMxRecord_(domain);
    if (!hasMx) {
      console.warn('Domain has no active MX records: ' + domain);
      continue;
    }

    var decisionMakers = [];

    // Attempt Hunter.io lookup if key available
    if (hunterKey) {
      decisionMakers = queryHunterForDecisionMakers_(domain, hunterKey);
    }

    // Fallback: Synthesize standard executive roles and generate permutation candidates
    if (decisionMakers.length === 0) {
      decisionMakers = synthesizeExecutivePersonas_(comp.Company_Name, domain);
    }

    // Add each decision maker to Contacts Outreach tab
    for (var d = 0; d < decisionMakers.length; d++) {
      var dm = decisionMakers[d];
      var emailLower = String(dm.email).toLowerCase().trim();

      if (existingContactEmails.indexOf(emailLower) === -1 && dm.email) {
        var contactId = 'CONT-' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd') + '-' + (contacts.length + enrichedCount + 1);

        appendSheetRow(OPTRA_CONFIG.TABS.CONTACTS, {
          Contact_ID: contactId,
          Company_ID: comp.Company_ID,
          Company_Name: comp.Company_Name,
          Prospect_Name: dm.name,
          Designation: dm.title,
          Email: dm.email,
          Email_Status: dm.verificationStatus || 'MX Verified',
          LinkedIn_URL: dm.linkedInUrl || ('https://www.linkedin.com/search/results/all/?keywords=' + encodeURIComponent(dm.name + ' ' + comp.Company_Name)),
          Identified_Trigger: comp.Identified_Trigger || 'Process Breakdown',
          Outreach_Status: 'Pending Pitch Deck',
          Deck_5_Slide_PDF: '',
          Deck_12_Slide_PDF: '',
          LinkedIn_Note: ''
        });

        existingContactEmails.push(emailLower);
        enrichedCount++;
      }
    }

    // Mark company as enriched
    updateSheetRow(OPTRA_CONFIG.TABS.TARGET_COMPANIES, 'Company_ID', comp.Company_ID, {
      Status: 'Enriched'
    });

    if (enrichedCount >= 15) break; // Batch pacing
  }

  return {
    status: 'success',
    newContactsEnriched: enrichedCount
  };
}

/**
 * Checks DNS MX records via Google Public DNS HTTPS API.
 * @param {string} domain 
 * @return {boolean}
 */
function verifyDomainMxRecord_(domain) {
  try {
    var url = 'https://dns.google/resolve?name=' + encodeURIComponent(domain) + '&type=MX';
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    if (resp.getResponseCode() === 200) {
      var json = JSON.parse(resp.getContentText());
      if (json.Answer && json.Answer.length > 0) {
        return true;
      }
    }
  } catch (e) {
    console.warn('DNS MX check failed for ' + domain + ': ' + e.message);
  }
  return false;
}

/**
 * Generates standard corporate email permutations.
 * @param {string} firstName 
 * @param {string} lastName 
 * @param {string} domain 
 * @return {Array<string>}
 */
function generateEmailPermutations(firstName, lastName, domain) {
  var f = firstName.toLowerCase().replace(/[^a-z]/g, '');
  var l = lastName.toLowerCase().replace(/[^a-z]/g, '');
  var dom = cleanDomain_(domain);

  if (!f || !dom) return [];

  var permutations = [];
  if (l) {
    permutations.push(f + '.' + l + '@' + dom);
    permutations.push(f[0] + '.' + l + '@' + dom);
    permutations.push(f + l + '@' + dom);
    permutations.push(f[0] + l + '@' + dom);
    permutations.push(l + '.' + f + '@' + dom);
    permutations.push(f + '@' + dom);
  } else {
    permutations.push(f + '@' + dom);
  }

  return permutations;
}

/**
 * Cleans a domain string (strips protocol, path, and www).
 * @private
 */
function cleanDomain_(domain) {
  return String(domain)
    .replace(/^https?:\/\//i, '')
    .replace(/^www\./i, '')
    .split('/')[0]
    .split(':')[0]
    .trim()
    .toLowerCase();
}

/**
 * Synthesizes top executive personas for cold outreach when no API key is available.
 * @private
 */
function synthesizeExecutivePersonas_(companyName, domain) {
  var cleanName = companyName.split(' ')[0];
  return [
    {
      name: 'Operations Head',
      title: 'VP / Head of Operations',
      email: 'operations@' + domain,
      verificationStatus: 'MX Verified',
      linkedInUrl: 'https://www.linkedin.com/search/results/all/?keywords=' + encodeURIComponent('VP Operations ' + companyName)
    },
    {
      name: 'Human Resources Head',
      title: 'CHRO / Head of HR',
      email: 'hr@' + domain,
      verificationStatus: 'MX Verified',
      linkedInUrl: 'https://www.linkedin.com/search/results/all/?keywords=' + encodeURIComponent('CHRO ' + companyName)
    }
  ];
}

/**
 * Queries Hunter.io for email addresses.
 * @private
 */
function queryHunterForDecisionMakers_(domain, apiKey) {
  var results = [];
  try {
    var url = 'https://api.hunter.io/v2/domain-search?domain=' + encodeURIComponent(domain) + '&api_key=' + apiKey + '&seniority=executive,senior';
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });

    if (resp.getResponseCode() === 200) {
      var data = JSON.parse(resp.getContentText());
      if (data.data && data.data.emails) {
        for (var i = 0; i < Math.min(data.data.emails.length, 3); i++) {
          var em = data.data.emails[i];
          results.push({
            name: (em.first_name + ' ' + em.last_name).trim() || em.position || 'Executive Leader',
            title: em.position || 'Head of Department',
            email: em.value,
            verificationStatus: em.confidence > 70 ? 'Valid' : 'MX Verified',
            linkedInUrl: em.linkedin || ''
          });
        }
      }
    }
  } catch (e) {
    console.warn('Hunter API error: ' + e.message);
  }
  return results;
}
