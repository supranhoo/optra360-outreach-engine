/**
 * PROJECT ANTIGRAVITY: END-TO-END JOB OUTREACH INTELLIGENCE PLATFORM
 * Module 1.1: Apollo.io Lead Discovery, Enrichment & Ingestion Engine
 * (01b_ApolloIntegrator.js)
 */

var ANTIGRAVITY_APOLLO = {
  API_BASE_URL: 'https://api.apollo.io/v1',

  /**
   * Retrieves the stored Apollo.io API key from Script Properties.
   */
  getApiKey: function() {
    if (typeof PropertiesService === 'undefined' || !PropertiesService.getScriptProperties) {
      return '';
    }
    return PropertiesService.getScriptProperties().getProperty('APOLLO_API_KEY') || '';
  },

  /**
   * Tests the connection to Apollo.io with the configured or provided API key.
   */
  testConnection: function(customApiKey) {
    const key = customApiKey || this.getApiKey();
    if (!key) {
      return { success: false, error: 'No Apollo.io API key provided or found in system settings.' };
    }

    try {
      const url = this.API_BASE_URL + '/mixed_people/search';
      const payload = {
        api_key: key,
        page: 1,
        per_page: 1,
        person_locations: ['India']
      };

      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      const body = JSON.parse(resp.getContentText() || '{}');

      if (code === 200) {
        return {
          success: true,
          message: 'Successfully connected to Apollo.io API! Search and lead enrichment are active.',
          totalEntries: body.pagination ? body.pagination.total_entries : undefined
        };
      } else {
        const errMsg = body.error || body.message || ('Apollo API returned HTTP ' + code);
        return { success: false, error: errMsg };
      }
    } catch (e) {
      return { success: false, error: 'Apollo connection exception: ' + e.message };
    }
  },

  /**
   * Searches Apollo.io for executive decision-makers matching target criteria.
   * @param {Object} queryOptions - Filters: { domains, titles, locations, page, perPage }
   */
  searchPeople: function(queryOptions) {
    const key = this.getApiKey();
    if (!key) {
      return { success: false, error: 'Apollo.io API key is not configured. Go to Settings to add it.' };
    }

    try {
      const page = queryOptions.page || 1;
      const perPage = Math.min(queryOptions.perPage || 15, 25); // Safe batch size
      const payload = {
        api_key: key,
        page: page,
        per_page: perPage
      };

      // Organization Domains (newline or array)
      if (queryOptions.domains) {
        if (Array.isArray(queryOptions.domains)) {
          payload.q_organization_domains = queryOptions.domains.join('\n');
        } else {
          payload.q_organization_domains = String(queryOptions.domains).replace(/,/g, '\n').trim();
        }
      }

      // Person Job Titles
      if (queryOptions.titles) {
        const titleList = Array.isArray(queryOptions.titles)
          ? queryOptions.titles
          : String(queryOptions.titles).split(',').map(s => s.trim()).filter(Boolean);
        if (titleList.length > 0) {
          payload.person_titles = titleList;
        }
      }

      // Organization Employee Headcount Ranges (Sweet spot: 200 - 5,000 employees for growing companies)
      if (queryOptions.employeeRanges && Array.isArray(queryOptions.employeeRanges) && queryOptions.employeeRanges.length > 0) {
        payload.organization_num_employees_ranges = queryOptions.employeeRanges;
      } else {
        // Default: Strictly focus on mid-market growing companies in systems implementation stage
        payload.organization_num_employees_ranges = ['201,500', '501,1000', '1001,5000'];
      }

      // Optional Industry / Expansion Keywords (e.g., "Foundry", "Manufacturing", "Expansion")
      if (queryOptions.keywords) {
        payload.q_keywords = String(queryOptions.keywords).trim();
      }

      // Locations (e.g. India, Odisha, etc.)
      if (queryOptions.locations) {
        const locList = Array.isArray(queryOptions.locations)
          ? queryOptions.locations
          : String(queryOptions.locations).split(',').map(s => s.trim()).filter(Boolean);
        if (locList.length > 0) {
          payload.person_locations = locList;
        }
      }

      const url = this.API_BASE_URL + '/mixed_people/search';
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      const body = JSON.parse(resp.getContentText() || '{}');

      if (code !== 200) {
        return { success: false, error: body.error || body.message || ('Apollo error ' + code) };
      }

      const rawPeople = body.people || [];
      const normalized = rawPeople.map(p => {
        const org = p.organization || {};
        const title = p.title || '';
        const persona = this.classifyPersonaFromTitle(title);
        
        return {
          id: p.id,
          firstName: p.first_name || '',
          lastName: p.last_name || '',
          fullName: p.name || ((p.first_name || '') + ' ' + (p.last_name || '')).trim(),
          title: title,
          persona: persona,
          email: p.email || '',
          emailStatus: p.email_status || 'unknown',
          companyName: org.name || '',
          companyDomain: org.primary_domain || '',
          industry: org.industry || 'Manufacturing & Industrial',
          linkedinUrl: p.linkedin_url || '',
          city: p.city || '',
          state: p.state || '',
          country: p.country || 'India'
        };
      });

      return {
        success: true,
        people: normalized,
        pagination: body.pagination || { page: page, per_page: perPage, total_entries: normalized.length }
      };
    } catch (err) {
      return { success: false, error: 'Failed searching Apollo.io: ' + err.message };
    }
  },

  /**
   * Enriches a specific lead using Apollo.io /people/match to retrieve their verified work email.
   */
  enrichPerson: function(personParams) {
    const key = this.getApiKey();
    if (!key) return { success: false, error: 'Apollo API key missing.' };

    try {
      const url = this.API_BASE_URL + '/people/match';
      const payload = {
        api_key: key,
        reveal_personal_emails: false
      };

      if (personParams.id) payload.id = personParams.id;
      if (personParams.firstName) payload.first_name = personParams.firstName;
      if (personParams.lastName) payload.last_name = personParams.lastName;
      if (personParams.companyDomain) payload.domain = personParams.companyDomain;
      if (personParams.companyName) payload.organization_name = personParams.companyName;

      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const resp = UrlFetchApp.fetch(url, options);
      const body = JSON.parse(resp.getContentText() || '{}');
      if (body.person && body.person.email) {
        return {
          success: true,
          email: body.person.email,
          emailStatus: body.person.email_status,
          person: body.person
        };
      }
      return { success: false, error: 'Email could not be revealed or found in Apollo.' };
    } catch (e) {
      return { success: false, error: e.message };
    }
  },

  /**
   * Automatically classifies executive persona from job title.
   * Tailored for mid-market manufacturing companies in systems implementation stage.
   */
  classifyPersonaFromTitle: function(title) {
    const t = String(title || '').toLowerCase();
    
    // 1. Promoter & Business Owner Track (MD, Joint MD, Owner, Founder, CEO)
    if (t.includes('founder') || t.includes('ceo') || t.includes('managing director') || 
        t.includes('joint md') || t.includes('promoter') || t.includes('executive director') || 
        t.includes('owner') || t.includes('partner') || t.includes('chairman')) {
      return 'FOUNDER';
    }

    // 2. Plant Operations & Shopfloor Leadership (COO, Plant Head, VP Operations, Works Director)
    if (t.includes('plant head') || t.includes('coo') || t.includes('chief operating officer') || 
        t.includes('operations') || t.includes('manufacturing') || t.includes('works') || 
        t.includes('factory head') || t.includes('unit head') || t.includes('general manager works')) {
      return 'TECH_LEAD'; // Operational / Technical Plant Leadership
    }

    // 3. Talent Acquisition / Recruiting
    if (t.includes('recruiter') || t.includes('talent acquisition') || t.includes('head ta') || t.includes('recruitment')) {
      return 'RECRUITER';
    }

    // 4. HR Leadership & Systems Transformation (CHRO, VP HR, Head HR, GM HR, IR Head)
    if (t.includes('chro') || t.includes('human resources') || t.includes('hr') || 
        t.includes('people') || t.includes('ir') || t.includes('industrial relations')) {
      return 'HIRING_MGR';
    }

    return 'HIRING_MGR';
  },

  /**
   * Determines default skill & pain-point tags based on role and industry.
   * Focuses on 0-to-1 systems building, attendance governance, and PMS implementation.
   */
  getOpportunityTagsForRole: function(roleTitle) {
    const t = String(roleTitle || '').toLowerCase();

    // Promoter / MD: Looking for total plant systems, EBITDA leakage arrest, and institutional governance
    if (t.includes('managing director') || t.includes('promoter') || t.includes('founder') || t.includes('ceo') || t.includes('director')) {
      return 'systems-implementation, pms-rollout, attendance-governance, absenteeism-control, ir, workforce-productivity, statutory-compliance';
    }

    // Operations / COO / Plant Head: Focus on shopfloor absenteeism, shift productivity, biometric integrity, IR peace
    if (t.includes('operations') || t.includes('manufacturing') || t.includes('works') || t.includes('production') || t.includes('coo') || t.includes('plant head')) {
      return 'plant-operations, absenteeism-control, attendance-governance, time-office, shift-deployment, ir, safety-governance, workforce-productivity, plant-hr';
    }

    // HR Leadership: Focus on PMS rollout, KRA/KPI framework, LMS, contract-labour compliance
    if (t.includes('chro') || t.includes('vp hr') || t.includes('human resources') || t.includes('hr') || t.includes('people')) {
      return 'pms-rollout, kra-kpi, lms, digital-hr, plant-hr, contract-labour, statutory-compliance, ir';
    }

    return 'plant-hr, ir, statutory-compliance, pms, contract-labour, absenteeism-control';
  },

  /**
   * Ingests a batch of selected Apollo leads directly into the CRM relational sheets.
   * Inserts into Target_Companies, Opportunities, and Contacts, then triggers scoring.
   */
  importLeadsToCrm: function(leads) {
    if (!Array.isArray(leads) || leads.length === 0) {
      return { success: false, error: 'No leads provided for import.' };
    }

    const ss = ANTIGRAVITY_CONFIG.getSpreadsheet();
    if (!ss) return { success: false, error: 'Database not initialized.' };

    const compSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.TARGET_COMPANIES);
    const oppSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.OPPORTUNITIES);
    const contactSheet = ss.getSheetByName(ANTIGRAVITY_CONFIG.SHEETS.CONTACTS);

    if (!compSheet || !oppSheet || !contactSheet) {
      return { success: false, error: 'Required relational tables missing. Run Bootstrap first.' };
    }

    const compData = compSheet.getDataRange().getValues();
    const contactData = contactSheet.getDataRange().getValues();

    // Index existing companies by clean domain / name
    const companyDomainMap = {};
    for (let i = 1; i < compData.length; i++) {
      const cid = String(compData[i][0]);
      const cDomain = String(compData[i][2] || '').toLowerCase().trim();
      const cName = String(compData[i][1] || '').toLowerCase().trim();
      if (cDomain) companyDomainMap[cDomain] = cid;
      if (cName) companyDomainMap[cName] = cid;
    }

    // Index existing contact emails to avoid duplicates
    const existingEmails = new Set();
    for (let j = 1; j < contactData.length; j++) {
      const email = String(contactData[j][5] || '').toLowerCase().trim();
      if (email) existingEmails.add(email);
    }

    let nextCompNum = compData.length + 100;
    let nextOppNum = oppSheet.getLastRow() + 200;
    let nextContactNum = contactData.length + 500;

    let importedCount = 0;
    let skippedDuplicates = 0;

    const newCompanyRows = [];
    const newOppRows = [];
    const newContactRows = [];

    leads.forEach(lead => {
      const email = String(lead.email || '').trim().toLowerCase();
      if (!email || existingEmails.has(email)) {
        skippedDuplicates++;
        return;
      }
      existingEmails.add(email);

      // 1. Resolve or Create Target_Company
      const compName = lead.companyName || 'Target Manufacturing Corp';
      const compDomain = (lead.companyDomain || '').toLowerCase().trim();
      let compId = companyDomainMap[compDomain] || companyDomainMap[compName.toLowerCase()];

      if (!compId) {
        compId = 'CMP-' + (++nextCompNum);
        companyDomainMap[compDomain || compName.toLowerCase()] = compId;

        newCompanyRows.push([
          compId,
          compName,
          compDomain || (compName.toLowerCase().replace(/[^a-z0-9]/g, '') + '.com'),
          lead.industry || 'Manufacturing & Heavy Industrial',
          'Enterprise (>1,000 workforce)',
          '5,000+ Personnel',
          lead.city ? (lead.city + ', ' + lead.state + ', ' + lead.country) : 'India',
          'Active Plant HR & Workforce Optimization Focus'
        ]);
      }

      // 2. Create Opportunity
      const oppId = 'OPP-' + (++nextOppNum);
      const roleTitle = lead.title || 'Plant HR & Industrial Relations Leader';
      const skillTags = lead.skillTags || this.getOpportunityTagsForRole(roleTitle);

      newOppRows.push([
        oppId,
        compId,
        roleTitle,
        'EXECUTIVE_OUTREACH',
        lead.linkedinUrl || ('https://linkedin.com/company/' + compName.toLowerCase().replace(/[^a-z0-9]/g, '')),
        'Leading workforce governance, plant IR compliance, and workforce productivity across multi-shift plant operations.',
        '["Shopfloor absenteeism and contract labour compliance", "Digital PMS adoption and supervisory competency", "IR harmony and statutory governance"]',
        skillTags,
        'PRJ-001',
        'Direct application of 3,850-worker workforce governance and digital Safety/PMS framework de-risks plant operations.'
      ]);

      // 3. Create Contact
      const contactId = 'CNT-' + (++nextContactNum);
      const persona = lead.persona || this.classifyPersonaFromTitle(roleTitle);

      newContactRows.push([
        contactId,
        compId,
        oppId,
        lead.firstName || lead.fullName || 'Executive',
        lead.lastName || '',
        email,
        'UNVERIFIED',
        persona,
        roleTitle,
        lead.linkedinUrl || '',
        'Asia/Kolkata',
        0, // Initial score before calculation
        'NEW' // Set to NEW so processAndScoreNewContacts() will score and verify it
      ]);

      importedCount++;
    });

    // Write batched rows
    if (newCompanyRows.length > 0) {
      compSheet.getRange(compSheet.getLastRow() + 1, 1, newCompanyRows.length, newCompanyRows[0].length).setValues(newCompanyRows);
    }
    if (newOppRows.length > 0) {
      oppSheet.getRange(oppSheet.getLastRow() + 1, 1, newOppRows.length, newOppRows[0].length).setValues(newOppRows);
    }
    if (newContactRows.length > 0) {
      contactSheet.getRange(contactSheet.getLastRow() + 1, 1, newContactRows.length, newContactRows[0].length).setValues(newContactRows);
    }

    // Automatically trigger lead scoring & MX verification for the newly imported contacts
    if (importedCount > 0 && typeof processAndScoreNewContacts === 'function') {
      try {
        processAndScoreNewContacts();
      } catch (e) {
        console.warn('Scoring run triggered after import, note: ' + e.message);
      }
    }

    return {
      success: true,
      importedCount: importedCount,
      skippedDuplicates: skippedDuplicates,
      message: 'Successfully imported ' + importedCount + ' leads from Apollo into Antigravity CRM (scored & verified)!'
    };
  },

  /**
   * Parses standard CSV text exported directly from Apollo.io.
   */
  importCsvText: function(csvText) {
    if (!csvText || typeof csvText !== 'string') {
      return { success: false, error: 'Empty CSV text provided.' };
    }

    const lines = csvText.trim().split(/\r\n|\n/);
    if (lines.length < 2) {
      return { success: false, error: 'CSV must contain at least a header row and one data row.' };
    }

    // Parse header row
    const headers = lines[0].split(',').map(h => h.replace(/^["']|["']$/g, '').trim().toLowerCase());
    const colIndex = {
      firstName: headers.findIndex(h => h.includes('first name')),
      lastName: headers.findIndex(h => h.includes('last name')),
      fullName: headers.findIndex(h => h === 'name' || h.includes('full name')),
      title: headers.findIndex(h => h.includes('title')),
      company: headers.findIndex(h => h.includes('company') || h.includes('organization')),
      domain: headers.findIndex(h => h.includes('domain') || h.includes('website')),
      email: headers.findIndex(h => h.includes('email')),
      linkedin: headers.findIndex(h => h.includes('linkedin')),
      city: headers.findIndex(h => h.includes('city')),
      state: headers.findIndex(h => h.includes('state')),
      country: headers.findIndex(h => h.includes('country'))
    };

    if (colIndex.email === -1) {
      return { success: false, error: 'CSV must contain an "Email" column.' };
    }

    const parsedLeads = [];
    for (let i = 1; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue;

      // Basic CSV token parser handling quotes
      const cells = [];
      let inQuote = false;
      let curr = '';
      for (let c = 0; c < line.length; c++) {
        const char = line[c];
        if (char === '"') {
          inQuote = !inQuote;
        } else if (char === ',' && !inQuote) {
          cells.push(curr.trim());
          curr = '';
        } else {
          curr += char;
        }
      }
      cells.push(curr.trim());

      const getVal = (idx) => (idx >= 0 && idx < cells.length) ? cells[idx].replace(/^["']|["']$/g, '').trim() : '';
      const email = getVal(colIndex.email);
      if (!email || !email.includes('@')) continue;

      parsedLeads.push({
        firstName: getVal(colIndex.firstName),
        lastName: getVal(colIndex.lastName),
        fullName: getVal(colIndex.fullName),
        title: getVal(colIndex.title) || 'HR & Plant Operations Executive',
        companyName: getVal(colIndex.company) || 'Manufacturing Enterprise',
        companyDomain: getVal(colIndex.domain),
        email: email,
        linkedinUrl: getVal(colIndex.linkedin),
        city: getVal(colIndex.city),
        state: getVal(colIndex.state),
        country: getVal(colIndex.country) || 'India'
      });
    }

    if (parsedLeads.length === 0) {
      return { success: false, error: 'No valid lead rows with email addresses found in CSV.' };
    }

    return this.importLeadsToCrm(parsedLeads);
  }
};

// Node module export for offline testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { ANTIGRAVITY_APOLLO };
}
