/**
 * Optra360 & Zoptro Automated Outreach & Opportunity Engine
 * 
 * 02_ResearchEngine.js: Daily automated research, news signal tracking & buying trigger detection.
 */

/**
 * Executes the daily research pipeline to discover prospective enterprise companies.
 * @return {Object} Summary of discovered companies.
 */
function runDailyResearch() {
  var props = PropertiesService.getScriptProperties();
  var geminiKey = props.getProperty(OPTRA_CONFIG.PROPERTY_KEYS.GEMINI_KEY);
  var searchKey = props.getProperty(OPTRA_CONFIG.PROPERTY_KEYS.SEARCH_KEY);
  var searchCx = props.getProperty(OPTRA_CONFIG.PROPERTY_KEYS.SEARCH_CX);

  var discovered = [];

  // 1. Try Google Custom Search API if configured, otherwise fallback to Google News RSS
  for (var s = 0; s < OPTRA_CONFIG.SECTORS.length; s++) {
    var sector = OPTRA_CONFIG.SECTORS[s];
    var sectorResults = [];

    if (searchKey && searchCx) {
      sectorResults = searchCompaniesViaCustomSearch_(sector, searchKey, searchCx);
    }

    if (sectorResults.length === 0) {
      sectorResults = searchCompaniesViaNewsRss_(sector);
    }

    discovered = discovered.concat(sectorResults);
    if (discovered.length >= 10) break; // Limit daily new targets
  }

  // 2. Filter out already logged companies from CRM
  var existing = getSheetDataAsObjects(OPTRA_CONFIG.TABS.TARGET_COMPANIES);
  var existingNames = existing.map(function(c) { return String(c.Company_Name || '').toLowerCase().trim(); });

  var newLoggedCount = 0;
  for (var i = 0; i < discovered.length; i++) {
    var comp = discovered[i];
    var normalizedName = String(comp.companyName).toLowerCase().trim();

    if (existingNames.indexOf(normalizedName) === -1 && comp.companyName) {
      var companyId = 'COMP-' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd') + '-' + (existing.length + newLoggedCount + 1);

      // Append to Target Companies Tab
      appendSheetRow(OPTRA_CONFIG.TABS.TARGET_COMPANIES, {
        Company_ID: companyId,
        Company_Name: comp.companyName,
        Industry_Sector: comp.sector,
        Est_Headcount: comp.headcount || '500-2,500',
        Domain_Website: comp.domain,
        Identified_Trigger: comp.trigger,
        Source_Signal: comp.signalTitle,
        Signal_URL: comp.signalUrl,
        Sourcing_Date: Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd'),
        Status: 'New Lead'
      });

      // Also seed an entry in Problem_Solution tab
      appendSheetRow(OPTRA_CONFIG.TABS.PROBLEM_SOLUTION, {
        Company_ID: companyId,
        Company_Name: comp.companyName,
        Identified_Trigger: comp.trigger,
        Raw_Problem_Signal: comp.signalDescription,
        Diagnosis_Status: 'Pending AI Analysis'
      });

      existingNames.push(normalizedName);
      newLoggedCount++;
    }
  }

  return {
    status: 'success',
    newCompaniesLogged: newLoggedCount,
    totalDiscovered: discovered.length
  };
}

/**
 * Zero-Cost Fallback: Parses Google News RSS feeds for target sectors and corporate signals.
 * @private
 */
function searchCompaniesViaNewsRss_(sector) {
  var results = [];
  try {
    var query = encodeURIComponent(sector + ' India expansion plant commissioning management leadership restructuring');
    var feedUrl = 'https://news.google.com/rss/search?q=' + query + '&hl=en-IN&gl=IN&ceid=IN:en';

    var resp = UrlFetchApp.fetch(feedUrl, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return results;

    var xml = resp.getContentText();
    var document = XmlService.parse(xml);
    var root = document.getRootElement();
    var channel = root.getChild('channel');
    if (!channel) return results;

    var items = channel.getChildren('item');
    for (var i = 0; i < Math.min(items.length, 5); i++) {
      var item = items[i];
      var title = item.getChildText('title') || '';
      var link = item.getChildText('link') || '';
      var pubDate = item.getChildText('pubDate') || '';
      var desc = item.getChildText('description') || '';

      // Infer company name & trigger
      var parsed = inferCompanyAndTriggerFromTitle_(title, sector);
      if (parsed.companyName) {
        results.push({
          companyName: parsed.companyName,
          sector: sector,
          domain: parsed.domain || (parsed.companyName.toLowerCase().replace(/[^a-z0-9]/g, '') + '.com'),
          trigger: parsed.trigger,
          signalTitle: title,
          signalUrl: link,
          signalDescription: desc || title,
          headcount: '500-3,000'
        });
      }
    }
  } catch (e) {
    console.warn('News RSS fetch error for sector ' + sector + ': ' + e.message);
  }
  return results;
}

/**
 * Searches via Google Custom Search API.
 * @private
 */
function searchCompaniesViaCustomSearch_(sector, apiKey, cx) {
  var results = [];
  try {
    var query = encodeURIComponent(sector + ' manufacturing companies India plant operations');
    var url = 'https://www.googleapis.com/customsearch/v1?key=' + apiKey + '&cx=' + cx + '&q=' + query;

    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() === 200) {
      var data = JSON.parse(resp.getContentText());
      if (data.items) {
        for (var i = 0; i < Math.min(data.items.length, 5); i++) {
          var it = data.items[i];
          results.push({
            companyName: it.title.split('-')[0].split('|')[0].trim(),
            sector: sector,
            domain: it.displayLink || '',
            trigger: OPTRA_CONFIG.BUYING_TRIGGERS[i % OPTRA_CONFIG.BUYING_TRIGGERS.length],
            signalTitle: it.title,
            signalUrl: it.link,
            signalDescription: it.snippet,
            headcount: '1,000-5,000'
          });
        }
      }
    }
  } catch (e) {
    console.warn('Custom search error: ' + e.message);
  }
  return results;
}

/**
 * Parses article titles to infer company entities and likely buying trigger.
 * @private
 */
function inferCompanyAndTriggerFromTitle_(title, sector) {
  var clean = title.replace(/\s*-\s*[^-]+$/, '').trim(); // Remove publisher suffix (e.g. " - Economic Times")
  var trigger = OPTRA_CONFIG.BUYING_TRIGGERS[0];

  var lower = title.toLowerCase();
  if (lower.indexOf('loss') !== -1 || lower.indexOf('drop') !== -1 || lower.indexOf('margin') !== -1) {
    trigger = 'Performance Stagnation';
  } else if (lower.indexOf('strike') !== -1 || lower.indexOf('attrition') !== -1 || lower.indexOf('union') !== -1) {
    trigger = 'Workforce Misalignment';
  } else if (lower.indexOf('delay') !== -1 || lower.indexOf('halt') !== -1 || lower.indexOf('shutdown') !== -1) {
    trigger = 'Process Breakdown';
  } else if (lower.indexOf('ceo') !== -1 || lower.indexOf('resign') !== -1 || lower.indexOf('appoint') !== -1) {
    trigger = 'Leadership Disconnect';
  } else if (lower.indexOf('erp') !== -1 || lower.indexOf('software') !== -1 || lower.indexOf('legacy') !== -1) {
    trigger = 'System Fragmentation';
  }

  // Extract first 1-3 capitalized words as potential company name
  var words = clean.split(' ');
  var compCandidates = [];
  for (var w = 0; w < Math.min(words.length, 3); w++) {
    if (words[w] && words[w][0] === words[w][0].toUpperCase() && words[w].length > 2) {
      compCandidates.push(words[w]);
    }
  }
  var compName = compCandidates.join(' ');
  if (!compName) compName = clean.substring(0, 30);

  return {
    companyName: compName,
    domain: compName.toLowerCase().replace(/[^a-z0-9]/g, '') + '.com',
    trigger: trigger
  };
}
