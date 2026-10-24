/**
 * Optra360 & Zoptro Automated Outreach & Opportunity Engine
 * 
 * 07_DemoPortfolioHub.js: Configurable demo repository & live portfolio showcase.
 */

/**
 * Returns all active demo items from 6_Demos_&_Portfolio tab or default fallback repository.
 * @return {Array<Object>}
 */
function getDemoPortfolio() {
  var demos = getSheetDataAsObjects(OPTRA_CONFIG.TABS.DEMOS);
  if (demos.length > 0) return demos;

  return [
    {
      Demo_ID: 'DEMO-01',
      Title: 'Zoptro HRMS Enterprise Cloud Architecture',
      Category: 'HR Tech & Operations',
      Operating_Scale: '~2,500 Active Employees',
      Live_URL: 'https://github.com/supranhoo/Recruitment-CRM',
      Description: 'Complete single-tenant cloud HRMS managing onboarding, lifecycle tracking, and biometric attendance.',
      Trigger_Alignment: 'System Fragmentation'
    },
    {
      Demo_ID: 'DEMO-02',
      Title: 'Zoptro PMS & Point Factor Evaluation Engine',
      Category: 'Performance & Compensation',
      Operating_Scale: 'Enterprise-wide',
      Live_URL: 'https://docs.google.com/spreadsheets/d/1KEJfOBK8PobuGQ6TI-91AKebXehOphDD3LwQw_zaXwU/edit',
      Description: 'Scientific job evaluation matrix aligning employee contribution scores directly to salary bands.',
      Trigger_Alignment: 'Performance Stagnation'
    },
    {
      Demo_ID: 'DEMO-03',
      Title: 'Biometric-SQL Attendance & Payroll Engine',
      Category: 'Workforce & Statutory Compliance',
      Operating_Scale: 'Multi-Plant Heavy Industry',
      Live_URL: 'https://github.com/supranhoo/Recruitment-CRM',
      Description: 'Automated biometric sync feeding directly into SQL payroll tables with zero variance validation.',
      Trigger_Alignment: 'Workforce Misalignment'
    },
    {
      Demo_ID: 'DEMO-04',
      Title: 'Retail & Plant Maintenance MIS',
      Category: 'Operational Engineering',
      Operating_Scale: 'Multi-Location Commercial',
      Live_URL: 'https://github.com/supranhoo',
      Description: 'Preventive and breakdown maintenance tracking with automated SLA escalation workflows.',
      Trigger_Alignment: 'Process Breakdown'
    },
    {
      Demo_ID: 'DEMO-05',
      Title: 'Corporate Governance & Executive KPI Cascade',
      Category: 'Leadership & Strategy',
      Operating_Scale: 'C-Suite & HODs',
      Live_URL: 'https://github.com/supranhoo',
      Description: 'Real-time strategic KPI monitoring system linking board mandates to operational outputs.',
      Trigger_Alignment: 'Leadership Disconnect'
    }
  ];
}

/**
 * Seeds default demos into 6_Demos_&_Portfolio tab if empty.
 */
function seedDemoPortfolioSheet() {
  var existing = getSheetDataAsObjects(OPTRA_CONFIG.TABS.DEMOS);
  if (existing.length > 0) return;

  var defaults = getDemoPortfolio();
  for (var i = 0; i < defaults.length; i++) {
    appendSheetRow(OPTRA_CONFIG.TABS.DEMOS, defaults[i]);
  }
}
