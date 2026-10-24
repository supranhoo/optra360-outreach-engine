/**
 * Optra360 & Zoptro Automated Outreach & Opportunity Engine
 * 
 * 00_Config.js: System constants, brand parameters, and configuration settings.
 */

var OPTRA_CONFIG = {
  // Google Drive & Workspace Document IDs
  SPREADSHEET_ID: '1KEJfOBK8PobuGQ6TI-91AKebXehOphDD3LwQw_zaXwU',
  SLIDES_TEASER_ID: '1S-BHa8yngtruRdkvhZf9EdNDlMYHHD6k8vdpaS_RJ9Q',
  SLIDES_DETAILED_ID: '1MLtvM1TWEl7hBZ-xI5MTQ73mI0irUipKzzFU8UkEJ10',

  // Google Sheets Tab Names
  TABS: {
    PROFILES: '1_Profiles_Config',
    TARGET_COMPANIES: '2_Target_Companies',
    PROBLEM_SOLUTION: '3_Problem_Solution',
    CONTACTS: '4_Contacts_Outreach',
    DECK_TEMPLATES: '5_Deck_Templates',
    DEMOS: '6_Demos_&_Portfolio',
    KPI: '7_KPI_Dashboard'
  },

  // Optra360 Brand Expression Guidelines
  BRAND: {
    NAME: 'Optra360',
    TAGLINE: 'Prove Before You Promise',
    COLORS: {
      AUTHORITY_BLUE: '#09293D',
      CLARITY_IVORY: '#F4F2EA',
      GROUNDED_OLIVE: '#666D57',
      PROGRESS_ORANGE: '#F65800',
      WHITE: '#FFFFFF',
      MUTED_GRAY: '#6B7280',
      BORDER_GRAY: '#E5E7EB'
    },
    FONTS: {
      HEADING: 'Outfit, sans-serif',
      BODY: 'Figtree, sans-serif'
    },
    RADIUS: '0px' // Sharp 90-degree corners per brand expression rules
  },

  // Optra360 5 Core Buying Triggers
  BUYING_TRIGGERS: [
    'Process Breakdown',
    'Performance Stagnation',
    'Leadership Disconnect',
    'System Fragmentation',
    'Workforce Misalignment'
  ],

  // Optra360 6-Stage Delivery Method
  DELIVERY_STAGES: [
    'Diagnose',
    'Design',
    'Digitize',
    'Deploy',
    'Stabilize',
    'Govern'
  ],

  // Target Sectors
  SECTORS: [
    'Manufacturing',
    'Steel',
    'Cement',
    'Power',
    'Pharma',
    'Mining',
    'Automotive'
  ],

  // Script Property Storage Keys
  PROPERTY_KEYS: {
    ACTIVE_PROFILE: 'OPTRA_ACTIVE_PROFILE_ID',
    GEMINI_KEY: 'OPTRA_GEMINI_API_KEY',
    APOLLO_KEY: 'OPTRA_APOLLO_API_KEY',
    HUNTER_KEY: 'OPTRA_HUNTER_API_KEY',
    SEARCH_KEY: 'OPTRA_SEARCH_API_KEY',
    SEARCH_CX: 'OPTRA_SEARCH_CX',
    DAILY_LIMIT: 'OPTRA_DAILY_LIMIT'
  },

  // Default Email Batch Limits
  LIMITS: {
    DEFAULT_DAILY_EMAILS: 30,
    PERMUTATION_MAX_PER_DOMAIN: 8
  }
};
