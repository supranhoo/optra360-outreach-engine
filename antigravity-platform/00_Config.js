/**
 * PROJECT ANTIGRAVITY: END-TO-END JOB OUTREACH INTELLIGENCE PLATFORM
 * Module 0: Central Configuration, Table Schemas & Platform Constants
 * 
 * Target Stack: Google Apps Script (V8 Runtime) • Google Sheets Relational CRM • Gmail API • Google Drive • Gemini API
 */

var ANTIGRAVITY_CONFIG = {
  // Operational Quota & Safety Limits
  DAILY_HARD_LIMIT: 100,          // Maximum total outbound sends per calendar day (Initial + Follow-ups)
  SAFETY_QUOTA_BUFFER: 15,       // Halt dispatch if Google MailApp remaining quota drops to or below this
  MAX_EXECUTION_MS: 240000,      // 4 minutes safe ceiling (well below Google's 360s hard ceiling)
  BATCH_DISPATCH_LIMIT: 8,       // Max emails dispatched per 15-minute trigger run to avoid timeouts
  JITTER_MIN_MS: 15000,          // 15 seconds minimum human delay between outbound sends
  JITTER_MAX_MS: 30000,          // 30 seconds maximum human delay
  STAGGER_INTERVAL_MINS: 7,      // 7 minutes stagger between scheduled candidate outreach slots
  BUSINESS_HOURS_START: 8,       // 08:00 AM local prospect time
  BUSINESS_HOURS_END: 17,        // 05:00 PM local prospect time
  LEAD_SCORE_QUALIFY_THRESHOLD: 65, // Minimum score to enter APPROVED_QUEUED / QUEUED state
  MAX_CONTACTS_PER_COMPANY_7D: 2,   // Anti-collision guard: Max 2 contacts per target domain per 7-day window
  MAX_CATCHALL_CONTACTS_WEEK: 3,    // Max catch-all domain contacts per week to protect deliverability
  BOUNCE_CIRCUIT_BREAKER_LIMIT: 3,  // If daily bounces reach 3, automatically pause all active campaigns

  // Active User Email Resolution
  getCandidateEmail: function() {
    try {
      const email = (typeof Session !== 'undefined' && Session.getActiveUser) 
        ? Session.getActiveUser().getEmail() 
        : '';
      return email || 'ankitchoudharymba@gmail.com';
    } catch (e) {
      return 'ankitchoudharymba@gmail.com';
    }
  },

  // Default Portfolio, Video & CV URLs
  DEFAULT_PORTFOLIO_VIDEO_URL: 'https://linkedin.com/in/ankit-choudhary-human-resources',
  DEFAULT_RESUME_URL: 'https://linkedin.com/in/ankit-choudhary-human-resources',
  DEFAULT_CALENDLY_URL: 'https://calendly.com/ankitchoudharymba/hr-leadership-discussion',

  // Dynamic Relational Spreadsheet Resolution (Standalone Apps Script Compatible)
  DEFAULT_SPREADSHEET_ID: '1KEJfOBK8PobuGQ6TI-91AKebXehOphDD3LwQw_zaXwU',

  getSpreadsheet: function() {
    if (typeof PropertiesService === 'undefined' || !PropertiesService.getScriptProperties) {
      return null;
    }
    const props = PropertiesService.getScriptProperties();
    let sheetId = props.getProperty('SPREADSHEET_ID') || this.DEFAULT_SPREADSHEET_ID;

    if (sheetId && typeof SpreadsheetApp !== 'undefined') {
      try {
        const ss = SpreadsheetApp.openById(sheetId);
        if (ss) {
          props.setProperty('SPREADSHEET_ID', sheetId);
          return ss;
        }
      } catch (e) {
        console.warn('Could not open spreadsheet by ID ' + sheetId + ': ' + e.message);
      }
    }

    if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet) {
      try {
        const active = SpreadsheetApp.getActiveSpreadsheet();
        if (active) {
          props.setProperty('SPREADSHEET_ID', active.getId());
          return active;
        }
      } catch (e) {}
    }

    // Auto-create standalone CRM database spreadsheet if none exists
    if (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.create) {
      const newSs = SpreadsheetApp.create('⚡ Project Antigravity CRM Database');
      props.setProperty('SPREADSHEET_ID', newSs.getId());
      console.log('Created dedicated Antigravity CRM Spreadsheet: ' + newSs.getUrl());
      return newSs;
    }

    return null;
  },

  getSpreadsheetUrl: function() {
    try {
      const ss = this.getSpreadsheet();
      return ss ? ss.getUrl() : '';
    } catch (e) {
      return '';
    }
  },

  // 14 Relational Sheets + Tab 0 Command Center
  SHEETS: {
    COMMAND_CENTER: 'Command_Center',
    CANDIDATE_PROFILE: 'Candidate_Profile',
    PROJECTS: 'Projects',
    WORK_EXPERIENCE: 'Work_Experience',
    SKILLS: 'Skills',
    TARGET_COMPANIES: 'Target_Companies',
    OPPORTUNITIES: 'Opportunities',
    CONTACTS: 'Contacts',
    EMAIL_CAMPAIGNS: 'Email_Campaigns',
    OUTREACH_HISTORY: 'Outreach_History',
    FOLLOW_UPS: 'Follow_Ups',
    RESPONSES: 'Responses',
    INTERVIEW_PIPELINE: 'Interview_Pipeline',
    NOTES: 'Notes',
    KPI_METRICS: 'KPI_Metrics'
  },

  // Table Schema Column Specifications (Frozen Row 1 snake_case headers)
  SCHEMAS: {
    Candidate_Profile: [
      'candidate_id', 'full_name', 'headline_positioning', 'years_experience',
      'core_value_pillars', 'portfolio_base_url', 'github_profile_url',
      'linkedin_url', 'master_resume_drive_id', 'calendar_booking_url', 'daily_send_limit'
    ],
    Projects: [
      'project_id', 'candidate_id', 'project_title', 'project_type',
      'problem_domain_tags', 'tech_stack_tags', 'architecture_summary',
      'quantified_business_impact', 'github_repo_url', 'case_study_url', 'target_persona_fit'
    ],
    Work_Experience: [
      'experience_id', 'candidate_id', 'company_name', 'role_title',
      'industry_vertical', 'challenge_context', 'action_executed',
      'roi_metric_statement', 'verification_url'
    ],
    Skills: [
      'skill_id', 'skill_name', 'skill_category', 'proficiency_score',
      'primary_project_fk', 'primary_exp_fk', 'synonym_keywords'
    ],
    Target_Companies: [
      'company_id', 'company_name', 'website_domain', 'industry_vertical',
      'funding_stage_headcount', 'recent_trigger_event', 'inferred_pain_points',
      'company_fit_score', 'careers_page_url'
    ],
    Opportunities: [
      'opportunity_id', 'company_id', 'role_title', 'opportunity_type',
      'jd_source_url', 'raw_jd_text', 'extracted_pain_points',
      'required_skills_tags', 'matched_project_1_fk', 'custom_roi_hypothesis'
    ],
    Contacts: [
      'contact_id', 'company_id', 'opportunity_id', 'first_name', 'last_name',
      'email_address', 'verification_status', 'persona_type', 'role_title',
      'linkedin_url', 'timezone_offset', 'lead_score', 'outreach_status'
    ],
    Email_Campaigns: [
      'campaign_id', 'campaign_name', 'target_persona', 'template_family_id',
      'daily_cap_allocation', 'send_window_start_hr', 'send_window_end_hr',
      'followup_1_delay_days', 'followup_2_delay_days', 'campaign_status'
    ],
    Outreach_History: [
      'outreach_id', 'campaign_id', 'contact_id', 'opportunity_id', 'matched_project_id',
      'subject_line', 'email_body_html', 'scheduled_send_at', 'actual_sent_at',
      'gmail_thread_id', 'gmail_message_id', 'dispatch_status'
    ],
    Follow_Ups: [
      'followup_id', 'outreach_id', 'contact_id', 'gmail_thread_id', 'step_number',
      'followup_angle', 'followup_body_html', 'scheduled_for', 'sent_at', 'followup_status'
    ],
    Responses: [
      'response_id', 'outreach_id', 'contact_id', 'gmail_thread_id',
      'received_at', 'reply_snippet', 'sentiment_category', 'time_to_reply_hours', 'action_required'
    ],
    Interview_Pipeline: [
      'pipeline_id', 'company_id', 'opportunity_id', 'primary_contact_id',
      'current_stage', 'next_meeting_datetime', 'prep_doc_drive_url',
      'comp_target_range', 'win_probability_pct'
    ],
    Notes: [
      'note_id', 'entity_type', 'entity_fk_id', 'created_at', 'note_author', 'note_content'
    ],
    KPI_Metrics: [
      'metric_date', 'initial_emails_sent', 'followup_emails_sent', 'total_emails_sent',
      'bounces_detected', 'total_replies_received', 'positive_replies_count',
      'interviews_booked_count', 'rolling_7d_reply_rate', 'top_performing_project_fk'
    ]
  },

  // Scoring Weights & Allocation Rules
  SCORING: {
    PERSONA_WEIGHTS: {
      FOUNDER: 20,    // Promoter, Managing Director, Owner (Decision-maker in mid-market)
      HIRING_MGR: 20, // CHRO, VP HR, Head HR (First institutional HR leadership)
      TECH_LEAD: 18,  // COO, Plant Head, Works Director (Shopfloor operations leader)
      RECRUITER: 12,
      PEER: 10
    },
    STACK_MAX_POINTS: 25,
    STACK_POINTS_PER_TAG: 5,
    URGENCY_MAX_POINTS: 15,
    VERIFICATION_POINTS: {
      VERIFIED_VALID: 10,
      CATCH_ALL: 0,
      INVALID: 0
    },
    DYNAMIC_SIGNALS: {
      LINK_CLICKED: 15,
      POSITIVE_REPLY: 30,
      REFERRAL_FORWARD: 20,
      BOUNCE_OR_UNSUB: 0
    }
  },

  // Valid ENUM Values for Data Validation Dropdowns
  ENUMS: {
    PROJECT_TYPE: ['GITHUB_REPO', 'CASE_STUDY', 'LIVE_SAAS'],
    PERSONA_TYPE: ['FOUNDER', 'HIRING_MGR', 'TECH_LEAD', 'RECRUITER', 'PEER'],
    SKILL_CATEGORY: ['BACKEND', 'AI_ML', 'INFRA', 'PRODUCT', 'GTM'],
    OPPORTUNITY_TYPE: ['POSTED_JD', 'UNPOSTED_EXPANSION', 'EXECUTIVE_OUTREACH', 'FRACTIONAL'],
    VERIFICATION_STATUS: ['VERIFIED_VALID', 'UNVERIFIED', 'CATCH_ALL', 'INVALID'],
    OUTREACH_STATUS: ['NEW', 'QUEUED', 'IN_SEQUENCE', 'LOW_SCORE_HOLD', 'REPLIED', 'OPT_OUT', 'BOUNCED', 'DUPLICATE_SKIPPED'],
    CAMPAIGN_STATUS: ['ACTIVE', 'PAUSED', 'COMPLETED'],
    DISPATCH_STATUS: ['DRAFT_READY', 'APPROVED_QUEUED', 'SENDING_LOCKED', 'SENT', 'BOUNCED', 'FAILED'],
    FOLLOWUP_STATUS: ['PENDING', 'SENT', 'CANCELLED_REPLIED', 'CANCELLED_BOUNCED'],
    FOLLOWUP_ANGLE: ['ARCHITECTURE_DIAGRAM', 'CASE_STUDY_DROP', 'SOFT_BREAKUP'],
    SENTIMENT_CATEGORY: ['POSITIVE_INTERVIEW', 'REFERRAL_FORWARD', 'OOO_AUTO', 'NOT_NOW', 'UNSUBSCRIBE'],
    PIPELINE_STAGE: ['1_INTRO_CALL', '2_TECH_SCREEN', '3_SYSTEM_DESIGN', '4_EXEC_FINAL', '5_OFFER', 'CLOSED'],
    ENTITY_TYPE: ['COMPANY', 'CONTACT', 'PIPELINE'],
    NOTE_AUTHOR: ['CANDIDATE', 'AI_COPILOT']
  }
};

// Node.js export compatibility for offline testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = ANTIGRAVITY_CONFIG;
}
