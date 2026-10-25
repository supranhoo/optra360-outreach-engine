# Project Antigravity: End-to-End Job Outreach Intelligence Platform

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-V8%20Runtime-4285F4?style=flat-square&logo=google&logoColor=white)](https://script.google.com)
[![Google Sheets Relational CRM](https://img.shields.io/badge/Relational%20CRM-14%20Tables-34A853?style=flat-square&logo=googlesheets&logoColor=white)](https://sheets.google.com)
[![Gemini 1.5 Pro](https://img.shields.io/badge/AI%20Copilot-Gemini%201.5%20Pro-8E75B2?style=flat-square&logo=google&logoColor=white)](https://ai.google.dev)
[![Gmail API](https://img.shields.io/badge/Outbound%20MTA-Gmail%20API-EA4335?style=flat-square&logo=gmail&logoColor=white)](https://mail.google.com)
[![clasp](https://img.shields.io/badge/Deploy%20with-clasp-24292e?style=flat-square&logo=github)](https://github.com/google/clasp)

A self-hosted, single-tenant Job Outreach Intelligence Platform engineered entirely on the native **Google Workspace** stack (Google Sheets 14-table CRM, Google Apps Script V8, Gmail API / MailApp, Google Drive, Looker Studio, and Gemini REST API). 

Antigravity transforms a candidate's static portfolio—Resume/CV, LinkedIn Profile, GitHub Repositories, Technical Documentation, and Case Studies—into a structured, relational **Proof-of-Work Knowledge Graph**. By mapping verified engineering and business outcomes directly against an employer's live technical bottlenecks, Antigravity executes hyper-personalized, multi-touch outbound campaigns up to 100 emails per day with thread-locked follow-ups and automated reply/bounce detection.

---

## 1. System Architecture & Web App UI

Antigravity operates as an **independent standalone Google Apps Script Web Application** with an integrated HTML/CSS/JS Command Center:

* **Apps Script Web App**: Runs directly in Google Apps Script via `doGet(e)` / `HtmlService`.
* **Automated CRM Database Binding**: Automatically initializes and connects to its own dedicated Google Sheets CRM (`⚡ Project Antigravity CRM Database`) or any designated spreadsheet ID.
* **Direct Web App URL**: [Open Web App](https://script.google.com/macros/s/AKfycbw3ZJDuex74gJWR1s_brxaDAelpWI_PZVeB4SBsJ9REP_ty8x-cEOyO3ZLm4du57LMkew/exec)
* **Script Editor**: [Open Script Project](https://script.google.com/d/1b1oJEtAVGcB1FeVFJ7V7tWV-Lo-beGkoQgreIryArcDmD9rSfV6Powqf/edit)

---

## 2. Platform Documentation, Specifications & Governance

| Document | File Path | Scope & Authority |
|---|---|---|
| **Product Requirements Document (PRD)** | [`PRD.md`](./PRD.md) | Exhaustive problem statement, candidate personas, user journeys, KPIs, FRs, NFRs, and compliance guardrails. |
| **Single Source of Truth (SSOT)** | [`SSOT.md`](./SSOT.md) | Authoritative reference for 14-table relational schemas, state machine lifecycles, 100-pt scoring math, and pacing limits. |
| **Architecture Decision Records (ADRs)** | [`docs/adr/`](./docs/adr/) | 8 technical decision records documenting architectural choices, tradeoffs, and rationale. |

### Architecture Decision Records Index:
* [ADR-001: Serverless Google Workspace Stack vs Third-Party SaaS Sequencers](./docs/adr/ADR-001-serverless-google-workspace.md)
* [ADR-002: Google Sheets as a 14-Table Relational CRM Database with In-Memory Hash Joins](./docs/adr/ADR-002-relational-crm-sheets.md)
* [ADR-003: Quota-Safe Jitter Dispatcher (100/day Ceiling, 15 Quota Buffer, 240s Ceiling)](./docs/adr/ADR-003-quota-safe-dispatch-engine.md)
* [ADR-004: RFC 2822 In-Reply-To Thread-Locked Follow-Ups via GmailThread API](./docs/adr/ADR-004-thread-locked-followups.md)
* [ADR-005: Zero-Cost DNS MX Verification via Google Public DNS-over-HTTPS](./docs/adr/ADR-005-dns-mx-verification.md)
* [ADR-006: Bipartite Tag Intersection Matching Algorithm for Proof-of-Work Injection](./docs/adr/ADR-006-bipartite-matching-algorithm.md)
* [ADR-007: Gemini 1.5 Pro Structured JSON API Bridge via UrlFetchApp](./docs/adr/ADR-007-gemini-ai-copilot-bridge.md)
* [ADR-008: Standalone Apps Script Web App Architecture with Dynamic Database Provisioning](./docs/adr/ADR-008-standalone-webapp-ui.md)

```
+--------------------------------------------------------------------------------------------------+
|                        CANDIDATE ASSET & PROOF-OF-WORK LAYER (GOOGLE DRIVE)                       |
|   [Master Resume/CV]   [GitHub Repo Index]   [Case Study PDFs]   [Portfolio Links]   [Arch Docs]  |
+------------------------------------------------+-------------------------------------------------+
                                                 | (Metadata & File IDs Indexed)
                                                 v
+--------------------------------------------------------------------------------------------------+
|                    RELATIONAL DATA STORE & CRM UI LAYER (GOOGLE SHEETS - 14 TABS)                |
|  Tab 0: Command_Center (KPI Scorecards, Sparklines, Formulas, Active Deal Pipeline)              |
|  1. Candidate_Profile     5. Target_Companies    8. Email_Campaigns    11. Responses             |
|  2. Projects              6. Opportunities       9. Outreach_History   12. Interview_Pipeline    |
|  3. Work_Experience       7. Contacts           10. Follow_Ups         13. Notes                 |
|  4. Skills                                                             14. KPI_Metrics           |
+------------------------------------------------+-------------------------------------------------+
                                                 | (Read/Write Batch Arrays)
                                                 v
+--------------------------------------------------------------------------------------------------+
|                     APPLICATION & ORCHESTRATION LAYER (GOOGLE APPS SCRIPT V8)                    |
|  • Module 0: Central Config & Schema Bootstrapper (00_Config.js, 00_BootstrapAndMenu.js)        |
|  • Module 1: Contact Manager, DNS MX Verification & 100-Point Scorer (01_ContactManager.js)     |
|  • Module 2 & 3: Campaign Queue Builder & Multi-Persona Personalizer (02_03_CampaignAnd...)     |
|  • Module 4 & 5: Quota-Safe Mail Dispatcher & Thread Follow-Up Engine (04_05_Scheduler...)       |
|  • Module 6, 7 & 8: Inbound Reply/Bounce Sensor, Pipeline Promoter & KPIs (06_07_08_Response...)  |
|  • Module 9: AI Enhancement Layer & Interview Dossier Generator (09_AiCopilotAndDossier.js)     |
|  • Module 10: Trigger Management & Cron Automation (10_TriggersAndAutomations.js)                |
+------------------------------------------------+-------------------------------------------------+
                         |                                          |
     (UrlFetchApp HTTPS) v                                          v (GmailApp / Advanced Services)
+----------------------------------------+     +---------------------------------------------------+
|      EXTERNAL AI INTELLIGENCE LAYER    |     |             COMMUNICATION & BI LAYER              |
|          [Gemini 1.5 Pro REST]         |     |          [Gmail Sent Threads & Inbound]           |
|      - JD Bottleneck Extraction        |     |          [Automated Google Docs Dossiers]         |
|      - Tailored Copy & Hook Synthesis  |     |          [Looker Studio Executive Dashboards]     |
+----------------------------------------+     +---------------------------------------------------+
```

---

## 2. 14-Table Relational Schema Architecture

| # | Tab Name | Role & Description | Primary Key | Foreign Keys |
|---|---|---|---|---|
| **0** | `Command_Center` | Live KPI scorecards, formulas, velocity sparkline, and actionable interview deals | - | Pulls live formulas from tabs |
| **1** | `Candidate_Profile` | Master candidate positioning, headline, URLs, drive IDs, and daily send limits | `candidate_id` | - |
| **2** | `Projects` | Proof-of-work repositories, case studies, problem domain tags, tech stack tags, and ROI impact metrics | `project_id` | `candidate_id` &rarr; Tab 1 |
| **3** | `Work_Experience` | Quantified enterprise career achievements, challenges, architectural actions, and verification links | `experience_id` | `candidate_id` &rarr; Tab 1 |
| **4** | `Skills` | Competency taxonomy, categories, proficiency scores, and proof mappings | `skill_id` | `primary_project_fk`, `primary_exp_fk` |
| **5** | `Target_Companies` | Account intelligence, funding stage, headcount, news triggers, and inferred pain points | `company_id` | - |
| **6** | `Opportunities` | Job descriptions, unposted expansion hypotheses, AI pain points, and matched candidate project | `opportunity_id` | `company_id` &rarr; Tab 5, `matched_project_1_fk` &rarr; Tab 2 |
| **7** | `Contacts` | Decision makers, founders, VPs, recruiters, DNS MX verification status, lead score, and outreach state | `contact_id` | `company_id` &rarr; Tab 5, `opportunity_id` &rarr; Tab 6 |
| **8** | `Email_Campaigns` | Multi-day sequence orchestration, persona tracks, daily caps, and timezone windows | `campaign_id` | - |
| **9** | `Outreach_History` | Immutable transaction log of scheduled & dispatched initial emails with thread IDs and atomic status | `outreach_id` | `campaign_id` &rarr; Tab 8, `contact_id` &rarr; Tab 7 |
| **10** | `Follow_Ups` | Thread-locked multi-touch sequence queue (Stages 1–3), angles, scheduled business dates, and auto-cancellation | `followup_id` | `outreach_id` &rarr; Tab 9, `contact_id` &rarr; Tab 7 |
| **11** | `Responses` | Inbound reply detection, sentiment classification, reply latency, and action required flags | `response_id` | `outreach_id` &rarr; Tab 9, `contact_id` &rarr; Tab 7 |
| **12** | `Interview_Pipeline` | Active deal stages (`1_INTRO_CALL` &rarr; `5_OFFER`), meeting times, prep dossiers, and win probabilities | `pipeline_id` | `company_id` &rarr; Tab 5, `contact_id` &rarr; Tab 7 |
| **13** | `Notes` | Polymorphic qualitative stakeholder intelligence and call notes | `note_id` | `entity_fk_id` (polymorphic) |
| **14** | `KPI_Metrics` | Daily time-series rollups, sent counts, reply rates, positive replies, and top-performing projects | `metric_date` | `top_performing_project_fk` &rarr; Tab 2 |

---

## 3. 6 Multi-Persona Template Families

Antigravity provides 6 purpose-built outbound copy tracks interpolated with dynamic candidate tokens:
1. **Founder / CEO Outreach**: Focuses on speed-to-market, runway protection, and custom ROI hypotheses. Low-friction 10-minute builder CTA.
2. **Technical Leadership (CTO / VP Eng / Staff Lead)**: Emphasizes concurrency, architectural tradeoffs, production benchmarks, and peer repo code-review CTA.
3. **Hiring Manager**: De-risks the 90-day roadmap with dual proof points and near-zero ramp-up time.
4. **Recruiter / Talent Partner**: Provides a 30-second bulletproof fit checklist (Core Stack, Verified Impact, Live Proof-of-Work, Tailored CV).
5. **Peer / Alumni Referral Request**: Low-friction inquiry about engineering culture and autonomy, paired with a 3-bullet blurb for easy internal forwarding.
6. **General Initial Problem-First Pitch**: Anchored directly around the candidate's highest-matching open-source project and quantified business metrics.

---

## 4. Quota Safety & Deliverability Guardrails

* **Daily Send Hard Ceiling**: 100 outbound emails/day max (Initial + Follow-ups combined).
* **Safety Quota Buffer**: Execution halts if Google's remaining daily quota drops to or below 15.
* **Execution Timeout Prevention**: Batched to max 8 emails per 15-minute trigger run (keeps execution under 240 seconds, well below Google's 360-second limit).
* **Atomic Concurrency Locking**: Acquires `LockService.getScriptLock()` and atomically sets row status to `SENDING_LOCKED` with `SpreadsheetApp.flush()` before dispatching.
* **Human Jitter**: 15,000ms – 30,000ms randomized sleep between outbound sends.
* **Zero-Cost DNS MX Verification**: Queries Google Public DNS-over-HTTPS (`https://dns.google/resolve?name=domain&type=MX`) to verify target domain mail servers before queueing.
* **Thread-Locked Follow-Ups**: Uses `GmailThread.reply()` to guarantee that multi-stage follow-ups nest inside the exact same email conversation thread with proper RFC 2822 `In-Reply-To` and `References` headers.
* **Automated Circuit Breaker**: Automatically pauses all campaigns if daily bounces reach 3.

---

## 5. Daily 20-Minute Operating Cadence

1. **08:30 AM – Pipeline Review (5 mins)**: Open the `Command_Center` sheet; inspect overnight positive replies and review auto-generated Google Docs interview prep dossiers.
2. **08:35 AM – Lead Ingestion (10 mins)**: Add new target companies and decision makers into `Contacts` and click **🚀 Antigravity Platform &rarr; 🎯 2. Verify Emails & Score Leads**.
3. **08:45 AM – Draft QA & Approval (5 mins)**: Scan generated hooks in `Outreach_History`; flip approved rows from `DRAFT_READY` to `APPROVED_QUEUED`.
4. **Autonomous Execution**: Background 15-minute triggers handle all dispatching, thread tracking, reply parsing, and follow-up suppression automatically.

---

## 6. Deployment & Installation Guide

### Prerequisites
* Google Workspace or personal Gmail account
* Target Google Sheet
* Google Apps Script command-line tool (`@google/clasp`)

### Pushing Code to Google Apps Script
```bash
# From within the antigravity-platform directory:
cd "c:\Users\ANKIT CHOUDHARY\OneDrive\Project working\Self work\antigravity-platform"

# Push all files to Google Apps Script
clasp push
```

### Initializing the Platform in Google Sheets
1. Open your master Google Sheet.
2. Refresh the browser to load the **🚀 Antigravity Platform** custom menu.
3. Click **🚀 Antigravity Platform &rarr; 🛠 1. Bootstrap 14-Table CRM & Command Center**.
4. Authorize Google Workspace permissions when prompted.
5. In Google Apps Script Editor (**Extensions &rarr; Apps Script**), navigate to **Project Settings &rarr; Script Properties** and add:
   * `GEMINI_API_KEY`: Your Google AI Studio API key.
6. Return to Google Sheets and click **🚀 Antigravity Platform &rarr; ⏰ 7. Install 15-Minute Automation Triggers**.

---

## 7. Verification & Offline Testing

Antigravity includes a comprehensive Node.js unit test suite:

```bash
node tests/test_antigravity.js
```

### Test Coverage:
- RFC 5322 email syntax validation & disposable/system prefix filter
- 100-point deterministic lead scoring math and edge cases
- Bipartite tag intersection and candidate project matching
- 6 multi-persona email template renderers & UTM telemetry generation
- Business day calendar arithmetic skipping weekends
- Inbound response sentiment classification rules (Opt-out, Referral, Interview, Not-Now)
