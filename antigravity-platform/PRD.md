# Product Requirements Document (PRD)
## Project Antigravity: End-to-End Job Outreach Intelligence Platform

**Document Version:** 1.0.0  
**Status:** Approved & Implemented  
**Owner:** Ankit Choudhary  
**Target Stack:** Google Apps Script (V8 Runtime) • Google Sheets Relational CRM • Gmail API / MailApp • Google Drive • Looker Studio • Gemini 1.5 Pro REST API  

---

## 1. Executive Summary

Project Antigravity is a self-hosted, single-tenant Job Outreach Intelligence Platform engineered entirely on the native Google Workspace stack (Google Sheets 14-table Relational CRM, Google Apps Script V8, Gmail API, Google Drive, Google Contacts, Looker Studio, and Gemini REST API). 

Traditional career acquisition relies on reactive application portals where Applicant Tracking Systems (ATS) reject up to 75% of qualified technical practitioners through keyword heuristics. Antigravity fundamentally inverts this paradigm: it converts a candidate's static portfolio—Resume/CV, LinkedIn Profile, GitHub Repositories, Technical Whitepapers, System Design Docs, and Business Case Studies—into an actionable, relational **Proof-of-Work Knowledge Graph**. 

By mapping verified architectural achievements directly against an employer's live technical bottlenecks, funding catalysts, and open engineering requisitions, Antigravity executes hyper-personalized, multi-touch outbound email campaigns up to 100 emails per day, positioning the candidate as a high-leverage business asset and problem solver rather than a commodity applicant.

---

## 2. Problem Statement

Experienced senior software engineers, systems architects, engineering leaders, and technical product builders face four structural market failures when relying on standard job boards and resume submissions:

1. **The ATS Signal-to-Noise Collapse**: One-click apply buttons flood job openings with 500 to 2,000+ applicants within 48 hours. Standard resumes flatten distributed systems achievements, production code repositories, and quantified revenue/cost impacts into generic bullet points that fail to differentiate top-decile talent.
2. **Asymmetric Decision-Maker Access**: Non-technical HR screeners evaluate candidates through rigid keyword matching, whereas Founders, CTOs, VPs of Engineering, and Hiring Managers evaluate talent on their capability to de-risk roadmap milestones, solve latency/throughput bottlenecks, or eliminate cloud spend overhead.
3. **Manual Personalization Bottlenecks**: Deeply researching a company's architecture, recent press releases, and engineering blog posts to craft a bespoke pitch referencing relevant candidate GitHub code takes 30–45 minutes per prospect. Achieving consistent 50–100/day outreach is mathematically impossible without automation.
4. **Fragmented Follow-Up & Pipeline Leakage**: Over 60% of positive executive responses occur between touchpoints 2 and 4. Without automated thread monitoring, NDR bounce interception, and deterministic reply suppression, candidates either drop warm leads or accidentally send tone-deaf follow-ups after an executive has already responded.

---

## 3. User Personas

| Persona Name & Archetype | Core Assets Available | Primary Outreach Targets | Primary Value Proposition & Objective |
|---|---|---|---|
| **1. The Systems Builder**<br>*(Senior Software / AI / Data / Infrastructure Engineer)* | Active GitHub repos, system architecture diagrams, benchmark docs, technical blog posts, CV. | CTOs, VPs of Engineering, Staff Engineers, Engineering Managers at Series A–D startups. | Bypass HR/ATS filters by proving production code, latency reductions, and architectural solutions tailored to the company's tech stack. |
| **2. The Growth & Product Operator**<br>*(Senior PM, Growth Lead, GTM Architect)* | Detailed product teardown case studies, Figma prototypes, revenue attribution models, live portfolio. | Founders, CEOs, CPOs, VPs of Product, Heads of Growth. | Pitch concrete 30-60-90 day growth hypotheses, funnel audits, and feature ROI backed by prior shipped products. |
| **3. The Fractional / Consulting Technologist**<br>*(Principal Architect / Advisory Lead)* | Client ROI case studies, open-source tooling, implementation playbooks, thought leadership. | COOs, Managing Directors, Department Heads, Technical Founders. | Generate high-ticket full-time or fractional advisory interviews by quantifying immediate operational cost savings and automation ROI. |

---

## 4. End-to-End User Journey

1. **Asset Ingestion & Indexing (Day 1)**: The candidate populates `Candidate_Profile`, `Projects`, `Work_Experience`, and `Skills` in Google Sheets, linking specific GitHub URLs, live demos, and quantified business impact metrics to problem-domain tags (e.g., `#LLMLatency`, `#HighConcurrency`, `#FinOps`, `#KafkaEventStreaming`).
2. **Compliant Lead Discovery & Enrichment (Daily Batch)**: The candidate imports target companies, job openings, and decision-makers into `Target_Companies`, `Opportunities`, and `Contacts`.
3. **AI Pain-Point Mapping & Value Synthesis**: Apps Script invokes the Personalization Engine, parsing the JD and news triggers, matching the highest-scoring candidate project via bipartite tag intersection, and synthesizing a tailored subject line and HTML email draft.
4. **Human-in-the-Loop QA & Queue Approval (5 mins daily)**: The candidate inspects generated drafts in the Web App UI, reviews proof links, and flips rows from `DRAFT_READY` to `APPROVED_QUEUED`.
5. **Autonomous Dispatch & Thread-Locked Follow-Ups**: Background 15-minute triggers dispatch batches across business hours with randomized jitter (15s–30s). When prospects reply, follow-ups halt immediately, lead scores increase by +30, and deals advance to `Interview_Pipeline` with auto-generated Google Docs prep dossiers.

---

## 5. Business Goals & Success Metrics (KPIs)

| Metric Category | Metric Name | Baseline (Cold Apply) | Antigravity Target Benchmark | Measurement Formula in Platform |
|---|---|---|---|---|
| **Volume & Pacing** | Daily Safe Dispatch Rate | 5–10 manual emails/day | 50–100 personalized emails/day | `COUNTIFS(Outreach_History!sent_at, TODAY())` |
| **Deliverability** | Hard + Soft Bounce Rate | 5%–12% (unverified) | < 1.8% (strict threshold) | `Bounced_Emails / Total_Emails_Sent` |
| **Engagement** | Overall Thread Reply Rate | 1.5%–3.0% | 12.0%–22.0% | `Unique_Contacts_Replied / Unique_Contacts_Emailed` |
| **Conversion (North Star)**| Positive Reply / Intro Call Rate | 0.5%–1.0% | 4.5%–8.5% | `Positive_Sentiment_Replies / Unique_Contacts_Emailed` |
| **Asset Attribution** | Proof-of-Work Click-Through | Untracked | > 25% of positive replies cite project | UTM telemetry (`?utm_source=antigravity`) + reply body parse |
| **Pipeline Velocity** | First Touch to Technical Screen | 21–45 days | 7–14 days | `AVG(Interview_Pipeline!screen_date - Outreach_History!first_sent)` |

---

## 6. Functional Requirements (FR)

* **FR-01 Relational Asset Repository**: Normalized storage for candidate profile data, GitHub repos, live case study links, and quantifiable impact metrics categorized by problem tags.
* **FR-02 Multi-Persona Template Engine**: 6 distinct persona outreach tracks (Founder, Hiring Manager, Technical Leadership, Recruiter, Peer, Initial Pitch) with token interpolation and dynamic proof blocks.
* **FR-03 Quota-Aware Jitter Scheduler**: Dispatches up to 100 emails/day during prospect business hours (08:00–17:00 local time) in max 8-email batches with 15–30s randomized human jitter while maintaining a 15-email safety buffer against `MailApp.getRemainingDailyQuota()`.
* **FR-04 Thread-Locked Follow-Up Engine**: Dispatches multi-stage follow-ups (Stage 1 Day +3 architecture benchmark; Stage 2 Day +7 soft breakup) inside the exact same Gmail thread via `thread.reply()`, maintaining RFC 2822 `In-Reply-To` and `References` headers.
* **FR-05 Bidirectional Response Detection & Bounce Guard**: Automatically polls Gmail sent threads every 15 minutes to detect prospect replies, vacation auto-responders, and NDR bounces (`mailer-daemon`), immediately cancelling pending follow-ups.
* **FR-06 Standalone Apps Script Web App UI**: A standalone browser-based Command Center running via `HtmlService.createHtmlOutputFromFile('WebAppUI')` providing real-time telemetry, lead scoring controls, draft preview drawer, and trigger management.
* **FR-07 Apollo.io Executive Lead Discovery & Ingestion Engine**: Native integration with Apollo.io REST API (`/mixed_people/search`, `/people/match`) and CSV ingestion parser, allowing 1-click executive discovery across target companies and roles, verified email extraction, automatic relational CRM ingestion, and immediate 100-point lead fit scoring.

---

## 7. Non-Functional Requirements (NFR)

* **NFR-01 Execution Time Resilience**: Every Apps Script trigger execution must complete in under 240 seconds (well below Google's 360-second execution ceiling) via atomic batch slicing and state tracking.
* **NFR-02 Concurrency & Idempotency**: Concurrency protection enforced via `LockService.getScriptLock()` and atomic row locking (`SENDING_LOCKED` &rarr; `SENT`) with `SpreadsheetApp.flush()`.
* **NFR-03 Zero External Infrastructure Cost**: Operates 100% within standard Google Workspace storage, compute, and Gmail limits without requiring paid SaaS sequencers.

---

## 8. Compliance, Privacy & Deliverability Guardrails

* **Regulatory Compliance**: Operates under GDPR Art. 6(1)(f) Legitimate Interests, CAN-SPAM, and India DPDP Act. One-to-one professional communications with clear opt-out phrasing.
* **Instant Opt-Out Honoring**: Any reply expressing unsubscribe intent immediately updates `Contacts.outreach_status = 'OPT_OUT'` and permanently cancels all future campaign communications.
* **Automated Circuit Breaker**: If daily bounces reach 3 in `KPI_Metrics`, all active campaigns automatically pause to protect domain reputation.
