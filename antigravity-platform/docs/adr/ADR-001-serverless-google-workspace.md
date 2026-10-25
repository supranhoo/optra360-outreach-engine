# ADR-001: Serverless Google Workspace Architecture vs Third-Party SaaS Sequencers

**Date:** 2026-10-25  
**Status:** Accepted  
**Deciders:** Solution Architect, Engineering Lead  

## Context
Traditional cold email outreach platforms (e.g., Lemlist, Apollo, Instantly) require recurring paid subscriptions ($50–$300/month), store candidate private data and prospect lists on third-party servers, and often route emails through shared third-party SMTP servers with erratic IP reputation. Technical job seekers and candidates need an ultra-low-cost, private, single-tenant system operating natively through their own verified Google Workspace or Gmail account.

## Decision
We architected Project Antigravity 100% on the native Google Workspace stack:
- Compute: Google Apps Script (V8 runtime serverless engine)
- Data Store: Google Sheets (14 normalized tables + Tab 0 Command Center)
- Outbound Mail: Gmail API / MailApp
- Storage: Google Drive (PDF resumes, whitepapers, case study docs)
- Document Generation: Google Docs via DocumentApp (AI interview dossiers)
- User Interface: Standalone Web App via HtmlService

## Consequences
### Positive
- Zero recurring infrastructure or SaaS subscription cost ($0/month).
- 100% single-tenant data isolation — no third-party database stores candidate resumes or prospect contact details.
- High deliverability: Emails are sent directly from the authenticated user's genuine Gmail mailbox rather than a shared bulk pool.

### Negative / Trade-offs
- Execution time limits: Google imposes a 360s hard ceiling per execution (mitigated by batching max 8 emails per run).
- Daily sending limits: Google imposes a 100–1,500/day recipient quota (mitigated by the 100/day hard cap and 15-email safety buffer).
