# ADR-005: Zero-Cost DNS MX Verification via Google Public DNS-over-HTTPS

**Date:** 2026-10-25  
**Status:** Accepted  
**Deciders:** Deliverability Specialist, Automation Engineer  

## Context
Sending cold emails to invalid, defunct, or unroutable email addresses results in high hard-bounce rates (>5%). Exceeding a 2% bounce rate triggers Google Workspace anti-spam penalties and domain reputation degradation. Third-party email verification APIs (e.g., ZeroBounce, NeverBounce) cost $0.005–$0.01 per verification and require external API keys.

## Decision
We implement a zero-cost DNS MX verification gate in `01_ContactManager.js`:
1. Validate RFC 5322 email syntax with strict regular expressions.
2. Filter out generic administrative prefixes (`noreply@`, `mailer-daemon@`, `support@`, `info@`).
3. Query Google Public DNS-over-HTTPS (`https://dns.google/resolve?name=${domain}&type=MX`) via `UrlFetchApp` to confirm the domain possesses active, valid mail exchanger records.
4. Contacts on domains lacking MX records are immediately marked `INVALID` and set to `BOUNCED` status, preventing them from ever entering campaign queues.

## Consequences
### Positive
- Zero API cost ($0.00) with no third-party credits required.
- Hard bounce rate reliably maintained below the 1.8% threshold.
- Fast execution with high caching efficiency across standard company domains.

### Negative / Trade-offs
- DNS MX checks confirm domain deliverability, but do not perform deep SMTP RCPT TO handshakes for catch-all domains. Catch-all domains are capped at 3 per week to mitigate risk.
