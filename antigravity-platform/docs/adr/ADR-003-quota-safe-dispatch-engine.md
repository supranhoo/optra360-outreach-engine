# ADR-003: Quota-Safe Jitter Dispatcher (100/day Limit, 15 Safety Buffer & 240s Ceiling)

**Date:** 2026-10-25  
**Status:** Accepted  
**Deciders:** Automation Engineer, Deliverability Specialist  

## Context
Google monitors outbound email velocity heuristics. Blasting dozens of emails simultaneously from a cold mailbox triggers temporary 24-hour sending bans. Additionally, exceeding Google's daily sending quota causes `MailApp` exceptions, failing active campaigns.

## Decision
We enforce three strict operational guardrails in `04_05_SchedulerAndFollowUp.js`:
1. **Hard Send Ceiling**: Maximum 100 total outbound emails per calendar day (Initial + Follow-ups combined).
2. **Quota Safety Buffer**: If `MailApp.getRemainingDailyQuota() <= 15`, dispatching immediately halts.
3. **Execution Ceiling & Batch Throttling**: Each 15-minute cron execution processes a maximum of 8 emails and bounds execution time to 240 seconds (leaving a 120-second safety cushion before Google's 360-second cutoff).
4. **Randomized Human Jitter**: An artificial delay of 15,000ms – 30,000ms (`Utilities.sleep`) is inserted between sequential emails.
5. **Atomic State Locking**: Before calling GmailApp, the row status is atomically flipped to `SENDING_LOCKED` and committed with `SpreadsheetApp.flush()`.

## Consequences
### Positive
- Zero risk of hitting Google's temporary 24-hour velocity blocks.
- Guaranteed idempotency: Overlapping triggers cannot double-send emails to the same recipient.
- Prevents script timeouts.

### Negative / Trade-offs
- Maximum daily throughput is capped at 100 emails/day, which is ideal for targeted career outreach but unsuitable for bulk spam.
