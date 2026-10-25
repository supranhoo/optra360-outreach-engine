# ADR-004: RFC 2822 In-Reply-To Thread-Locked Follow-Ups via GmailThread API

**Date:** 2026-10-25  
**Status:** Accepted  
**Deciders:** Recruiting Strategist, Apps Script Developer  

## Context
When candidates send follow-up emails as separate, newly generated message threads with generic subjects (e.g., "Follow up"), recipients lose context, and corporate email clients fragment the interaction. Executives prefer follow-ups that appear directly within the original message thread with full historical context.

## Decision
We enforce thread-locked follow-ups across the platform:
1. Upon dispatching an initial email, `GmailApp.search('to:... subject:...')` queries the sent message and extracts its exact `gmail_thread_id` and RFC 2822 `message_id`, persisting both in `Outreach_History`.
2. Follow-up tasks in `Follow_Ups` store this `gmail_thread_id`.
3. When executing follow-ups, `GmailApp.getThreadById(gmail_thread_id)` retrieves the live thread and dispatches the follow-up using `thread.reply('', {htmlBody: body})`.
4. Prior to dispatching any follow-up, the scanner verifies the thread for any prospect replies. If a human reply is found, pending follow-ups are cancelled immediately (`CANCELLED_REPLIED`).

## Consequences
### Positive
- Strict RFC 2822 adherence with `In-Reply-To` and `References` headers intact.
- Follow-ups nest directly in the recipient's primary Gmail/Outlook conversation thread.
- Zero risk of embarrassing duplicate follow-ups after a prospect has already replied.

### Negative / Trade-offs
- Requires capturing the thread ID immediately after send, introducing a 2.5-second indexing pause.
