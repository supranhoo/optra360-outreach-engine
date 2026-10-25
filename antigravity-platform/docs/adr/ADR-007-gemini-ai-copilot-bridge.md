# ADR-007: Gemini 1.5 Pro Structured JSON API Bridge via UrlFetchApp

**Date:** 2026-10-25  
**Status:** Accepted  
**Deciders:** AI Lead, Solution Architect  

## Context
Antigravity requires AI capabilities to deconstruct unstructured job descriptions into pain points and to generate 3-page Google Docs interview dossiers when a positive reply is detected. The integration must be deterministic, prevent formatting breakdowns, maintain low latency, and keep monthly API expenses under $3.

## Decision
We implement a direct REST bridge to Gemini 1.5 Pro via `UrlFetchApp` in `09_AiCopilotAndDossier.js`:
1. Use `temperature: 0.2` to ensure low variability and high factual precision.
2. Enforce strict JSON output schemas using `response_mime_type: "application/json"`.
3. Store the API key securely in `PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')` — never in spreadsheet cells or source code.
4. Provide graceful deterministic fallbacks if the API key is not yet configured.

## Consequences
### Positive
- Strict JSON structure eliminates parsing failures and malformed outputs.
- Keeps monthly AI API costs well under $3.00 for active candidate volumes.
- API keys remain completely isolated from client-side code and shared sheets.

### Negative / Trade-offs
- Requires internet connectivity and an active Google AI Studio or OpenAI API key.
