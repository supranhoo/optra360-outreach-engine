# ADR-008: Standalone Apps Script Web App Architecture with Dynamic Database Provisioning

**Date:** 2026-10-25  
**Status:** Accepted  
**Deciders:** Solution Architect, Front-End Engineer  

## Context
Initial implementations relied on container-bound Google Apps Script embedded inside an existing Google Sheet, requiring the user to operate primarily within Google Sheets rows and custom menus. The user requested that the platform run as an independent, standalone application in Google Apps Script with its own dedicated Web App interface, without being tied to an existing spreadsheet.

## Decision
We decouple the backend compute and user interface from any single Google Sheet container:
1. Re-initialize the project as an independent standalone Google Apps Script (`1b1oJEtAVGcB1FeVFJ7V7tWV-Lo-beGkoQgreIryArcDmD9rSfV6Powqf`).
2. Implement a complete Single Page Application (SPA) in `WebAppUI.html` served natively via `doGet(e)` and `HtmlService`.
3. Provide dynamic database provisioning in `00_Config.js`:
   - Checks `PropertiesService` for `SPREADSHEET_ID`.
   - If not found, automatically creates a dedicated Google Sheet database (`⚡ Project Antigravity CRM Database`) via `SpreadsheetApp.create()`, storing its ID in `PropertiesService`.
4. Expose all CRM operations (lead scoring, campaign draft generation, batch dispatching, reply polling, AI key settings) through client-server RPC functions (`google.script.run`) in `WebAppController.js`.

## Consequences
### Positive
- Fully standalone web application accessible via direct browser URL.
- Zero dependency on a pre-existing sheet; database is provisioned and managed automatically.
- Rich, modern SaaS user experience with live progress bars, draft preview drawer, and interactive controls.

### Negative / Trade-offs
- Requires Web App deployment configuration in `appsscript.json` (`executeAs: "USER_ACCESSING"`, `access: "MYSELF"`).
