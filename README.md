# Optra360 & Zoptro Automated Outreach & Opportunity Engine

[![Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=flat-square&logo=google&logoColor=white)](https://script.google.com)
[![clasp](https://img.shields.io/badge/built%20with-clasp-24292e?style=flat-square&logo=github)](https://github.com/google/clasp)
[![Brand](https://img.shields.io/badge/Brand-Optra360%20%7C%20Prove%20Before%20You%20Promise-09293D?style=flat-square)](#brand-expression-system-governance)

An enterprise-grade, dual-mode outreach automation engine built for **Optra360** and **Zoptro**. It automates market research, decision-maker discovery, DNS MX email verification, Gemini-powered operational diagnosis, 2-tier personalized pitch deck generation via Google Slides, draft-first Gmail delivery, and LinkedIn connection sequencing.

---

## Architecture & Relational Data Model

### 1. 7-Tab CRM Command Center (`Google Sheets`)
* `1_Profiles_Config`: Dual-mode profiles (Mode B for Ankit Choudhary; Mode A for co-founders Jaspal, Divyanee, Sunehri, Vivek).
* `2_Target_Companies`: Sourced enterprises (Manufacturing, Steel, Cement, Power, Pharma, Mining, Auto) with buying trigger tagging.
* `3_Problem_Solution`: Problem diagnosis mapped to the Optra360 6-Stage Delivery Framework and engagement tiers.
* `4_Contacts_Outreach`: Decision makers, MX-verified emails, customized PDF deck links, and LinkedIn notes.
* `5_Deck_Templates`: Master Google Slides template IDs and `{{Merge_Tags}}`.
* `6_Demos_&_Portfolio`: Live portfolio links (`DEMO-01` to `DEMO-06`) including Zoptro HRMS (~2,500 employee scale).
* `7_KPI_Dashboard`: Live formula-driven metrics tracking conversion rates across each stage.

### 2. Apps Script Backend Modules
* `00_Config.js`: Central configuration, brand colors, fonts, document IDs, and API property keys.
* `01_Profiles_CRM.js`: Relational CRM reader/writer and dual-mode profile manager.
* `02_ResearchEngine.js`: Daily research runner with Google Custom Search API and zero-cost Google News RSS parser.
* `03_ContactEnrichment.js`: Email permutation engine with live Google Public DNS MX verification (`dns.google`).
* `04_SolutionMapper.js`: Gemini AI integration mapping operational bottlenecks to the 6-stage transformation delivery model.
* `05_PitchDeckBuilder.js`: Dynamic Google Slides cloner, merge-tag replacer, and 2-tier PDF generator (5-Slide Teaser + 12-Slide Detailed).
* `06_OutreachDispatcher.js`: Branded HTML email composer with email signature and Gmail Drafts bundler.
* `07_DemoPortfolioHub.js`: Configurable demo repository for Zoptro HRMS, PMS, and Industrial Engineering tools.
* `08_SidebarUI.html`: Optra360 Command Sidebar UI with Outfit/Figtree typography and sharp 90° corners.
* `09_SidebarController.js`: Bridge between sidebar UI and Apps Script server functions.
* `10_MenuTriggers.js`: Custom spreadsheet menu, initial sheet bootstrapper, and 08:00 AM IST scheduled triggers.

---

## Brand Expression System Governance

The engine strictly enforces Optra360’s brand identity rules:
* **Brand Philosophy**: *"Prove before you promise"*
* **Typography**:
  * Headlines & KPIs: **Outfit** (Bold / SemiBold / Medium)
  * Body & Tables: **Figtree** (Regular / Medium / SemiBold)
* **Color Palette**:
  * Authority Blue: `#09293D`
  * Clarity Ivory: `#F4F2EA`
  * Grounded Olive: `#666D57`
  * Progress Orange: `#F65800`
* **Edge Geometry**: **Strict 0px border-radius** (sharp 90° corners on all cards, buttons, tables, and PDF elements).

---

## CLI Deployment & Updating in Google

This repository is pre-configured with Google's official command-line tool `@google/clasp`.

### Pushing Changes to Google Apps Script
```bash
# Push all code to Google Apps Script
node "C:\Users\ANKIT CHOUDHARY\AppData\Roaming\npm\node_modules\@google\clasp\build\src\index.js" push

# Or run status check
node "C:\Users\ANKIT CHOUDHARY\AppData\Roaming\npm\node_modules\@google\clasp\build\src\index.js" status
```

### Git & GitHub Workflow
```bash
git add .
git commit -m "feat: complete optra360 outreach engine architecture"
git push -u origin main
```

---

## Quick Start in Google Sheets

1. Open your master spreadsheet: [`Optra360 & Zoptro CRM Engine`](https://docs.google.com/spreadsheets/d/1KEJfOBK8PobuGQ6TI-91AKebXehOphDD3LwQw_zaXwU/edit)
2. Reload the spreadsheet to load the custom menu: **Optra360 Engine**
3. Click **Optra360 Engine &rarr; 🛠 Seed 7-Tab CRM Architecture & Demos** (bootstraps tabs if blank).
4. Click **Optra360 Engine &rarr; 🚀 Open Command Sidebar** to launch the control center.
5. Click **Optra360 Engine &rarr; ⏰ Set Up Daily 08:00 AM IST Trigger** to automate daily sourcing and drafting.
