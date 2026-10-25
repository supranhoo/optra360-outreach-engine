# ADR-002: Google Sheets as a 14-Table Relational CRM Database with In-Memory Hash Joins

**Date:** 2026-10-25  
**Status:** Accepted  
**Deciders:** CRM Architect, Apps Script Developer  

## Context
Google Sheets is traditionally a flat tabular grid without built-in relational integrity constraints or SQL query engines. Making repetitive cell-by-cell `getValue()` and `setValue()` calls against multiple sheets incurs significant latency, causing Apps Script executions to time out quickly.

## Decision
We enforce a normalized 14-table relational database architecture within Google Sheets:
1. Every sheet uses row 1 as an immutable frozen header row with snake_case column identifiers.
2. Column A serves as the Primary Key (e.g., `CAND-01`, `PRJ-001`, `OUT-9001`).
3. Foreign keys are strictly maintained and validated via Google Sheets Data Validation rules.
4. All database queries in Apps Script execute as single batched `getDataRange().getValues()` calls. Joins are executed in memory using $O(1)$ JavaScript lookup hash maps (`buildLookupMap_`).
5. Updates are written back to Google Sheets in single atomic `setValues()` operations.

## Consequences
### Positive
- Read/write latency is reduced by up to 90% compared to cell-by-cell operations.
- Data remains completely human-auditable and editable in Google Sheets with familiar spreadsheet tools.
- Preserves relational integrity between candidates, projects, skills, companies, opportunities, contacts, outreach logs, and responses.

### Negative / Trade-offs
- Concurrency conflicts: Multiple simultaneous script triggers writing to the same sheet could cause overwrite hazards. Mitigated by `LockService.getScriptLock(10000)` and atomic row locking (`SENDING_LOCKED`).
