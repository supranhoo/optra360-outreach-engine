# ADR-006: Bipartite Tag Intersection Matching Algorithm for Proof-of-Work Injection

**Date:** 2026-10-25  
**Status:** Accepted  
**Deciders:** Senior Product Manager, Systems Architect  

## Context
Candidates possess diverse technical projects, case studies, and achievements. Generic cold emails fail because they do not align the candidate's exact technical proof with the specific bottleneck described in the employer's job description. Relying solely on manual selection is unscalable, while unconstrained LLM matching is non-deterministic and can produce hallucinated project claims.

## Decision
We implement a deterministic bipartite tag intersection algorithm in `01_ContactManager.js` and `02_03_CampaignAndPersonalizer.js`:
```
Match_Score(Project_i, Opp_j) = (2.0 * |Pain_Tags_i ∩ Pain_Tags_j|) + (1.2 * |Tech_Tags_i ∩ Tech_Tags_j|) + Persona_Bonus
```
1. Extract required skills and pain tags from `Opportunities`.
2. Compute tag intersections against candidate project problem tags (weighted 2.0x) and tech stack tags (weighted 1.2x).
3. Apply a +1.0 persona affinity bonus if `Projects.target_persona_fit` directly matches the recipient persona (`FOUNDER`, `HIRING_MGR`, `TECH_LEAD`).
4. Inject the top-scoring project's architecture summary, verified impact metrics, and UTM-instrumented GitHub repository into the rendered email.

## Consequences
### Positive
- Fully deterministic, defensible matching without hallucination risks.
- Every email is anchored around genuine code and real production metrics.
- Sub-millisecond execution overhead in Apps Script.

### Negative / Trade-offs
- Requires thoughtful initial tagging of candidate projects in the `Projects` table.
