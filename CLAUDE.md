# CLAUDE.md

> **Last reviewed:** April 2026

This file contains always-on rules for this repository. Detailed workflows are split into specialized skills under `.claude/skills/`.

## Purpose

This repository is a generic knowledge base and toolkit for IT carve-out project document generation.

Use reference assets for methodology only. Do not copy reference project facts into active engagement deliverables.

## Global Guardrails (Always On)

### 1) Reference Projects vs Active Engagements

- Reference folders and historical files are methodology-only.
- Never copy reference parties, scope, dates, or organization names into active engagement deliverables.
- Every generated artifact must reflect the current engagement only.

### 2) Mandatory Inputs Before Any Deliverable Generation

Confirm all fields before generating any deliverable:

- Project name
- Seller
- Buyer
- Business being carved out
- Carve-out model (Stand Alone / Integration / Combination)
- PMO / methodology lead
- Number of worldwide sites
- Number of IT users

If any are missing:

1. Block generation.
2. Ask for all missing fields at once.
3. Do not use estimates or reference-project placeholders for missing values.
4. Budget exception only: if the user explicitly says budget is unknown, use `TBC - to be approved at QG1`.

### 3) Buyer/Seller Derived Rules

- Sponsor Customer = Buyer.
- Sponsor Contractor = Seller.
- IT flow direction = Seller IT -> Merger Zone (if Integration model) -> Buyer IT.
- TSA = Seller operates services until buyer-side readiness.
- Stand Alone model = no Merger Zone.

### 4) Deliverable Orchestration

For a confirmed new engagement, create all mandatory deliverables in dependency order:

1. Schedule (CSV + XML)
2. Cost plan
3. Risk register
4. Project charter
5. Executive dashboard
6. Management KPI dashboard
7. Monthly status report

Do not generate Open Points List unless explicitly requested.

## Skill Routing

Use these skills for detailed execution:

- intake-compliance-gate
- schedule-generation
- cost-plan-generation
- risk-register-generation
- executive-dashboard-generation
- management-kpi-dashboard-generation
- monthly-status-report-generation
- repository-governance-updates

## Skill Locations

- `.claude/skills/intake-compliance-gate/SKILL.md`
- `.claude/skills/schedule-generation/SKILL.md`
- `.claude/skills/cost-plan-generation/SKILL.md`
- `.claude/skills/risk-register-generation/SKILL.md`
- `.claude/skills/executive-dashboard-generation/SKILL.md`
- `.claude/skills/management-kpi-dashboard-generation/SKILL.md`
- `.claude/skills/monthly-status-report-generation/SKILL.md`
- `.claude/skills/repository-governance-updates/SKILL.md`

## Repository Maintenance

Keep repository metadata current when:

- A new project folder is created
- A new generator script is added
- A new active project file is introduced
- A template/spec changes
- A project status moves active <-> closed

When repository metadata changes, update Last reviewed.
