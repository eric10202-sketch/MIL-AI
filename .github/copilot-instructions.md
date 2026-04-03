# Carveout AI Toolkit — Copilot Instructions

> **Last reviewed:** April 2026

This file is automatically loaded by GitHub Copilot for all users of this workspace.
Detailed skill workflows are in `.claude/skills/` (relative to this workspace root).

## Purpose

This workspace is a generic knowledge base and toolkit for IT carve-out project document generation.

Use reference assets (in `references/`, `archive/`) for methodology only. Do not copy reference project facts into active engagement deliverables.

---

## Global Guardrails (Always On)

### 1) Reference Projects vs Active Engagements

- Reference folders and historical files are methodology-only.
- Never copy reference parties, scope, dates, or organization names into active engagement deliverables.
- Every generated artifact must reflect the current engagement only.

### 2) Mandatory Inputs Before Any Deliverable Generation

Confirm ALL of the following fields before generating any deliverable:

- Project name
- Seller
- Buyer
- Business being carved out
- Carve-out model (Stand Alone / Integration / Combination)
- PMO / methodology lead
- Number of worldwide sites
- Number of IT users
- Project start date
- Project GoLive date
- Project completion date

If any are missing:

1. **Block generation.**
2. Ask for all missing fields at once.
3. Do not use estimates or reference-project placeholders for missing values.
4. Budget exception only: if the user explicitly says budget is unknown, use `TBC - to be approved at QG1`.

### 3) Buyer/Seller Derived Rules

- Sponsor Customer = Buyer.
- Sponsor Contractor = Seller.
- IT flow direction = Seller IT → Merger Zone (if Integration model) → Buyer IT.
- TSA = Seller operates services until buyer-side readiness.
- Stand Alone model = no Merger Zone.

### 4) Deliverable Orchestration

For a confirmed new engagement, create all mandatory deliverables in this exact dependency order:

1. Schedule (CSV + XML)
2. Risk register (Excel template format: XLS/XLSX)
3. Cost plan
4. Project charter
5. Executive dashboard
6. Management KPI dashboard
7. Monthly status report

Do not generate Open Points List unless explicitly requested.

**Each deliverable depends on all preceding ones being complete and consistent.**
The cost plan must be cross-checked against both the schedule and the risk register before generation.

### 5) Output Format Standards

- Schedule must always be delivered in both CSV and MS Project XML formats.
  - XML must always be generated via `generate_msp_xml.py` — never hand-written.
  - The XML generator enforces critical MSPDI rules (TaskMode ordering, ManualStart/Finish, ConstraintType) that are required for MS Project to honour task dates. See `schedule-generation` skill for full rules.
- Risk register must always be delivered in Excel template format (XLS/XLSX) aligned with `Risk_analysis_template.xlsx` structure.
- Project charter, executive dashboard, and management KPI dashboard must include an embedded Bosch logo in the document header/cover.
  - Default logo file: `Bosch.png` (workspace root) — embed as `data:image/png;base64,...` at `height:36px`.
  - `.bosch-logo` container CSS: `display:flex; align-items:center;` (do **not** use fixed `width`/`height` or `display:grid` — those clip the image).

---

## Skill Routing

When executing a task, read the relevant skill file from this workspace before proceeding:

| Task | Skill file |
|------|-----------|
| Validate intake / check missing fields | `.claude/skills/intake-compliance-gate/SKILL.md` |
| Create or adapt project schedule | `.claude/skills/schedule-generation/SKILL.md` |
| Derive cost plan from schedule | `.claude/skills/cost-plan-generation/SKILL.md` |
| Generate risk register | `.claude/skills/risk-register-generation/SKILL.md` |
| Create executive dashboard | `.claude/skills/executive-dashboard-generation/SKILL.md` |
| Create management KPI dashboard | `.claude/skills/management-kpi-dashboard-generation/SKILL.md` |
| Generate monthly status report | `.claude/skills/monthly-status-report-generation/SKILL.md` |
| Update repository metadata | `.claude/skills/repository-governance-updates/SKILL.md` |

> All paths above are relative to the workspace root. Use the `read_file` tool to load the skill content before generating the deliverable.

---

## Repository Maintenance

Keep repository metadata current when:

- A new project folder is created
- A new generator script is added
- A new active project file is introduced
- A template/spec changes
- A project status moves active ↔ closed

When repository metadata changes, update **Last reviewed** at the top of this file.
