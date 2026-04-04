# CLAUDE.md

> **Last reviewed:** April 2026 — risk register template lookup and matrix formatting rules updated

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
- Project start date
- Project GoLive date
- Project completion date

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

For a confirmed new engagement, create all mandatory deliverables in this exact dependency order:

1. Schedule (XLSX + XML)
2. Risk register (Excel template format: XLS/XLSX)
3. Cost plan
4. Project charter
5. Executive dashboard
6. Management KPI dashboard
7. Monthly status report

Do not generate Open Points List unless explicitly requested.

**Each deliverable depends on all preceding ones being complete and consistent.**
The cost plan must be cross-checked against both the schedule and the risk register before generation (see cost-plan-generation skill).

### 5) Output Format Standards

- Schedule must always be delivered in two files: **XLSX** (primary human-readable) and MS Project XML.
  - XLSX: formatted Bosch blue theme workbook (see `schedule-generation` skill for formatting rules).
  - XML must always be generated via `generate_msp_xml.py` — never hand-written. The script uses a temporary CSV internally; no CSV file is kept in the output folder.
  - The XML generator enforces critical MSPDI rules (TaskMode ordering, ManualStart/Finish, ConstraintType) that are required for MS Project to honour task dates. See `schedule-generation` skill for full rules.
- Cost plan must always be delivered as a **formatted XLSX workbook** (Bosch blue theme). CSV is no longer the primary cost plan format. See `cost-plan-generation` skill for formatting rules.
- Risk register must always be delivered in Excel template format (XLS/XLSX) aligned with `BD_Risk-Register_template_en_V1.0_Dec2023.xlsx` structure.
  - Use the template-owned lookup blocks as-is: categories `D140:D156`, sources `D171:D175`, impact `D182:E186`, probability `D189:E193`, traffic-light thresholds `D196:E199`, strategy `D202:D208`, status `D213:D217`.
  - Impact labels must match the template exactly: `Very Low`, `Low`, `Moderate`, `High`, `Very High`.
  - For generated lookup formulas in columns `M`, `O`, `AC`, and `AE`, use the template-compatible `=_xlfn.IFNA(VLOOKUP(...),"")` pattern.
  - If a workbook is re-saved through `openpyxl`, explicitly force black font on yellow-filled cells in the `Matrix ` sheet so Excel does not render unreadable light-on-yellow text.
  - Prefer plain ASCII punctuation in generated workbook text to avoid special-character corruption in Excel-bound outputs.
- Project charter, executive dashboard, and management KPI dashboard must include an embedded Bosch logo in the document header/cover.
  - Default logo file: `Bosch.png` (workspace root) — embed as `data:image/png;base64,...` at `height:36px`.
  - `.bosch-logo` container CSS: `display:flex; align-items:center;` (do **not** use fixed `width`/`height` or `display:grid` — those clip the image).

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
