---
name: cost-plan-generation
description: "Use when deriving labor-only carve-out cost plans from schedule resource assignments, including category subtotals, totals, phase and category breakdowns, and CAPEX exclusions. Output is XLSX."
---

# Cost Plan Generation

## Purpose
Create a standardized draft IT labor cost plan from schedule assignments, cross-checked against the risk register.
Output is a **formatted Excel workbook (XLSX)**. CSV is no longer the primary deliverable format.

## Canonical Sources
- `{ProjectName}/{ProjectName}_Project_Schedule.xlsx` — **required input** (task/resource data)
- `{ProjectName}/{ProjectName}_Risk_Register.xlsx` — **required input**
- Existing project cost plan generators (e.g. `generate_bravo_cost_plan.py`) — **XLSX format/structure reference only**; never copy their resource names, cost figures, categories, or project-specific content

## Generation Sequence
The cost plan is the **3rd deliverable** in the mandatory generation sequence:

| Step | Deliverable | Dependency |
|---|---|---|
| 1 | Project Schedule (CSV + XML) | None |
| 2 | Risk Register (.xlsx) | Schedule |
| **3** | **Cost Plan** | **Schedule + Risk Register** |
| 4 | Project Charter | Cost Plan |
| 5 | Executive Dashboard | Cost Plan + Risk Register |
| 6 | Management KPI Dashboard | Cost Plan + Risk Register |
| 7 | Monthly Status Report | All above |

**Block cost plan generation if either schedule or risk register is missing.**

## Pre-Generation Cross-Check (Mandatory)
Before writing any cost figures, perform all checks below. Fix any gap before proceeding.

### 1 — Schedule alignment
- Phase names and date ranges in cost plan **must match** phase names and dates in `{ProjectName}_Project_Schedule.xlsx`.
- Phase breakdown rows must cover every phase present in the schedule — no phase missing, none added.
- Resource names in cost plan category blocks must map to resource names used in the schedule `Resource Names` column. If a schedule resource has no matching cost line, add one. If a cost line has no schedule counterpart, remove or rename it.

### 2 — Risk register alignment
- Review every risk in `{ProjectName}_Risk_Register.xlsx` with Status = **Amber or Red** or Risk Rating ≥ 12.
- For any such risk whose mitigation involves **external cost** (contractor, consultant, legal counsel, MSP, licensing top-up, tooling), add a corresponding CAPEX / contingency line to the cost plan with:
  - Description referencing the risk (e.g. `Risk Register #5`)
  - Estimated range (or `TBC` with confirmation QG noted)
- Cross-reference the risk register in the cost plan header line.

### 3 — Resource name consistency
- Schedule `Resource Names` use `+` to separate multiple resources per task.
- Each `+`-split token that appears in the schedule must be traceable to at least one cost plan line.
- Do not invent resource labels that are not present in the schedule (e.g. `Integration Team` if the schedule says `App Teams + Test Team`).

## Required Columns (data schema)
`CATEGORY, RESOURCE, TOTAL DAYS, TOTAL HRS, HOURLY RATE (EUR), TOTAL COST (EUR)`

## Mandatory Worksheet Sections (in order)
1. Header block — include `Based on: {ProjectName}_Project_Schedule.xlsx` **and** `Risk-aligned: {ProjectName}_Risk_Register.xlsx`, grey fill
4. **Cost category blocks** — one section per category; mid blue category header, alternating white/light-blue data rows, light blue SUBTOTAL row
5. **Overall Project Total** — dark Bosch blue fill, white bold, number-formatted EUR
6. **Cost breakdown by category** — mid blue section header, alternating rows
7. **Cost breakdown by phase** — phase names and dates must match schedule exactly
8. **CAPEX / additional costs** — excluded from labour total; must include risk-driven contingency lines
9. **Notes** — grey fill, italic; cross-reference risk register for each contingency line

## XLSX Formatting Standards
Follow the Bosch blue theme used in `generate_bravo_cost_plan.py`:

| Element | Fill | Font |
|---|---|---|
| Title banner | `#003B6E` | White bold 13pt |
| Column headers | `#0066CC` | White bold 9pt |
| Metadata rows | `#F2F2F2` | Black 8pt |
| Category section header | `#0066CC` | White bold 9pt |
| Detail rows (alternating) | White / `#EFF4FB` | Black 9pt |
| Subtotal rows | `#C6D4E8` | Black bold 9pt |
| Overall total | `#003B6E` | White bold 10pt |
| Notes rows | `#F2F2F2` | Black italic 8pt |

- Number format `#,##0` on all EUR cost and rate cells.
- Column widths: A=52, B=32, C=10, D=10, E=16, F=16.
- Freeze pane below column header row.

## Generator Script Pattern
Create `generate_{ProjectName}_cost_plan.py` as a new script from scratch:
- Derive CATEGORIES, PHASE_BREAKDOWN, CAPEX_ROWS, NOTES from the current project's schedule and risk register — **never copy these data structures from another project's generator**.
- Use existing generators only as format/structure reference for the `write_xlsx(path)` function and Bosch blue theme styling.
- Output path: `{ProjectName}/{ProjectName}_Cost_Plan.xlsx`.

## Rules
- Labour only totals in core plan.
- Hardware, licences, WAN, co-lo, travel excluded from labour total.
- Stand Alone and Integration model notes must differ.
- Budget baseline state must be explicit (`TBC - to be approved at QG0` or QG1 if unknown).
- Do not recycle resource names or CAPEX lines from reference projects (AlphaX, Trinity).

## Output Completeness
Cost plan output is complete when this file exists:
- `{ProjectName}/{ProjectName}_Cost_Plan.xlsx` ← **primary deliverable** (XLSX, not CSV)
