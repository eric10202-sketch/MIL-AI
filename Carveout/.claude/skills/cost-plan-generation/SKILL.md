---
name: cost-plan-generation
description: "Use when deriving labor-only carve-out cost plans from schedule resource assignments, including category subtotals, totals, phase and category breakdowns, and CAPEX exclusions."
---

# Cost Plan Generation

## Purpose
Create a standardized draft IT labor cost plan from schedule assignments, cross-checked against the risk register.

## Canonical Sources
- `{ProjectName}/{ProjectName}_Project_Schedule.csv` — **required input**
- `{ProjectName}/{ProjectName}_Risk_Register.xlsx` — **required input**
- `AlphaX/AlphaX_Cost_Plan.csv` — reference structure only
- `Trinity_Project_Cost_Plan.csv` — reference structure only

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
- Phase names and date ranges in cost plan **must match** phase names and dates in `{ProjectName}_Project_Schedule.csv`.
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

## Required Columns
`CATEGORY, RESOURCE, TOTAL DAYS, TOTAL HRS, HOURLY RATE (EUR), TOTAL COST (EUR)`

## Mandatory Sections
1. Header block — include `Based on: {ProjectName}_Project_Schedule.csv` **and** `Risk-aligned: {ProjectName}_Risk_Register.xlsx`
2. Cost category blocks with SUBTOTAL rows
3. OVERALL PROJECT TOTAL section
4. Cost breakdown by category
5. Cost breakdown by phase (phase names and dates must match schedule exactly)
6. CAPEX / additional costs (excluded from labor total) — must include risk-driven contingency lines
7. Notes — include cross-reference note to risk register for each contingency line

## Rules
- Labor only totals in core plan.
- Hardware, licences, WAN, co-lo, travel excluded from labor total.
- Stand Alone and Integration model notes must differ.
- Budget baseline state must be explicit (`TBC - to be approved at QG1` if unknown).
- Do not recycle resource names or CAPEX lines from reference projects (AlphaX, Trinity).
