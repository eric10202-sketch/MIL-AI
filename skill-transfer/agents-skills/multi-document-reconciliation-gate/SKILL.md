---
name: multi-document-reconciliation-gate
description: "Use when performing the final cross-document quality gate to verify schedule, risk, cost, charter, dashboards, monthly report, and stakeholder presentation are mutually consistent before handoff."
---

# Multi-Document Reconciliation Gate

## Purpose
Provide the final quality gate before stakeholder handoff by checking that all generated deliverables for the current project tell the same story.

## Required Inputs
The following deliverables must already exist for the same project:
- Schedule (`XLSX` and `XML`)
- Risk register (`XLSX`)
- Cost plan (`XLSX`)
- Project charter (`HTML`)
- Executive dashboard (`HTML`)
- Management KPI dashboard (`HTML`)
- Monthly status report (`PDF`)
- Stakeholder presentation (`PPTX`)

## Gate Objective
Detect drift between deliverables before handoff. If any material mismatch is found, block completion until corrected.

## Reconciliation Checks

### 1. Identity and Intake Consistency
- Project name matches across all deliverables.
- Seller and buyer match intake.
- Business being carved out matches intake.
- Carve-out model matches intake.
- PMO / methodology lead is consistent where referenced.

### 2. Date Consistency
- Project start date matches schedule and downstream documents.
- QG1, QG2&3, QG4, GoLive, and QG5 dates match the generated schedule.
- Monthly report timing logic and countdown messaging are aligned to the reporting date.

### 3. Risk Consistency
- Risk count and top-risk prioritization match the generated risk register.
- Any risks referenced in dashboards, report, charter, or deck exist in the register.
- Risk-driven CAPEX or contingency references map to real risk IDs.

### 4. Cost Consistency
- Labour totals and phase totals match the generated cost plan.
- Budget notes such as `TBC - to be approved at QG1` are consistent everywhere.
- Cost references in charter, dashboards, report, and deck align to the same baseline.

### 5. Model and Scope Consistency
- Stand Alone projects do not mention a merger zone.
- Integration projects consistently describe seller IT to merger zone to buyer IT flow.
- Sites, users, and application scope are consistent across narrative artifacts.

### 6. Handoff Readiness
- No deliverable contains copied legacy project names, parties, or dates.
- Output filenames follow repository conventions.
- Deliverables open successfully in their native formats where feasible.

## Outcome Rules
- If all checks pass, the project is ready for stakeholder handoff.
- If any check fails, document the mismatch clearly and fix the source artifact rather than masking it downstream.

## Recommended Execution Pattern
1. Use the schedule, risk register, and cost plan as the factual baseline.
2. Compare each downstream deliverable against that baseline.
3. Record any mismatches.
4. Correct the affected deliverable.
5. Re-run the reconciliation gate until clean.

## Completion Checklist
- [ ] All eight deliverables exist.
- [ ] Identity fields reconcile across artifacts.
- [ ] Key milestone dates reconcile across artifacts.
- [ ] Risk references reconcile across artifacts.
- [ ] Budget and baseline cost references reconcile across artifacts.
- [ ] No copied reference-project facts remain.
- [ ] Stakeholder handoff is unblocked.