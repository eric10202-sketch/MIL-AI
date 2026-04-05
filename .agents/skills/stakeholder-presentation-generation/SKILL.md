---
name: stakeholder-presentation-generation
description: "Use when generating Bosch-branded stakeholder presentation PPTX decks from the current project's schedule, risk, cost, charter, dashboard, and monthly report outputs."
---

# Stakeholder Presentation Generation

## Purpose
Create the management stakeholder presentation as the final executive communication artifact for the current carve-out engagement.

## Prerequisites
- Intake is confirmed.
- Schedule, risk register, cost plan, charter, executive dashboard, management KPI dashboard, and monthly status report are already complete and consistent.

## Canonical Sources
- Current project outputs in `active-projects/{ProjectName}/`.
- `Reference/Bosch presentation template.pptx` as the required PowerPoint base template.
- Existing presentation generators (for example `generate_bravo_stakeholder_presentation.py`, `generate_trinity_cam_stakeholder_presentation.py`) are structure references only. Never copy their project-specific slide text, decisions, risks, dates, or budget content.

## Steps
1. Create a new `generate_{ProjectName}_stakeholder_presentation.py` script from scratch.
2. Load the Bosch presentation template as the base presentation.
3. Read the current project's generated schedule, risk register, cost plan, charter, dashboards, and monthly report inputs.
4. Build the slide narrative from the current engagement only.
5. Save output to `active-projects/{ProjectName}/{ProjectName}_Stakeholder_Presentation.pptx`.

## Minimum Slide Set
- Cover slide with project name, carve-out model, and report date.
- Executive summary.
- Scope and operating model.
- Timeline and quality gates.
- Budget and cost structure.
- Top risks and management actions.
- Decisions required / steering focus.

## Content Rules
- Never copy slide bullets from another project.
- Facts must reconcile to the current project's generated upstream artifacts.
- Budget statements must distinguish labour baseline from CAPEX or contingency items when those are handled separately.
- Integration model decks must clearly show seller IT to merger zone to buyer IT flow.
- Use management-level language: concise, decision-oriented, and grounded in current outputs.

## Format Rules
- Output format is PPTX only.
- Use `python-pptx` and the Bosch presentation template.
- Apply Bosch branding consistently across slides, including header, footer, accent shapes, and logo placement where appropriate.

## Verification Checklist
- [ ] Output PPTX exists in the active project folder.
- [ ] Cover slide uses current project name and current report date.
- [ ] GoLive, QG1, and QG5 dates match the generated schedule.
- [ ] Budget figures reconcile to the current cost plan.
- [ ] Top risks reconcile to the current risk register.
- [ ] No copied reference-project content appears in the deck.