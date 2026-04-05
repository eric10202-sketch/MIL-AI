---
name: project-charter-generation
description: "Use when generating carve-out project charters in HTML from confirmed intake, schedule, risk, and cost data with Bosch branding and engagement-specific narrative."
---

# Project Charter Generation

## Purpose
Create the project charter as a Bosch-branded HTML deliverable that establishes the programme mandate, scope, governance, timeline, risks, and budget framing for the current engagement.

## Prerequisites
- Intake fields are fully confirmed.
- Schedule is complete and consistent with the current project.
- Risk register is complete and aligned to the schedule.
- Cost plan is complete and cross-checked against schedule and risks.

## Canonical Sources
- Current project deliverables only:
  - `active-projects/{ProjectName}/{ProjectName}_Project_Schedule.xlsx`
  - `active-projects/{ProjectName}/{ProjectName}_Risk_Register.xlsx`
  - `active-projects/{ProjectName}/{ProjectName}_Cost_Plan.xlsx`
- Existing charter generators (for example `generate_bravo_charter.py`, `generate_charlie_charter.py`) are format and structure references only. Never copy project facts, dates, risk summaries, or narrative text.

## Steps
1. Create a new `generate_{ProjectName}_charter.py` script from scratch.
2. Read the current project's generated schedule, risk register, and cost plan rather than hardcoding copied values.
3. Build a self-contained HTML document with Bosch branding and embedded logo.
4. Write charter narrative fresh for the current engagement only.
5. Save output to `active-projects/{ProjectName}/{ProjectName}_Project_Charter.html`.

## Required Sections
- Cover or header with project name, seller, buyer, carve-out model, and report date.
- Executive summary describing the carve-out purpose and target operating path.
- Scope and objectives.
- Timeline and key gates from QG0 through QG5.
- Governance and delivery model.
- Risk summary based on the current risk register.
- Budget summary based on the current cost plan.
- Assumptions, dependencies, and success criteria.

## Content Rules
- Never copy charter prose from any other project.
- Buyer, seller, business scope, dates, and model must exactly match intake and upstream deliverables.
- Timeline, milestone dates, risk counts, and budget figures must be read from current generated artifacts.
- If budget is still unknown, use `TBC - to be approved at QG1` exactly.
- Integration model charters must describe the merger zone as a temporary transition layer, not the target-state platform.

## Formatting Rules
- Output format is HTML.
- Include the Bosch logo from `Bosch.png`.
- Use the existing Bosch blue visual language already established in repository deliverables.
- Keep the document self-contained so it can be opened directly in a browser.

## Verification Checklist
- [ ] Output file exists in the active project folder.
- [ ] Project name, seller, buyer, and carve-out model match intake.
- [ ] Milestone dates match the generated schedule.
- [ ] Risk summary matches the generated risk register.
- [ ] Budget summary matches the generated cost plan.
- [ ] No copied reference-project narrative appears in the charter.