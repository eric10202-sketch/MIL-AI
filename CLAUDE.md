	# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Purpose

This repository is a **knowledge base for IT carve-out project document generation**. It contains no source code. The files here are reference material from a completed real-world M&A IT carve-out (Project FRAME: Bosch → Keenfinity) plus a methodology guide. Use them to generate artifacts for new carve-out engagements.

## Files in This Repository

| File | Type | Purpose |
|------|------|---------|
| `20240718_FRAME_IT Risk Assessment_completed.csv` | Reference | Historical risk register — realistic risk categories, probability/impact ratings, mitigation patterns, and lifecycle of risks across a full carve-out project |
| `FRAME_IT_OPL_vFINAL.csv` | Reference | Historical Open Points List — 220+ items across 8 IT sub-workstreams; shows dependency patterns, decision tracking, and typical issues per workstream |
| `IT INFRA TIMELINE FRAME CURENT.pptx` | Reference | IT infrastructure milestone timeline — phase sequencing, critical dependencies (AD → apps, WAN → network cut), and realistic durations |
| `Carveout best practices guide and template.docx` | Reference | Bosch's internal IT carve-out methodology — carve-out models, guiding principles, and full workstream templates (Infrastructure, ERP/SAP, Other Apps, HR IT, IT Org, Contracts/Licenses, IT Security) |
| `generate_project_schedule.js` | Generator | Node.js script that produces `Trinity_Project_Schedule.csv` — the MS Project-importable schedule for Project Trinity |
| `Trinity_Project_Schedule.csv` | Output | Generated MS Project schedule for Project Trinity (192 tasks, 6 phases, 16 milestones) — regenerate with `node generate_project_schedule.js` |

## Document Generators

### Project Schedule — `generate_project_schedule.js`
Node.js script (no external dependencies) that writes `Trinity_Project_Schedule.csv`.

```bash
node generate_project_schedule.js
```

Columns written: `ID | Outline Level | Name | Duration | Start | Finish | Predecessors | Resource Names | Notes | Milestone`

- **Outline Level** drives the WBS hierarchy in MS Project (1 = phase, 2 = sub-group, 3 = task/milestone)
- **Milestone** = `Yes` for 0-day milestone rows (QGs, Day 1, Signing, UAT sign-off, etc.)
- Tasks defined via a compact `t()` helper — easy to add, remove, or resequence
- To create a schedule for a new project: copy the script, update the `tasks` array and the header comment block

### Output Formats for New Projects

When the user provides parameters for a new carve-out project, generate:

- **MS Project schedule** — run `generate_project_schedule.js` or produce equivalent CSV with the columns above
- **Risk register** — CSV matching the FRAME structure: Nb, Sub-project, Type, Priority, Risk category, Risk/Opportunity description, Effects, Root Cause, Probability (1–5), Impact (1–5), Risk Rating (P×I), EMV, Response strategy, Actions, Responsible, Deadline, Status
- **Open Points List** — CSV matching FRAME OPL structure: #, Source, Date Reported, Location, Sub-Workstream, WP, Category, Title, Action/Description/Impact, Execution Owner, Priority, Status, Due Date, Comments

## IT Sub-Workstream Structure (from FRAME — use as default template)

1. IT Infrastructure (WAN/LAN, AD, servers, M365/Azure, telephony)
2. Commercial IT incl. ERP (SAP migration, CRM, FSM, BPO)
3. Other Applications (~500 apps; SharePoint, Confluence/Jira, etc.)
4. Engineering IT (PLM/Windchill, FOSS compliance, developer network)
5. Production IT (OT Security, MES, DOT, plant telephony)
6. HR IT
7. IT Organization & Processes (TOM, ITO contracting, ITSM/ServiceNow)
8. IT Contracts & Licenses (change of control, SAM, FOSS)
9. IT Security (IAM/Saviynt, CISO, ISO 27001, BCM, GDPR)

## Key Carve-Out Concepts to Apply

- **"First make it work, then make it better"** — pragmatism over optimization; Day-1 readiness is the goal
- **Carve-out models**: Stand Alone (full independence), Integration with Buyer (buyer leads), or Combination
- **TSA (Transitional Service Agreement)**: Parent provides temporary services post-closing; minimize scope and duration
- **Big Bang vs. Staggered go-live**: FRAME used Big Bang due to tightly coupled logistics/ERP processes
- **Point of No Return**: An implemented production change that cannot be reversed without significant cost/compliance impact
- **Risk rating** = Probability × Impact (1–5 scale); categories: ScR=Schedule, SR=Scope, RR=Resource, BtR=Budget, QR=Quality, BR=Business, LR=Legality, CR=Customer Satisfaction
- **Country-specific complexity**: Brazil ERP (extreme tax law complexity), China (local FTS/customs), Mexico (legal entity delays), India (local IT solutions like RBIN)

## Typical Timeline Benchmarks (from FRAME)

- Total project duration: ~18–24 months from kick-off to closing
- WAN ordering lead time: 4–6 months minimum
- Active Directory build and go-live: ~6 months
- SAP migration (shell copy + testing): 9–12 months
- M365 tenant build + cutover: ~6 months
- Phone number migration: 4–8 weeks per operator (state-owned operators may take longer)
- IAM/IdM implementation: 6–9 months; Saviynt SC2 connector in Bosch-governed env restricted to 10 days before closing
- Infrastructure hub setup (co-locator): decision by month 4, implementation by month 9
- "Frozen Zone" begins at Signing — minimize changes to production environment
