# Carveout AI Toolkit — Copilot Instructions

> **Last reviewed:** April 2026 — skill coverage completed for charter, stakeholder presentation, and reconciliation gate

This file is automatically loaded by GitHub Copilot for all users of this workspace.
`CLAUDE.md` is the canonical repository policy for guardrails, mandatory inputs, orchestration, and output standards.
Detailed skill workflows are in `.agents/skills/` (relative to this workspace root).

## Copilot Execution Rules

Use `CLAUDE.md` as the single source of truth for repository policy.

Before generating or modifying any deliverable:
- Read the relevant skill file from `.agents/skills/` first.
- Follow the mandatory deliverable sequence defined in `CLAUDE.md`.
- Treat the schedule, risk register, and cost plan as the authoritative baseline documents.
- Reconcile every downstream artifact to those baseline documents before handoff.

---

## Skill Routing

When executing a task, read the relevant skill file from this workspace before proceeding:

| Task | Skill file |
|------|----------|
| Validate intake / check missing fields | `.agents/skills/intake-compliance-gate/SKILL.md` |
| Create or adapt project schedule | `.agents/skills/schedule-generation/SKILL.md` |
| Derive cost plan from schedule | `.agents/skills/cost-plan-generation/SKILL.md` |
| Generate risk register | `.agents/skills/risk-register-generation/SKILL.md` |
| Generate project charter | `.agents/skills/project-charter-generation/SKILL.md` |
| Create executive dashboard | `.agents/skills/executive-dashboard-generation/SKILL.md` |
| Create management KPI dashboard | `.agents/skills/management-kpi-dashboard-generation/SKILL.md` |
| Generate monthly status report | `.agents/skills/monthly-status-report-generation/SKILL.md` |
| Generate PowerPoint stakeholder presentation | `.agents/skills/stakeholder-presentation-generation/SKILL.md` |
| **Perform multi-document reconciliation & quality gate** | **`.agents/skills/multi-document-reconciliation-gate/SKILL.md`** |
| Update repository metadata | `.agents/skills/repository-governance-updates/SKILL.md` |

> All paths above are relative to the workspace root. Use the `read_file` tool to load the skill content before generating the deliverable.

---

## Repository Maintenance

When repository metadata changes, update `CLAUDE.md`, this file, and any affected inventories or folder-structure references.
