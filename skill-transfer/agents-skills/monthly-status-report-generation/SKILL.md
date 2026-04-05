---
name: monthly-status-report-generation
description: "Use when generating monthly carve-out PDF status reports with auto-calculated days-to-gate, phase/risk/budget sections, and scheduler-ready output naming."
---

# Monthly Status Report Generation

## Purpose
Create scheduled executive PDF status reports with current-date calculations.

## Canonical Source
- Existing project report generators (e.g. `generate_alphax_monthly_report.py`) — **format/layout reference only**; never copy their project-specific content, dates, parties, status text, or risk summaries

## Steps
1. Create a new `generate_{ProjectName}_monthly_report.py` script from scratch.
2. Define the project configuration block using the current project's actual parameters (name, seller, buyer, dates, sites, users, model).
3. Write all section content (overview, phase status, risk summary, budget status, next steps) fresh based on the current project's scope and timeline.
4. Confirm Bosch.png path and output folder.
5. Run script to produce `{ProjectName}_Monthly_Status_Report_{MMM_YYYY}.pdf`.

## Content Generation Rules
- **Never copy status text, risk summaries, budget figures, or narrative from any existing project** (AlphaX, Bravo, Falcon, Hamburger, or any other).
- All dates, days-to-gate calculations, and phase progress must reflect the current project's timeline.
- Risk highlights must come from the current project's risk register, not borrowed from another project.
- Budget status must reference the current project's cost plan.

## Rules
- No interactive prompts.
- Filename month/year must auto-refresh from runtime date.
- Maintain single-page A4 report layout.
- Keep Bosch Digital color mapping.

## Colour Theme — Blue as Primary
Use **Blue** as the primary/hero colour for all new PDF outputs. Do NOT use Bosch Red (`#E20015`) as the dominant header colour.

| Role | Recommended value |
|---|---|
| Header / hero band | `#003b6e` (deep navy blue) |
| Accent / section bar | `#0066CC` |
| Status: GREEN | `#007A33` |
| Status: AMBER | `#E8A000` |
| Status: RED | `#CC0000` |
| Body background | `#f4f6f9` |
| Body text | `#1a1a1a` |

Bosch Red may still appear for red-status RAG badges, but must **not** be the primary brand colour of the report.
