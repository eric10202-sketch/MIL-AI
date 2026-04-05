---
name: management-kpi-dashboard-generation
description: "Use when creating management KPI dashboards with 12-column card layout, SPI/CPI/readiness metrics, milestone controls, top risks, model differences, and 90-day action forecast."
---

# Management KPI Dashboard Generation

## Purpose
Create the operational steering dashboard for PMO and SteerCo.

## Canonical Source
- Existing project KPI dashboards (e.g. `AlphaX/AlphaX_Management_KPI_Dashboard.html`) — **HTML/CSS layout reference only**; never copy their project-specific KPI values, milestone dates, risk tables, or narrative content

## Content Generation Rules
- **Never copy KPI values, milestone data, risk tables, workstream confidence scores, or action items from another project.**
- All metrics (SPI, CPI, readiness) must be calculated from the current project's schedule, cost plan, and risk register.
- Milestone gate control timeline must use the current project's actual QG dates.
- Top risk table must come from the current project's risk register.
- 90-day action forecast must be based on the current project's upcoming milestones and deliverables.

## Non-Negotiables
- Self-contained HTML and Bosch visual standards.
- 12-column CSS grid card layout.
- SharePoint offline compatibility.
- Embedded Bosch logo in the dashboard header.

## Colour Theme — Blue as Primary
Use **Blue** as the primary/hero colour for all new HTML outputs. Do NOT use Bosch Red (`#E20015`) as the dominant background or header colour.

| Role | Recommended value |
|---|---|
| Hero / header background | `#003b6e` (deep navy blue) |
| Accent / highlight | `#0066CC` (Bosch mid-blue) |
| Section header bar | `#005199` |
| Link / interactive | `#0077BB` |
| Status: GREEN | `#007A33` |
| Status: AMBER | `#E8A000` |
| Status: RED | `#CC0000` |
| Body background | `#f4f6f9` |
| Card background | `#ffffff` |
| Body text | `#1a1a1a` |

Bosch Red may still appear for red-status RAG badges or critical-path indicators, but must **not** be the primary brand colour of the document.

## Bosch Logo Embedding
Read `Bosch.png` from workspace root and base64-encode it, then use:
```html
<img src="data:image/png;base64,<BASE64>" alt="Bosch — Invented for Life" style="height:36px;display:block;" />
```
Wrap in a white background container: `background:#fff; padding:4px 8px; border-radius:4px;`

`.bosch-logo` CSS must use `display:flex; align-items:center;` — do **not** use fixed `width`/`height` or `display:grid`.

## Required KPI Areas
- Schedule performance (SPI)
- Cost performance (CPI)
- Day 1 readiness
- Stand Alone / TSA confidence
- Workstream confidence bars
- Milestone gate control timeline
- Top risk table
- Model key differences
- Next 90 days action forecast
