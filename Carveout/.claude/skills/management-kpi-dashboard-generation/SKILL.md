---
name: management-kpi-dashboard-generation
description: "Use when creating management KPI dashboards with 12-column card layout, SPI/CPI/readiness metrics, milestone controls, top risks, model differences, and 90-day action forecast."
---

# Management KPI Dashboard Generation

## Purpose
Create the operational steering dashboard for PMO and SteerCo.

## Canonical Source
- AlphaX/AlphaX_Management_KPI_Dashboard.html

## Non-Negotiables
- Self-contained HTML and Bosch visual standards.
- 12-column CSS grid card layout.
- SharePoint offline compatibility.
- Embedded Bosch logo in the dashboard header.

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
