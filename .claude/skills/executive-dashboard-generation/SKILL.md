---
name: executive-dashboard-generation
description: "Use when creating executive carve-out dashboards in self-contained HTML with Bosch palette, timeline and milestones, workstream confidence, regional scope, risks, and critical path sections."
---

# Executive Dashboard Generation

## Purpose
Produce a strategic executive dashboard from schedule, cost, and risk data.

## Canonical Format — Project Trinity
The reference layout is **Project Trinity — Executive Dashboard.pdf** (A4, 3-page portrait).
All new dashboards must replicate this layout.

### Canonical Sources (in priority order)
1. `Project Trinity — Executive Dashboard.pdf` — visual/layout reference
2. `AlphaX/AlphaX_Executive_Dashboard.html` — working HTML implementation

## Non-Negotiables
- Self-contained HTML (no external CDN/fonts/URLs).
- Bosch palette and system font stack only.
- Bosch logo: embed `Bosch.png` as `data:image/png;base64,...` in an `<img>` tag.
  - File: `Bosch.png` (workspace root)
  - Size the `<img>` at `height:36px` inside the header.
- Print-safe page-break rules (`page-break-before:always` on `.page-break`).
- SharePoint offline rendering support.

## Bosch Logo Embedding
Read `Bosch.png` from workspace root and base64-encode it, then use:
```html
<img src="data:image/png;base64,<BASE64>" alt="Bosch — Invented for Life" style="height:36px;display:block;" />
```
Wrap it in a white background container (e.g. `background:#fff; padding:4px 8px; border-radius:4px;`).

## Trinity 3-Page Layout Specification

### Page 1
1. **Header band** (dark navy gradient): Bosch logo (top-left) | Programme title + buyer/seller subtitle | date/countdown (top-right)
2. **Days-to-key-events strip** (Bosch red): countdown boxes for Kickoff, Day-1 Closing, TSA Exit
3. **PROJECT OVERVIEW section**: 2-column — left: narrative paragraph; right: Carve-Out Model box + Key Parties + Programme Budget + Governance
4. **Stats row** (6 icons): Global Sites · Employees · Client Devices · Applications · Project Duration · TSA Duration
5. **Phase Timeline bar**: horizontal colour-coded segments with date labels below
6. **Two-column lower half**:
   - Left: KEY MILESTONES & QUALITY GATES table (icon | name+desc | date | days-from-today | status pill)
   - Right: BUDGET DISTRIBUTION (total figure + labour-only caveat + donut/bar breakdown by resource category)

### Page 2
7. **Continued milestones** (QG4 through QG6 if overflow from Page 1)
8. **IT WORKSTREAM COVERAGE**: 3×3 grid (WS1–WS9) with title, bullet detail, confidence tag
9. **QUALITY GATE TRACKER**: list — date | gate name+days-from-today | criteria paragraph
10. **SCOPE & SCALE INDICATORS** (2-column):
    - Left: REGIONAL SITE DISTRIBUTION — bar per region with site count
    - Right: KEY RISK INDICATORS — HIGH/MEDIUM/LOW cards with P·I·Rating scores

### Page 3
11. **APPLICATION MIGRATION WAVES**: wave rows with bar chart widths and app counts
12. **COUNTRY-SPECIFIC COMPLEXITY HOTSPOTS**: 2×2 or 2×3 grid per country flag
13. **Stats strip** (dark navy background): Total Tasks · Resource Groups · Person-Hours · Regional DC Hubs
14. **CRITICAL PATH & GUIDING PRINCIPLES**: 4-column grid — Infrastructure CP | ERP CP | Client Workplace | Programme Principles
15. **Footer**: project name | dashboard date | data source CSV files | confidentiality notice

## Required Content (summary)
- Programme overview and countdown
- Timeline and milestones
- Budget summary with resource-category breakdown
- Workstream confidence (9 workstreams)
- QG tracker (all gates with criteria)
- Regional scope and site distribution
- Risk indicators (HIGH/MEDIUM/LOW with P×I ratings)
- App waves and country hotspots
- Critical path and principles
