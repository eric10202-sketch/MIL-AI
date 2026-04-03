---
name: monthly-status-report-generation
description: "Use when generating monthly carve-out PDF status reports with auto-calculated days-to-gate, phase/risk/budget sections, and scheduler-ready output naming."
---

# Monthly Status Report Generation

## Purpose
Create scheduled executive PDF status reports with current-date calculations.

## Canonical Source
- generate_alphax_monthly_report.py

## Steps
1. Copy script to generate_{ProjectName}_monthly_report.py.
2. Update project configuration block.
3. Validate date dictionary and section content.
4. Confirm Bosch.png path and output folder.
5. Run script to produce {ProjectName}_Monthly_Status_Report_{MMM_YYYY}.pdf.

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
