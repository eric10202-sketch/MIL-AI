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
4. Confirm Bosch-Logo.png path and output folder.
5. Run script to produce {ProjectName}_Monthly_Status_Report_{MMM_YYYY}.pdf.

## Rules
- No interactive prompts.
- Filename month/year must auto-refresh from runtime date.
- Maintain single-page A4 report layout.
- Keep Bosch Digital color mapping.
