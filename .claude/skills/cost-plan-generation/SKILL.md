---
name: cost-plan-generation
description: "Use when deriving labor-only carve-out cost plans from schedule resource assignments, including category subtotals, totals, phase and category breakdowns, and CAPEX exclusions."
---

# Cost Plan Generation

## Purpose
Create a standardized draft IT labor cost plan from schedule assignments.

## Canonical Sources
- AlphaX/AlphaX_Cost_Plan.csv
- Trinity_Project_Cost_Plan.csv
- {ProjectName}_Project_Schedule.csv

## Required Columns
CATEGORY, RESOURCE, TOTAL DAYS, TOTAL HRS, HOURLY RATE (EUR), TOTAL COST (EUR)

## Mandatory Sections
1. Header block (4 lines)
2. Cost category blocks with SUBTOTAL rows
3. OVERALL PROJECT TOTAL section
4. Cost breakdown by category
5. Cost breakdown by phase
6. CAPEX / additional costs (excluded from labor total)
7. Notes

## Rules
- Labor only totals in core plan.
- Hardware, licenses, WAN, co-lo, travel excluded from labor total.
- Stand Alone and Integration model notes must differ.
- Budget baseline state must be explicit (TBC at QG1 if unknown).
