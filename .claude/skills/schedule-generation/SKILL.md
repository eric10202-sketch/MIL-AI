---
name: schedule-generation
description: "Use when creating or adapting carve-out schedules, phase and QG milestones, task dependencies, and CSV/XML outputs from the canonical schedule workflow."
---

# Schedule Generation

## Purpose
Generate a project schedule CSV and MS Project XML from the canonical script pattern.

## Canonical Sources
- generate_alphax_schedule.py
- generate_msp_xml.py
- Trinity_Project_Schedule.csv

## Steps
1. Copy generate_alphax_schedule.py to generate_{ProjectName}_schedule.py.
2. Update project configuration and output paths to {ProjectName}/.
3. Replace TASKS list using required column schema.
4. Preserve milestone sequence and QG1-QG5 structure.
5. Run schedule script to produce both CSV and XML.

## Task Schema
(ID, Outline Level, Name, Duration, Start, Finish, Predecessors, Resource Names, Notes, Milestone)

## Rules
- Date format in tasks: MM/DD/YY.
- Milestones: 0 days and Milestone = Yes.
- Predecessors are task IDs.
- Resources use plus-separated names.
- Schedule must be generated before cost, risk, charter, and dashboards.
