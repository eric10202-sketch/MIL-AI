---
name: schedule-generation
description: "Use when creating or adapting carve-out schedules, phase and QG milestones, task dependencies, and CSV/XML outputs from the canonical schedule workflow."
---

# Schedule Generation

## Purpose
Generate a project schedule CSV and MS Project XML from the canonical script pattern.

## Canonical Sources
- `generate_msp_xml.py` — canonical CSV→XML converter (do not duplicate its logic)
- `generate_falcon_schedule.py` — reference project-specific schedule generator
- `Trinity_Project_Schedule.csv` — reference CSV

## Steps
1. Copy `generate_falcon_schedule.py` to `generate_{ProjectName}_schedule.py`.
2. Update project header comment and output paths to `{ProjectName}/`.
3. Replace the TASKS list using the required column schema.
4. Preserve milestone sequence and QG1–QG5 structure.
5. In the `if __name__ == "__main__":` block, ensure the subprocess call to `generate_msp_xml.py` includes `"--project", "Project {ProjectName}"` (see pattern below).
6. Run the schedule script to produce both CSV and XML.

## Task Schema
`(ID, Outline Level, Name, Duration, Start, Finish, Predecessors, Resource Names, Notes, Milestone)`

## Rules — Task Data
- Date format in CSV tasks: `MM/DD/YY`.
- Milestones: `"0 days"` duration and `"Yes"` in Milestone column.
- Predecessors: comma-separated task IDs (e.g. `"14,15,16,17"` or `"14"`).
- Resources: plus-separated names (e.g. `"KPMG + IT PM"`).
- Semicolon-separated predecessor values (e.g. `"14;15"`) are **not** supported — use commas only.

## XML Output Rules — Critical (MS Project import)
These are the exact rules enforced by `generate_msp_xml.py`. Do NOT hand-write or override XML — always generate via the script.

| Element | Correct value | Wrong (old bug) |
|---|---|---|
| `<Duration>` | `PT{days*8}H0M0S` (working hours) | `P{days}D` (calendar ISO — will import with errors) |
| `<Start>` | `20YY-MM-DDT08:00:00` | `T00:00:00` |
| `<Finish>` (normal tasks) | `20YY-MM-DDT17:00:00` | `T00:00:00` |
| `<Finish>` (milestones) | `20YY-MM-DDT08:00:00` | `T00:00:00` |
| `<DurationFormat>` | `7` (must be present on every task) | missing |
| `<Calendars>` section | Must be present with Standard 5-day calendar | missing |
| `<Resources>` section | Must be present | missing |
| `<Assignments>` section | Must be present | missing |
| Task 0 (root summary) | Required; `<Duration>PT{total_working_hours}H0M0S` | missing or wrong |

## Subprocess Pattern for generate_{ProjectName}_schedule.py
```python
result = subprocess.run(
    [sys.executable,
     str(HERE / "generate_msp_xml.py"),
     "--csv", str(CSV_PATH),
     "--out", str(XML_PATH),
     "--project", "Project {ProjectName}"],
    capture_output=True, text=True
)
```
The `--project` argument sets the `<Name>`, `<Title>`, and root Task 0 `<Name>` in the XML.
`generate_msp_xml.py` derives project `<StartDate>` and `<FinishDate>` automatically from the task data — do **not** hardcode these.

## Python Interpreter
- Use `C:/Program Files/px/python.exe` to run scripts on this machine.

## Output Completeness
Schedule output is complete only when both files exist:
- `{ProjectName}/{ProjectName}_Project_Schedule.csv`
- `{ProjectName}/{ProjectName}_Project_Schedule.xml`
- Schedule must be generated before cost plan, risk register, charter, and dashboards.
