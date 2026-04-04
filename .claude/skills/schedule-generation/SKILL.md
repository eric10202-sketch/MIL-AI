---
name: schedule-generation
description: "Use when creating or adapting carve-out schedules, phase and QG milestones, task dependencies, and XLSX/XML outputs from the canonical schedule workflow."
---

# Schedule Generation

## Purpose
Generate a project schedule as a formatted Excel workbook (XLSX) and MS Project XML from the canonical script pattern.
The XML generator (`generate_msp_xml.py`) requires a CSV as input — a temporary CSV is written to a system temp path and deleted automatically after XML generation. No CSV file is kept in the output folder.

## Canonical Sources
- `generate_msp_xml.py` — canonical CSV→XML converter (do not duplicate its logic)
- `generate_bravo_schedule.py` — reference project-specific schedule generator (includes XLSX output)
- `generate_falcon_schedule.py` — legacy reference (CSV-only, XLSX not yet added)
- `Trinity_Project_Schedule.csv` — legacy reference CSV

## Steps
1. Copy `generate_bravo_schedule.py` to `generate_{ProjectName}_schedule.py`.
2. Update project header comment and output paths to `{ProjectName}/`.
3. Replace the TASKS list using the required column schema.
4. Preserve milestone sequence and QG0–QG5 structure.
5. In the `if __name__ == "__main__":` block: use `tempfile.mkstemp(suffix='.csv')` to create a temp CSV, pass it to `generate_msp_xml.py` via subprocess, then delete it in a `finally` block.
6. Ensure `XLSX_PATH` and `XML_PATH` are defined; no `CSV_PATH` in the output folder.
7. Run the schedule script to produce XLSX and XML.

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

## XML Task Scheduling — Mandatory Elements to Prevent Date Drift
MS Project recalculates all task dates from the predecessor chain on import unless ALL of the following are present and correctly ordered. Missing or mis-ordered elements cause dates to silently drift (e.g. GoLive appearing months late).

### Project header (required)
```xml
<DefaultTaskType>1</DefaultTaskType>
<NewTasksAreManual>1</NewTasksAreManual>
```

### Per-task element order (strict — MS Project parses sequentially)
```xml
<Task>
  <UID>…</UID>
  <ID>…</ID>
  <Name>…</Name>
  <TaskMode>1</TaskMode>          <!-- MUST be immediately after <Name> -->
  <Duration>PT…</Duration>
  <DurationFormat>7</DurationFormat>
  <ManualDuration>PT…</ManualDuration>
  <Start>20YY-MM-DDT08:00:00</Start>
  <ManualStart>20YY-MM-DDT08:00:00</ManualStart>
  <Finish>20YY-MM-DDT…</Finish>
  <ManualFinish>20YY-MM-DDT…</ManualFinish>
  <OutlineLevel>…</OutlineLevel>
  <Summary>…</Summary>
  <Milestone>…</Milestone>
  <ConstraintType>2</ConstraintType>      <!-- Must Start On — non-summary tasks only -->
  <ConstraintDate>20YY-MM-DDT08:00:00</ConstraintDate>
  <CalendarUID>-1</CalendarUID>
  …
</Task>
```

**Rules:**
- `<TaskMode>1</TaskMode>` placed anywhere other than immediately after `<Name>` is silently ignored — MS Project falls back to auto-scheduling.
- `<ManualStart>` / `<ManualFinish>` are the fields MS Project actually **displays** for manually-scheduled tasks. `<Start>` / `<Finish>` alone are insufficient.
- `<ConstraintType>2</ConstraintType>` ("Must Start On") + `<ConstraintDate>` provides a hard date pin independent of scheduling mode.
- Summary tasks (parent rows) should NOT have `<ConstraintType>` — they inherit bounds from their children.

## Subprocess Pattern for generate_{ProjectName}_schedule.py
```python
_fd, _tmp_csv = tempfile.mkstemp(suffix=".csv")
os.close(_fd)
_write_temp_csv(Path(_tmp_csv))
try:
    result = subprocess.run(
        [sys.executable,
         str(HERE / "generate_msp_xml.py"),
         "--csv", _tmp_csv,
         "--out", str(XML_PATH),
         "--project", "Project {ProjectName}"],
        capture_output=True, text=True
    )
finally:
    try:
        os.unlink(_tmp_csv)
    except OSError:
        pass
```
The `--project` argument sets the `<Name>`, `<Title>`, and root Task 0 `<Name>` in the XML.
`generate_msp_xml.py` derives project `<StartDate>` and `<FinishDate>` automatically from the task data — do **not** hardcode these.

## Python Interpreter
- Use `C:/Program Files/px/python.exe` to run scripts on this machine.

## XLSX Formatting Standards
All schedule XLSX outputs must follow the Bosch blue theme applied in `generate_bravo_schedule.py`:

| Row type | Fill colour | Font colour | Bold |
|---|---|---|---|
| Column header | `#002147` (near-black blue) | White | Yes |
| Phase (Outline Level 1) | `#003B6E` (dark Bosch blue) | White | Yes |
| Section (Outline Level 2, non-milestone) | `#0066CC` (mid blue) | White | Yes |
| Milestone (any level, Milestone = Yes) | `#FFF2CC` (amber) | **Black** | Yes |
| Detail row (even) | `#EFF4FB` (light blue) | Black | No |
| Detail row (odd) | White | Black | No |

- Freeze pane on row 2 (below header).
- Auto-filter on header row.
- Name column indented by outline level (2 spaces per level beyond 1).
- Milestone rows: **always black bold text** — the section-level white-font override must be skipped for milestone rows.

## Output Completeness
Schedule output is complete only when both files exist:
- `{ProjectName}/{ProjectName}_Project_Schedule.xlsx` ← **primary human-readable deliverable**
- `{ProjectName}/{ProjectName}_Project_Schedule.xml` ← MS Project import file
- No CSV file is produced in the output folder.
- Schedule must be generated before cost plan, risk register, charter, and dashboards.
