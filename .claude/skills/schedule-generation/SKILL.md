---
name: schedule-generation
description: "Use when creating or adapting carve-out schedules, phase and QG milestones, task dependencies, and XLSX/XML outputs from the canonical schedule workflow."
---

# Schedule Generation

## Purpose
Generate a project schedule as a formatted Excel workbook (XLSX) and MS Project XML from the canonical script pattern.
The XML generator (`generate_msp_xml.py`) requires a CSV as input — the CSV is written permanently to the output folder alongside XLSX and XML.

## Canonical Sources
- `generate_msp_xml.py` — canonical CSV→XML converter (do not duplicate its logic)
- Existing project generators (e.g. `generate_bravo_schedule.py`) — **format/structure reference only**; never copy their task lists, dates, resources, or project-specific content

## Steps
1. Create a new `generate_{ProjectName}_schedule.py` script from scratch.
2. Define the TASKS list using the required column schema, with all tasks, dates, durations, resources, and milestones **derived from the current project's input parameters** (start date, GoLive date, completion date, number of sites, number of users, number of applications, carve-out model, TSA involvement, etc.).
3. Plan realistic phase durations and task breakdowns based on the project's scope and complexity — do NOT reuse task lists from any other project.
4. Preserve the standard milestone sequence and QG0–QG5 structure per the Bosch project framework.
5. Set output paths to `{ProjectName}/`.
6. Define `XLSX_PATH`, `CSV_PATH`, and `XML_PATH` in the output folder.
7. In the `if __name__ == "__main__":` block: call `_generate_excel()`, then `_write_temp_csv(CSV_PATH)` (kept permanently), then call `generate_msp_xml.py` via subprocess passing `CSV_PATH`.
8. Run the schedule script to produce XLSX, CSV, and XML.

## Content Generation Rules
- **Never copy task lists, resource assignments, dates, or durations from any existing project** (Bravo, AlphaX, Falcon, Trinity, Hamburger, or any other).
- Derive all phase durations proportionally from the project timeline (start → GoLive → completion).
- Scale task granularity to the project scope: more sites/users/applications = more detailed breakdown tasks.
- Include workstream-specific tasks relevant to the carve-out model (Stand Alone vs Integration).
- If SAP is in scope, include SAP-specific tasks (system copy, data migration, cutover planning).
- If TSA is involved, include TSA setup, service definition, and exit planning tasks.
- Resource names must reflect the actual project's organizational structure, not copied from other projects.

## Task Schema
`(ID, Outline Level, Name, Duration, Start, Finish, Predecessors, Resource Names, Notes, Milestone)`

## Rules — Task Data
- Date format in CSV tasks: `YYYY-MM-DD` (ISO format — unambiguous, no 2-digit year misparse).
- `date_plus_days(base, days)` must use **calendar days only**: `return base + timedelta(days=days)`. Never multiply by 7/5 — that inflates all dates by ~40%.
- Milestones: `"0 days"` duration and `"Yes"` in Milestone column.
- Predecessors: comma-separated task IDs (e.g. `"14,15,16,17"` or `"14"`).
- Resources: plus-separated names (e.g. `"KPMG + IT PM"`).
- Semicolon-separated predecessor values (e.g. `"14;15"`) are **not** supported — use commas only.
- **Predecessors must NEVER reference a summary task (outline level 1 or 2)** — see XML circular reference rules below.

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

## XML Circular Reference Prevention — Critical
MS Project reports "circular relationship" and sets ALL task durations to 0 on import when these rules are violated:

1. **Never write `<PredecessorLink>` on summary tasks** (outline level 1 or 2). MS Project derives summary timing from children; predecessors on summaries create circular scheduling. `generate_msp_xml.py` enforces this with `if is_summary: break`.
2. **Never reference a summary task ID as a predecessor of any detail task.** A detail task that depends on its own parent summary creates an indirect circular reference. `generate_msp_xml.py` enforces this with `if pred_id in summary_ids: continue`.
3. **Always point the first detail tasks of a new phase to the last DETAIL task of the previous phase** — not to the phase summary row.
4. **Gate milestones (QG0, QG1, QG2/3, QG4, QG5) must reference the last detail task(s) of their workstreams** — not the workstream summary rows.

Example — QG2/3 gate (task 81) should reference:
- Last task of 2.1 Infrastructure Design (e.g. task 62)
- Last task of 2.2 ERP Design (e.g. task 68)
- Last task of 2.3 App Migration Design (e.g. task 72)
- Last task of 2.4 Client Design (e.g. task 75)
- Last task of 2.5 Cutover Strategy (e.g. task 80)

NOT the workstream summaries (55, 63, 69, 73, 76).

## Subprocess Pattern for generate_{ProjectName}_schedule.py
```python
CSV_PATH  = HERE / "active-projects" / PROJECT_NAME / f"{PROJECT_NAME}_Project_Schedule.csv"

# In main:
_generate_excel()
_write_temp_csv(CSV_PATH)
result = subprocess.run(
    [sys.executable,
     str(HERE / "generate_msp_xml.py"),
     "--csv", str(CSV_PATH),
     "--out", str(XML_PATH),
     "--project", PROJECT_NAME],
    capture_output=True, text=True
)
```
CSV is kept permanently — do NOT use `tempfile` or delete the CSV.

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
Schedule output is complete only when all three files exist:
- `{ProjectName}/{ProjectName}_Project_Schedule.xlsx` ← **primary human-readable deliverable**
- `{ProjectName}/{ProjectName}_Project_Schedule.csv` ← kept permanently for inspection and XML re-generation
- `{ProjectName}/{ProjectName}_Project_Schedule.xml` ← MS Project import file
- Schedule must be generated before cost plan, risk register, charter, and dashboards.

## Mandatory Carve-Out Phase & Migration Rules
These rules are **non-negotiable** and must be validated before generating any schedule:

1. **ALL migrations are PRE-GoLive activities** — device reimaging, M365 mailbox migration, OneDrive migration, application waves — ALL must complete before QG4 gate.
2. **After GoLive: NO migrations allowed.** Phase 4 contains only: GoLive cutover, Hypercare, TSA exit, Programme Closure.
3. **Hypercare is mandatory 90 calendar days** — never less. It starts the day after GoLive Day 1.
4. **Application wave activation (Wave 2, Wave 3) during Hypercare is NOT migration** — it is phased activation of already-packaged applications and is allowed post-GoLive.
5. **Gate sequence (non-negotiable):** QG0 → QG1 → QG2/3 → *(all Phase 3 tasks complete)* → **QG4 pre-GoLive gate** → **GoLive Day 1** → 90-day Hypercare → TSA Exit → **QG5** → Project Closure.
6. **QG4 must be an explicit milestone** positioned after all Phase 3 tasks complete and before GoLive Day 1.
