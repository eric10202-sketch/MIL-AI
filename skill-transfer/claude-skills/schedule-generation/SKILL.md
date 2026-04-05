---
name: schedule-generation
description: "Use when creating or adapting carve-out schedules with strict quality gates, phase milestones, task dependencies, and XLSX/XML outputs."
---

# Schedule Generation

## Purpose
Generate a project schedule as Excel workbook (XLSX) + MS Project XML from canonical script pattern. The XML generator (`generate_msp_xml.py`) requires a CSV — kept permanently alongside XLSX and XML.

---

## Canonical Sources
- `generate_msp_xml.py` — canonical CSV→XML converter; do NOT duplicate logic
- Existing project generators (e.g. `generate_bravo_schedule.py`) — **format/structure reference ONLY**; never copy task lists, dates, resources, or project-specific content

---

## Generation Steps
1. Create new `generate_{ProjectName}_schedule.py` script from scratch
2. Define TASKS list with all tasks, dates, durations, resources derived from **current project parameters** (start date, GoLive, completion date, sites, users, applications, carve-out model, TSA)
3. Plan realistic phase durations based on scope/complexity — do NOT reuse other projects' task lists
4. Preserve standard milestone sequence: QG0 → QG1 → QG2&3 → QG4 → GoLive → Hypercare → QG5 → Closure
5. Set output paths to `active-projects/{ProjectName}/`
6. Define `XLSX_PATH`, `CSV_PATH`, `XML_PATH` in output folder
7. In `if __name__ == "__main__":` call `_generate_excel()`, `_write_temp_csv()`, subprocess `generate_msp_xml.py`
8. Run script to produce XLSX, CSV, and XML

---

## Content Generation Rules
- **Never copy** task lists, resources, dates, or durations from any existing project
- Derive phase durations proportionally from timeline (start → GoLive → completion)
- Scale task granularity to project scope: more sites/users/apps = more detail
- Include workstream tasks relevant to carve-out model (Stand Alone vs Integration)
- If SAP in scope: include system copy, data migration, cutover planning
- If TSA involved: include TSA setup, service definition, exit planning
- Resources must reflect actual project structure, not copied

---

## Task Schema
`(ID, Outline Level, Name, Duration, Start, Finish, Predecessors, Resource Names, Notes, Milestone)`

---

## Rules — Task Data

**Dates & Calculations:**
- Date format: `YYYY-MM-DD` (ISO, unambiguous)
- `date_plus_days(base, days)` must use calendar days only: `return base + timedelta(days=days)`
- Never multiply by 7/5 — inflates dates by ~40%

**Task Details:**
- Milestones: `1` day duration, `"Yes"` in Milestone column
- Predecessors: comma-separated (e.g. `"14,15,16,17"`)
- Resources: plus-separated (e.g. `"KPMG + IT PM"`)
- Semicolon-separated predecessors (`"14;15"`) NOT supported — use commas only
- **Predecessors MUST reference detail tasks only (level 3+), NEVER summary tasks (level 1–2)**

---

## THREE MANDATORY CARVE-OUT RULES

### **RULE 1: Strict Quality Gate Sequence (Binding)**

```
QG0 (Intake) → QG1 (Concept) → QG2&3 (Build & Test) → 
QG4 (Pre-GoLive Check) → [Final Readiness] → 
GoLive Day 1 → 90-day Hypercare → QG5 (Completion) → Closure
```

Each gate is a **1-day milestone** (`Milestone = Yes`) that gates entry to next phase. Cannot be modified.

---

### **RULE 2: QG4 and GoLive on DIFFERENT Dates**

QG4 approval and GoLive cutover **CANNOT be same calendar date.**

**Between QG4 and GoLive, insert mandatory Final Readiness & Open Item Closure phase (5 days minimum):**
- ✓ Verify all UAT defects resolved (QA sign-off)
- ✓ Confirm all QG4 action items closed
- ✓ Final infrastructure & system readiness checks (no critical alerts)
- ✓ Cutover rollback test executed
- ✓ Business & Steering Committee sign-off

**GoLive proceeds only after Final Readiness complete.** This prevents critical gaps cascading to production.

**Example:**
- QG4 Approval: Nov 1, 2026
- Final Readiness: Nov 2–6, 2026 (5 days)
- GoLive Cutover: Nov 7, 2026

---

### **RULE 3: No Development or Migration After GoLive**

**Post-GoLive scope = STABILIZATION ONLY.** Mandatory 90-day Hypercare contains:
- ✓ 24/7 L3 support & incident triage
- ✓ Hotfixes (bug fixes only, **NOT new features**)
- ✓ User training & enablement
- ✓ Knowledge transfer & documentation

**Explicit prohibitions post-GoLive:**
- ✗ No new development or features
- ✗ No migrations (data, device, M365)
- ✗ No deployment waves (pre-packaged only, configured pre-QG4)
- ✗ No architecture changes
- ✗ No unplanned enhancements

**Hypercare = support + stabilization ONLY.**

---

## Supporting Pre-GoLive Rules
- **ALL migrations complete before QG4:** device reimaging, M365 migration, OneDrive migration, app waves, data segregation
- **Hypercare duration:** static 90 calendar days minimum (never less), starts day after GoLive
- **Wave 2/3 post-GoLive:** allowed ONLY if pre-packaged and pre-configured before QG4 (phased activation, not migration)

---

## XML Output Rules — Critical (MS Project Import)

Generated via `generate_msp_xml.py` — do NOT hand-write XML.

| Element | Correct | Wrong (Causes Issues) |
|---|---|---|
| `<Duration>` | `PT{days*8}H0M0S` (working hours) | `P{days}D` (calendar ISO — errors) |
| `<Start>` | `20YY-MM-DDT08:00:00` | `T00:00:00` (wrong time) |
| `<Finish>` (normal) | `20YY-MM-DDT17:00:00` | `T00:00:00` (wrong time) |
| `<Finish>` (milestone) | `20YY-MM-DDT08:00:00` | `T17:00:00` (wrong time) |
| `<DurationFormat>` | `7` (on every task) | Missing (parsing fails) |
| `<Calendars>` | Standard 5-day calendar | Missing (date drift) |
| `<Resources>` | All assigned resources | Missing (import fails) |
| `<Assignments>` | All assignments | Missing (unassigned) |

---

## XML Circular Reference Prevention

MS Project sets **ALL durations to 0** on import if violated:

1. **Never write `<PredecessorLink>` on summary tasks** (level 1–2)
   - MS Project derives timing from children only
   - Creates circular scheduling
   - `generate_msp_xml.py` enforces: `if is_summary: break`

2. **Never reference summary task ID as predecessor** of detail task
   - Indirect circular reference (depends on own parent)
   - `generate_msp_xml.py` enforces: `if pred_id in summary_ids: continue`

3. **First detail tasks of new phase** → point to last DETAIL task of previous phase (not phase summary)

4. **Gate milestones** → reference last DETAIL task(s) of workstreams (not summaries)
   - Example: QG2&3 references final tasks of all Phase 2 workstreams (tasks 62, 68, 72, 75, 80), NOT workstream summaries

---

## MS Project Element Ordering (Strict Sequence)

Missing or mis-ordered elements cause silent date drift (e.g. GoLive months late).

```xml
<Task>
  <UID>…</UID>
  <ID>…</ID>
  <Name>…</Name>
  <TaskMode>1</TaskMode>              <!-- MUST follow <Name> immediately -->
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
  <ConstraintType>2</ConstraintType>   <!-- Must Start On (non-summary only) -->
  <ConstraintDate>20YY-MM-DDT08:00:00</ConstraintDate>
  <CalendarUID>-1</CalendarUID>
</Task>
```

**Critical Rules:**
- `<TaskMode>1</TaskMode>` anywhere other than immediately after `<Name>` is silently ignored
- `<ManualStart>` / `<ManualFinish>` are fields MS Project **displays** — `<Start>` / `<Finish>` alone insufficient
- `<ConstraintType>2</ConstraintType>` + `<ConstraintDate>` provides hard date pin
- Summary tasks should NOT have `<ConstraintType>` — inherit from children

---

## Subprocess Pattern for generate_{ProjectName}_schedule.py

```python
CSV_PATH = HERE / "active-projects" / PROJECT_NAME / f"{PROJECT_NAME}_Project_Schedule.csv"

# In main:
_generate_excel()
_write_temp_csv(CSV_PATH)
result = subprocess.run([sys.executable, str(HERE / "generate_msp_xml.py"),
                        "--csv", str(CSV_PATH), "--out", str(XML_PATH),
                        "--project", PROJECT_NAME], capture_output=True, text=True)
if result.returncode != 0:
    print(f"✗ XML generation failed:\n{result.stderr}")
    sys.exit(1)
```

**CSV kept permanently** — do NOT use `tempfile` or delete

---

## XLSX Formatting Standards (Bosch Blue Theme)

| Row Type | Fill | Font | Bold |
|---|---|---|---|
| Header | `#002147` (near-black blue) | White | Yes |
| Phase (L1) | `#003B6E` (dark Bosch blue) | White | Yes |
| Section (L2, non-milestone) | `#0066CC` (mid blue) | White | Yes |
| Milestone (any level) | `#FFF2CC` (amber) | **Black** | Yes |
| Detail (even) | `#EFF4FB` (light blue) | Black | No |
| Detail (odd) | White | Black | No |

**Additional:**
- Freeze pane on row 2 (below header)
- Auto-filter on header row
- Name indented by outline level (2 spaces per level beyond 1)
- Milestone rows: **always black bold** (overrides section white-font rule)

---

## Output Deliverables

Schedule complete only when **all three files exist:**
- `active-projects/{ProjectName}/{ProjectName}_Project_Schedule.xlsx` ← **primary human-readable deliverable**
- `active-projects/{ProjectName}/{ProjectName}_Project_Schedule.csv` ← kept permanently for inspection/XML re-generation
- `active-projects/{ProjectName}/{ProjectName}_Project_Schedule.xml` ← MS Project import file

**Schedule must be generated BEFORE cost plan, risk register, charter, and dashboards.**

---

## Verification Checklist

Before marking schedule complete, verify:
- [ ] QG0, QG1, QG2&3, QG4, GoLive, QG5 all present as 1-duration milestones (`Milestone = Yes`)
- [ ] **QG4 and GoLive on DIFFERENT dates** (minimum 5 day gap)
- [ ] Final Readiness phase tasks exist between QG4 and GoLive (UAT closure, open item verification, sign-offs)
- [ ] All pre-GoLive migrations/deployments complete before QG4
- [ ] No development or migration tasks after GoLive cutover milestone
- [ ] Hypercare spans exactly 90 calendar days post-GoLive
- [ ] Hypercare tasks = stabilization only (support, training, knowledge transfer)
- [ ] All XLSX formatting applied (Bosch colors, freeze pane, auto-filter)
- [ ] CSV and XML files generated and validated
- [ ] All predecessors reference detail tasks (level 3+), never summaries (level 1–2)
