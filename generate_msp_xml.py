"""
generate_msp_xml.py
Converts Trinity_Project_Schedule.csv → Trinity_Project_Schedule.xml
Microsoft Project XML format (schema: http://schemas.microsoft.com/project)

Usage:
    python generate_msp_xml.py
    python generate_msp_xml.py --csv path/to/schedule.csv --out path/to/output.xml

Requirements: Python 3.6+ (stdlib only — no external dependencies)
"""

import csv
import re
import argparse
from pathlib import Path
from xml.sax.saxutils import escape


# ─── Helpers ──────────────────────────────────────────────────────────────────

def parse_days(duration_str: str) -> int:
    """Extract integer day count from strings like '10 days', '0 days'."""
    m = re.search(r"(\d+)", duration_str or "")
    return int(m.group(1)) if m else 0


def days_to_iso(days: int) -> str:
    """Convert working days to MS Project ISO duration (PT{n}H0M0S, 8 hrs/day)."""
    return f"PT{days * 8}H0M0S" if days > 0 else "PT0H0M0S"


def to_xml_start(date_str: str) -> str:
    """Parse MM/DD/YY → 20YY-MM-DDT08:00:00."""
    date_str = (date_str or "").strip()
    if not date_str:
        return ""
    mm, dd, yy = date_str.split("/")
    return f"20{yy}-{mm.zfill(2)}-{dd.zfill(2)}T08:00:00"


def to_xml_finish(date_str: str, is_milestone: bool = False) -> str:
    """Parse MM/DD/YY → 20YY-MM-DDT17:00:00 (or T08:00:00 for 0-day milestones)."""
    date_str = (date_str or "").strip()
    if not date_str:
        return ""
    mm, dd, yy = date_str.split("/")
    time = "T08:00:00" if is_milestone else "T17:00:00"
    return f"20{yy}-{mm.zfill(2)}-{dd.zfill(2)}{time}"


def parse_predecessors(pred_str: str) -> list:
    """Return list of int predecessor IDs from strings like '14,15,16,17'."""
    return [
        int(p.strip())
        for p in (pred_str or "").replace('"', "").split(",")
        if p.strip().isdigit()
    ]


def split_resources(resource_str: str) -> list:
    """Split 'KPMG + IT PM' into ['KPMG', 'IT PM']."""
    return [r.strip() for r in (resource_str or "").split("+") if r.strip()]


# ─── Parse CSV ────────────────────────────────────────────────────────────────

def load_tasks(csv_path: Path) -> list:
    tasks = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        headers = next(reader)  # skip header row
        for row in reader:
            if not row or not row[0].strip().isdigit():
                continue

            task_id      = int(row[0])
            outline_lvl  = int(row[1])
            name         = row[2].strip()
            days         = parse_days(row[3])
            start_str    = row[4].strip()
            finish_str   = row[5].strip()
            pred_str     = row[6].strip() if len(row) > 6 else ""
            resource_str = row[7].strip() if len(row) > 7 else ""
            notes_str    = row[8].strip() if len(row) > 8 else ""
            milestone_flag = (row[9].strip().lower() == "yes") if len(row) > 9 else False

            is_milestone = milestone_flag or days == 0

            tasks.append({
                "id"          : task_id,
                "outline_level": outline_lvl,
                "name"        : name,
                "days"        : days,
                "duration"    : days_to_iso(days),
                "start"       : to_xml_start(start_str),
                "finish"      : to_xml_finish(finish_str, is_milestone),
                "predecessors": parse_predecessors(pred_str),
                "resources"   : resource_str,
                "resource_list": split_resources(resource_str),
                "notes"       : notes_str,
                "milestone"   : is_milestone,
            })
    return tasks


def find_summary_ids(tasks: list) -> set:
    """A task is a summary if the next task has a deeper outline level."""
    summary_ids = set()
    for i in range(len(tasks) - 1):
        if tasks[i + 1]["outline_level"] > tasks[i]["outline_level"]:
            summary_ids.add(tasks[i]["id"])
    return summary_ids


def build_resource_pool(tasks: list) -> tuple:
    """Return (sorted resource list, resource→UID dict)."""
    seen = set()
    for t in tasks:
        for r in t["resource_list"]:
            seen.add(r)
    resource_list = sorted(seen)
    resource_uid  = {r: i + 1 for i, r in enumerate(resource_list)}
    return resource_list, resource_uid


# ─── XML generation ───────────────────────────────────────────────────────────

def xml_task(t: dict, summary_ids: set) -> list:
    is_summary = t["id"] in summary_ids

    note_parts = []
    if t["resources"]:
        note_parts.append(f"Resources: {t['resources']}")
    if t["notes"]:
        note_parts.append(t["notes"])
    note_text = " | ".join(note_parts)

    lines = [
        "    <Task>",
        f"      <UID>{t['id']}</UID>",
        f"      <ID>{t['id']}</ID>",
        f"      <Name>{escape(t['name'])}</Name>",
        f"      <Duration>{t['duration']}</Duration>",
        "      <DurationFormat>7</DurationFormat>",
    ]
    if t["start"]:
        lines.append(f"      <Start>{t['start']}</Start>")
    if t["finish"]:
        lines.append(f"      <Finish>{t['finish']}</Finish>")
    lines += [
        f"      <OutlineLevel>{t['outline_level']}</OutlineLevel>",
        f"      <Summary>{1 if is_summary else 0}</Summary>",
        f"      <Milestone>{1 if t['milestone'] else 0}</Milestone>",
        "      <CalendarUID>-1</CalendarUID>",
        "      <IgnoreResourceCalendar>0</IgnoreResourceCalendar>",
        "      <EffortDriven>0</EffortDriven>",
    ]
    if note_text:
        lines.append(f"      <Notes>{escape(note_text)}</Notes>")

    for pred_id in t["predecessors"]:
        lines += [
            "      <PredecessorLink>",
            f"        <PredecessorUID>{pred_id}</PredecessorUID>",
            "        <Type>1</Type>",        # 1 = Finish-to-Start
            "        <CrossProject>0</CrossProject>",
            "        <LinkLag>0</LinkLag>",
            "        <LagFormat>7</LagFormat>",
            "      </PredecessorLink>",
        ]

    lines.append("    </Task>")
    return lines


def generate_xml(tasks: list, summary_ids: set,
                 resource_list: list, resource_uid: dict) -> list:
    lines = []

    # ── XML declaration + root ──
    lines += [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Project xmlns="http://schemas.microsoft.com/project">',
        "  <SaveVersion>14</SaveVersion>",
        "  <Name>Trinity IT Carve-Out Project</Name>",
        "  <Title>Project Trinity</Title>",
        "  <Manager>IT PM</Manager>",
        "  <StartDate>2026-07-01T08:00:00</StartDate>",
        "  <FinishDate>2029-12-31T17:00:00</FinishDate>",
        "  <ScheduleFromStart>1</ScheduleFromStart>",
        "  <CalendarUID>1</CalendarUID>",
        "  <DefaultStartTime>08:00:00</DefaultStartTime>",
        "  <DefaultFinishTime>17:00:00</DefaultFinishTime>",
        "  <MinutesPerDay>480</MinutesPerDay>",
        "  <MinutesPerWeek>2400</MinutesPerWeek>",
        "  <DaysPerMonth>20</DaysPerMonth>",
        "  <DefaultTaskType>0</DefaultTaskType>",
        "  <DefaultFixedCostAccrual>3</DefaultFixedCostAccrual>",
        "  <CriticalSlackLimit>0</CriticalSlackLimit>",
        "  <CurrencySymbol>€</CurrencySymbol>",
        "  <CurrencyCode>EUR</CurrencyCode>",
        "  <CurrencyDigits>2</CurrencyDigits>",
    ]

    # ── Calendar ──
    lines += [
        "  <Calendars>",
        "    <Calendar>",
        "      <UID>1</UID>",
        "      <Name>Standard</Name>",
        "      <IsBaseCalendar>1</IsBaseCalendar>",
        "      <IsBaselineCalendar>0</IsBaselineCalendar>",
        "      <WeekDays>",
        "        <WeekDay>",
        "          <DayType>1</DayType>",   # Sunday — non-working
        "          <DayWorking>0</DayWorking>",
        "        </WeekDay>",
        "        <WeekDay>",
        "          <DayType>7</DayType>",   # Saturday — non-working
        "          <DayWorking>0</DayWorking>",
        "        </WeekDay>",
        "      </WeekDays>",
        "    </Calendar>",
        "  </Calendars>",
    ]

    # ── Tasks ──
    lines.append("  <Tasks>")

    # Task 0 — MS Project requires a root project summary task
    lines += [
        "    <Task>",
        "      <UID>0</UID>",
        "      <ID>0</ID>",
        "      <Name>Project Trinity</Name>",
        "      <Duration>PT27520H0M0S</Duration>",
        "      <DurationFormat>7</DurationFormat>",
        "      <Start>2026-07-01T08:00:00</Start>",
        "      <Finish>2029-12-31T17:00:00</Finish>",
        "      <Summary>1</Summary>",
        "      <Milestone>0</Milestone>",
        "      <OutlineLevel>0</OutlineLevel>",
        "      <CalendarUID>-1</CalendarUID>",
        "      <IgnoreResourceCalendar>0</IgnoreResourceCalendar>",
        "    </Task>",
    ]

    for t in tasks:
        lines.extend(xml_task(t, summary_ids))

    lines.append("  </Tasks>")

    # ── Resources ──
    lines.append("  <Resources>")
    for r in resource_list:
        uid = resource_uid[r]
        lines += [
            "    <Resource>",
            f"      <UID>{uid}</UID>",
            f"      <ID>{uid}</ID>",
            f"      <Name>{escape(r)}</Name>",
            "      <Type>1</Type>",          # 1 = Work resource
            "      <IsNull>0</IsNull>",
            "      <CalendarUID>-1</CalendarUID>",
            "      <IsEnterprise>0</IsEnterprise>",
            "    </Resource>",
        ]
    lines.append("  </Resources>")

    # ── Assignments ──
    lines.append("  <Assignments>")
    a_uid = 1
    for t in tasks:
        for r in t["resource_list"]:
            rid = resource_uid.get(r)
            if rid is None:
                continue
            work_hrs = t["days"] * 8
            lines += [
                "    <Assignment>",
                f"      <UID>{a_uid}</UID>",
                f"      <TaskUID>{t['id']}</TaskUID>",
                f"      <ResourceUID>{rid}</ResourceUID>",
                "      <Units>1</Units>",
                f"      <Work>PT{work_hrs}H0M0S</Work>",
            ]
            if t["start"]:
                lines.append(f"      <Start>{t['start']}</Start>")
            if t["finish"]:
                lines.append(f"      <Finish>{t['finish']}</Finish>")
            lines.append("    </Assignment>")
            a_uid += 1
    lines.append("  </Assignments>")

    lines.append("</Project>")
    return lines, a_uid - 1


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Convert Trinity_Project_Schedule.csv to MS Project XML"
    )
    parser.add_argument(
        "--csv",
        type=Path,
        default=Path(__file__).parent / "Trinity_Project_Schedule.csv",
        help="Path to input CSV (default: same directory as script)",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=Path(__file__).parent / "Trinity_Project_Schedule.xml",
        help="Path to output XML (default: same directory as script)",
    )
    args = parser.parse_args()

    print(f"Reading : {args.csv}")
    tasks        = load_tasks(args.csv)
    summary_ids  = find_summary_ids(tasks)
    resource_list, resource_uid = build_resource_pool(tasks)

    xml_lines, assignment_count = generate_xml(
        tasks, summary_ids, resource_list, resource_uid
    )

    args.out.write_text("\n".join(xml_lines), encoding="utf-8")

    milestone_count = sum(1 for t in tasks if t["milestone"])
    print(f"Written : {args.out}")
    print(f"  Tasks      : {len(tasks)}")
    print(f"  Resources  : {len(resource_list)}")
    print(f"  Assignments: {assignment_count}")
    print(f"  Milestones : {milestone_count}")
    print(f"  Summary tasks: {len(summary_ids)}")
    print(f"  File size  : {args.out.stat().st_size / 1024:.1f} KB")


if __name__ == "__main__":
    main()
