#!/usr/bin/env python3
"""
generate_carveout_schedule_direct_xml.py
Produces MS Project XML schedule for a generic IT carve-out.
No intermediate CSV required.
"""

import argparse
import csv
import sys
import os
sys.path.insert(0, os.path.join(os.path.expanduser("~"), "py_packages"))
from pathlib import Path
from datetime import datetime
from xml.sax.saxutils import escape
import pandas as pd

REQUIRED_FIELDS = {
    "project_name": "Project name",
    "deal_closing_date": "Deal closing date (YYYY-MM-DD)",
    "signing_date": "Signing date (YYYY-MM-DD)",
    "tsa_exit_date": "TSA exit date (YYYY-MM-DD)",
    "day1_date": "Day 1 go-live date (YYYY-MM-DD)",
    "closure_date": "Project closure date (YYYY-MM-DD)",
    "num_sites": "Number of sites",
    "num_users": "Number of users",
    "carveout_model": "Carve-out model (Stand Alone / Integration / Combination)",
}


def prompt_for_required(cfg: dict) -> dict:
    for key, label in REQUIRED_FIELDS.items():
        while not cfg.get(key):
            value = input(f"{label}: ").strip()
            if value:
                cfg[key] = value
            else:
                print("  ➜ Required value missing; please enter value.")
    return cfg


def parse_date(value: str) -> datetime:
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError(f"Invalid date format for '{value}'. Expected YYYY-MM-DD.") from exc


def validate_date_sequence(cfg: dict):
    d = parse_date(cfg["deal_closing_date"])
    s = parse_date(cfg["signing_date"])
    t = parse_date(cfg["tsa_exit_date"])
    g = parse_date(cfg["day1_date"])
    c = parse_date(cfg["closure_date"])
    if not (d <= s <= t <= g <= c):
        raise ValueError(
            "Dates must be chronological: deal_closing_date <= signing_date <= tsa_exit_date <= day1_date <= closure_date"
        )


def iso_duration_from_days(days: int) -> str:
    return f"PT{max(0, days) * 8}H0M0S"


def to_xml_datetime(date_str: str, milestone: bool = False) -> str:
    dt = parse_date(date_str)
    if milestone:
        return dt.strftime("%Y-%m-%dT08:00:00")
    return dt.strftime("%Y-%m-%dT17:00:00")


def build_tasks(cfg: dict) -> list:
    core = [
        {
            "id": 1,
            "outline_level": 1,
            "name": "Deal closing",
            "days": 0,
            "start": cfg["deal_closing_date"],
            "finish": cfg["deal_closing_date"],
            "predecessors": [],
            "resources": "",
            "notes": "Deal closing milestone",
            "milestone": True,
        },
        {
            "id": 2,
            "outline_level": 1,
            "name": "Signing",
            "days": 0,
            "start": cfg["signing_date"],
            "finish": cfg["signing_date"],
            "predecessors": [1],
            "resources": "",
            "notes": "Signing milestone",
            "milestone": True,
        },
        {
            "id": 3,
            "outline_level": 1,
            "name": "TSA Exit",
            "days": 0,
            "start": cfg["tsa_exit_date"],
            "finish": cfg["tsa_exit_date"],
            "predecessors": [2],
            "resources": "",
            "notes": "TSA exit milestone",
            "milestone": True,
        },
        {
            "id": 4,
            "outline_level": 1,
            "name": "GoLive / Day 1",
            "days": 0,
            "start": cfg["day1_date"],
            "finish": cfg["day1_date"],
            "predecessors": [3],
            "resources": "",
            "notes": "Day 1 milestone",
            "milestone": True,
        },
        {
            "id": 5,
            "outline_level": 1,
            "name": "Project closure",
            "days": 0,
            "start": cfg["closure_date"],
            "finish": cfg["closure_date"],
            "predecessors": [4],
            "resources": "",
            "notes": "Project closure milestone",
            "milestone": True,
        },
    ]

    phase1_days = max(1, (parse_date(cfg["signing_date"]) - parse_date(cfg["deal_closing_date"])).days)
    phase2_days = max(1, (parse_date(cfg["tsa_exit_date"]) - parse_date(cfg["signing_date"])).days)
    phase3_days = max(1, (parse_date(cfg["day1_date"]) - parse_date(cfg["tsa_exit_date"])).days)
    phase4_days = max(1, (parse_date(cfg["closure_date"]) - parse_date(cfg["day1_date"])).days)

    core.extend([
        {
            "id": 6,
            "outline_level": 2,
            "name": "Governance and initiation",
            "days": phase1_days,
            "start": cfg["deal_closing_date"],
            "finish": cfg["signing_date"],
            "predecessors": [1],
            "resources": "PMO",
            "notes": f"Governance and kickoff in support of {cfg['project_name']} ({cfg['num_sites']} sites, {cfg['num_users']} users)",
            "milestone": False,
        },
        {
            "id": 7,
            "outline_level": 2,
            "name": "TSA service catalog and exit criteria",
            "days": phase2_days,
            "start": cfg["signing_date"],
            "finish": cfg["tsa_exit_date"],
            "predecessors": [2],
            "resources": "Legal + IT",
            "notes": "Define TSA scope and exit metrics",
            "milestone": False,
        },
        {
            "id": 8,
            "outline_level": 2,
            "name": "Day 1 cutover preparation",
            "days": phase3_days,
            "start": cfg["tsa_exit_date"],
            "finish": cfg["day1_date"],
            "predecessors": [3],
            "resources": "All Workstreams",
            "notes": "Plan Day 1 cutover in detail",
            "milestone": False,
        },
        {
            "id": 9,
            "outline_level": 2,
            "name": "Hypercare and project closure",
            "days": phase4_days,
            "start": cfg["day1_date"],
            "finish": cfg["closure_date"],
            "predecessors": [4],
            "resources": "IT Ops",
            "notes": "Stabilize and close project",
            "milestone": False,
        },
    ])
    return core


def find_summary_ids(tasks: list) -> set:
    summary = set()
    for i in range(len(tasks) - 1):
        if tasks[i + 1]["outline_level"] > tasks[i]["outline_level"]:
            summary.add(tasks[i]["id"])
    return summary


def build_resource_pool(tasks: list):
    seen = {}
    for t in tasks:
        for r in (t.get("resources") or "").split("+"):
            rr = r.strip()
            if rr:
                seen[rr] = seen.get(rr, 0) + 1
    ordered = sorted(seen.keys())
    return ordered, {n: i + 1 for i, n in enumerate(ordered)}


def task_to_xml(task: dict, summary_ids: set) -> str:
    is_summary = 1 if task["id"] in summary_ids else 0
    milestone = 1 if task["milestone"] else 0
    dur = iso_duration_from_days(task["days"])
    start = to_xml_datetime(task["start"], milestone)
    finish = to_xml_datetime(task["finish"], milestone)
    note_parts = []
    if task.get("resources"):
        note_parts.append(f"Resources: {task['resources']}")
    if task.get("notes"):
        note_parts.append(task["notes"])
    notes = escape(" | ".join(note_parts)) if note_parts else ""

    lines = [
        "    <Task>",
        f"      <UID>{task['id']}</UID>",
        f"      <ID>{task['id']}</ID>",
        f"      <Name>{escape(task['name'])}</Name>",
        f"      <Duration>{dur}</Duration>",
        "      <DurationFormat>7</DurationFormat>",
        f"      <Start>{start}</Start>",
        f"      <Finish>{finish}</Finish>",
        f"      <OutlineLevel>{task['outline_level']}</OutlineLevel>",
        f"      <Summary>{is_summary}</Summary>",
        f"      <Milestone>{milestone}</Milestone>",
        "      <CalendarUID>-1</CalendarUID>",
        "      <IgnoreResourceCalendar>0</IgnoreResourceCalendar>",
        "      <EffortDriven>0</EffortDriven>",
    ]
    if notes:
        lines.append(f"      <Notes>{notes}</Notes>")
    for p in task.get("predecessors", []):
        lines.extend([
            "      <PredecessorLink>",
            f"        <PredecessorUID>{p}</PredecessorUID>",
            "        <Type>1</Type>",
            "        <CrossProject>0</CrossProject>",
            "        <LinkLag>0</LinkLag>",
            "        <LagFormat>7</LagFormat>",
            "      </PredecessorLink>",
        ])
    lines.append("    </Task>")
    return "\n".join(lines)


def generate_xml(tasks: list, project_name: str) -> str:
    summary_ids = find_summary_ids(tasks)
    resource_list, resource_uid = build_resource_pool(tasks)
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Project xmlns="http://schemas.microsoft.com/project">',
        f"  <Name>{escape(project_name)}</Name>",
        f"  <Title>{escape(project_name)}</Title>",
        "  <Manager>IT PM</Manager>",
        "  <StartDate>2025-01-01T08:00:00</StartDate>",
        "  <FinishDate>2035-12-31T17:00:00</FinishDate>",
        "  <ScheduleFromStart>1</ScheduleFromStart>",
        "  <CalendarUID>1</CalendarUID>",
        "  <DefaultStartTime>08:00:00</DefaultStartTime>",
        "  <DefaultFinishTime>17:00:00</DefaultFinishTime>",
        "  <MinutesPerDay>480</MinutesPerDay>",
        "  <MinutesPerWeek>2400</MinutesPerWeek>",
        "  <DaysPerMonth>20</DaysPerMonth>",
        "  <DefaultTaskType>0</DefaultTaskType>",
        "  <DefaultFixedCostAccrual>3</DefaultFixedCostAccrual>",
        "  <CurrencySymbol>€</CurrencySymbol>",
        "  <CurrencyCode>EUR</CurrencyCode>",
        "  <CurrencyDigits>2</CurrencyDigits>",
        "  <Calendars>",
        "    <Calendar>",
        "      <UID>1</UID>",
        "      <Name>Standard</Name>",
        "      <IsBaseCalendar>1</IsBaseCalendar>",
        "      <IsBaselineCalendar>0</IsBaselineCalendar>",
        "      <WeekDays>",
        "        <WeekDay><DayType>1</DayType><DayWorking>0</DayWorking></WeekDay>",
        "        <WeekDay><DayType>7</DayType><DayWorking>0</DayWorking></WeekDay>",
        "      </WeekDays>",
        "    </Calendar>",
        "  </Calendars>",
    ]

    lines.append("  <Tasks>")
    lines.extend([
        "    <Task>",
        "      <UID>0</UID>",
        "      <ID>0</ID>",
        "      <Name>Project summary</Name>",
        "      <Duration>PT0H0M0S</Duration>",
        "      <DurationFormat>7</DurationFormat>",
        "      <Start>2025-01-01T08:00:00</Start>",
        "      <Finish>2035-12-31T17:00:00</Finish>",
        "      <OutlineLevel>0</OutlineLevel>",
        "      <Summary>1</Summary>",
        "      <Milestone>0</Milestone>",
        "      <CalendarUID>-1</CalendarUID>",
        "    </Task>",
    ])

    for task in tasks:
        lines.append(task_to_xml(task, summary_ids))

    lines.append("  </Tasks>")

    lines.append("  <Resources>")
    for r in resource_list:
        uid = resource_uid[r]
        lines.extend([
            "    <Resource>",
            f"      <UID>{uid}</UID>",
            f"      <ID>{uid}</ID>",
            f"      <Name>{escape(r)}</Name>",
            "      <Type>1</Type>",
            "      <IsNull>0</IsNull>",
            "      <CalendarUID>-1</CalendarUID>",
            "      <IsEnterprise>0</IsEnterprise>",
            "    </Resource>",
        ])
    lines.append("  </Resources>")

    lines.append("  <Assignments>")
    assignment_uid = 1
    for task in tasks:
        for r in [x.strip() for x in (task.get("resources") or "").split("+") if x.strip()]:
            if r not in resource_uid:
                continue
            rid = resource_uid[r]
            lines.extend([
                "    <Assignment>",
                f"      <UID>{assignment_uid}</UID>",
                f"      <TaskUID>{task['id']}</TaskUID>",
                f"      <ResourceUID>{rid}</ResourceUID>",
                "      <Units>1</Units>",
                f"      <Work>{iso_duration_from_days(task['days'])}</Work>",
                f"      <Start>{to_xml_datetime(task['start'], task['milestone'])}</Start>",
                f"      <Finish>{to_xml_datetime(task['finish'], task['milestone'])}</Finish>",
                "    </Assignment>",
            ])
            assignment_uid += 1

    lines.append("  </Assignments>")
    lines.append("</Project>")
    return "\n".join(lines)


def generate_risk_assessment_xlsx(cfg: dict, output: Path):
    columns = [
        "Nb", "Sub-project", "Type", "Priority", "Risk category", "Risk/Opportunity description",
        "Effects", "Root Cause", "Probability", "Impact", "Risk Rating", "EMV",
        "Response strategy", "Actions", "Responsible", "Deadline", "Status"
    ]
    sample = [
        [1, "IT Governance", "Risk", "High", "ScR", "Significant schedule slip if TSA is delayed",
         "GoLive delay", "TSA contract late", 4, 5, 20, 100000,
         "Close contract risks", "Do weekly TSA gating", "Program Manager", cfg["tsa_exit_date"], "Open"]
    ]
    df = pd.DataFrame(sample, columns=columns)
    df.to_excel(output, index=False)
    return output


def generate_cost_plan_csv(cfg: dict, output: Path):
    header = ["CATEGORY", "RESOURCE", "TOTAL DAYS", "TOTAL HRS", "HOURLY RATE (EUR)", "TOTAL COST (EUR)"]
    rows = [
        ["--- GOVERNANCE / PMO ---", "", "", "", "", ""],
        ["", "IT PM", 40, 320, 180, 57600],
        ["", "PMO", 20, 160, 150, 24000],
        ["", "SUBTOTAL - Governance / PMO", "", "", "", 81600],
    ]
    with output.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)
    return output


def generate_opl_csv(cfg: dict, output: Path):
    header = [
        "#", "Source", "Date Reported", "Location", "Sub-Workstream", "WP", "Category", "Title",
        "Action/Description/Impact", "Execution Owner", "Priority", "Status", "Due Date", "Comments"
    ]
    rows = [
        [1, "Initial Assessment", datetime.today().strftime("%Y-%m-%d"), "Global", "IT Infrastructure", "1.1",
         "Decision", "WAN lead time risk", "Need 4-6 months lead time for telecom circuits", "Network Lead", "High", "Open", cfg["day1_date"], ""],
    ]
    with output.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)
    return output


def generate_executive_dashboard_html(cfg: dict, output: Path):
    html = f"""<!DOCTYPE html>
<html lang='en'>
<head><meta charset='UTF-8'><title>Executive Dashboard - {escape(cfg['project_name'])}</title></head>
<body>
  <h1>Executive Dashboard: {escape(cfg['project_name'])}</h1>
  <p>Model: {escape(cfg['carveout_model'])}</p>
  <p>Sites: {escape(cfg['num_sites'])}, Users: {escape(cfg['num_users'])}</p>
  <h2>Milestones</h2>
  <ul>
    <li>Deal closing: {cfg['deal_closing_date']}</li>
    <li>Signing: {cfg['signing_date']}</li>
    <li>TSA exit: {cfg['tsa_exit_date']}</li>
    <li>Day 1: {cfg['day1_date']}</li>
    <li>Closure: {cfg['closure_date']}</li>
  </ul>
</body>
</html>"""
    output.write_text(html, encoding='utf-8')
    return output


def _pdf_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


def generate_biweekly_status_pdf(cfg: dict, output: Path):
    # Simple single-page PDF without extra dependencies.
    lines = [
        f"Status Update: {cfg['project_name']}",
        f"Date: {datetime.today().strftime('%Y-%m-%d')}",
        f"Deal closing: {cfg['deal_closing_date']}",
        f"Signing: {cfg['signing_date']}",
        f"TSA exit: {cfg['tsa_exit_date']}",
        f"Day 1: {cfg['day1_date']}",
        f"Closure: {cfg['closure_date']}",
        "\nKey action items:",
        " - Confirm TSA exit criteria",
        " - Execute Day 1 cutover plan",
    ]
    content = "BT /F1 12 Tf 72 740 Td "
    for i, line in enumerate(lines):
        txt = _pdf_escape(line)
        if i > 0:
            content += " T* "
        content += f"({txt}) Tj"
    content += " ET"
    content_bytes = content.encode('latin1')

    # Single-page PDF objects
    xref = []
    body = b""
    def add_obj(objnum, objstr):
        nonlocal body
        xref.append(len(body))
        body += f"{objnum} 0 obj\n".encode('latin1') + objstr + b"\nendobj\n"

    add_obj(1, b"<< /Type /Catalog /Pages 2 0 R >>")
    add_obj(2, b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>")
    add_obj(3, b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>")
    stream_obj = b"<< /Length %d >>\nstream\n" % len(content_bytes) + content_bytes + b"\nendstream"
    add_obj(4, stream_obj)
    add_obj(5, b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")

    xref_start = len(body)
    xref_section = b"xref\n0 %d\n0000000000 65535 f \n" % (len(xref) + 1)
    for offset in xref:
        xref_section += f"{offset:010d} 00000 n \n".encode('latin1')

    trailer = b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(xref) + 1, xref_start)

    output.write_bytes(b"%PDF-1.4\n" + body + xref_section + trailer)
    return output


def main():
    parser = argparse.ArgumentParser(description="Generate carve-out MS Project XML schedule")
    for arg_name, desc in REQUIRED_FIELDS.items():
        parser.add_argument(f"--{arg_name.replace('_', '-')}", help=desc)
    parser.add_argument("--out", default="Carveout_Schedule.xml", help="Output XML file path")
    args = parser.parse_args()

    cfg = {k: getattr(args, k) for k in REQUIRED_FIELDS}
    cfg = prompt_for_required(cfg)
    validate_date_sequence(cfg)

    tasks = build_tasks(cfg)
    xml_payload = generate_xml(tasks, cfg["project_name"])

    out_path = Path(args.out)
    out_path.write_text(xml_payload, encoding="utf-8")

    # Additional deliverables
    risk_output = Path(args.out).with_name("Risk_Assessment.xlsx")
    cost_output = Path(args.out).with_name("Cost_Plan.csv")
    opl_output = Path(args.out).with_name("Open_Points_List.csv")
    dashboard_output = Path(args.out).with_name("Executive_Dashboard.html")
    status_pdf_output = Path(args.out).with_name("Biweekly_Status_Update.pdf")

    generate_risk_assessment_xlsx(cfg, risk_output)
    generate_cost_plan_csv(cfg, cost_output)
    generate_opl_csv(cfg, opl_output)
    generate_executive_dashboard_html(cfg, dashboard_output)
    generate_biweekly_status_pdf(cfg, status_pdf_output)

    print(f"MS Project XML schedule generated: {out_path}")
    print(f"Tasks: {len(tasks)}")
    print(f"Resources: {len(build_resource_pool(tasks)[0])}")
    print(f"Risk register generated: {risk_output}")
    print(f"Cost plan generated: {cost_output}")
    print(f"OPL generated: {opl_output}")
    print(f"Executive dashboard generated: {dashboard_output}")
    print(f"Biweekly status PDF generated: {status_pdf_output}")


if __name__ == "__main__":
    main()
