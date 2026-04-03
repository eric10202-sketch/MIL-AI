"""
generate_bravo_schedule.py
Generates Bravo_Project_Schedule.csv and Bravo_Project_Schedule.xml

Project Bravo: BGSW AI Business Carve-Out into 50/50 JV with Tata
Seller: Bosch BGSW | Buyer: Tata
7 months: 01 Apr 2026 -> 30 Oct 2026 | 2 India sites | 70 users | 17 apps | No ERP | No TSA
PM: Riyaz Ahmed Syed Ahmed (BD/MIL-PSM4)
Bosch leadership control on JV; not antitrust relevant

Run:  python generate_bravo_schedule.py
Output: Bravo\Bravo_Project_Schedule.csv
        Bravo\Bravo_Project_Schedule.xml
"""

import csv
import os
import sys
from pathlib import Path

HERE = Path(__file__).parent
OUT_DIR = HERE / "Bravo"
OUT_DIR.mkdir(exist_ok=True)

CSV_PATH = OUT_DIR / "Bravo_Project_Schedule.csv"
XML_PATH = OUT_DIR / "Bravo_Project_Schedule.xml"

# Columns: ID, Outline Level, Name, Duration, Start, Finish,
#          Predecessors, Resource Names, Notes, Milestone
#
# Bosch Framework: exactly 5 phases (Phase 0–4). No Phase 5 or 6.
# QG structure for Bravo (Activity — Quality Manager combined QG1/2/3):
#   QG0  → end of Phase 0 (Initialization)
#   QG1/2/3 → end of Phase 3 (Build & Test)   [combined by QM decision]
#   QG4  → Pre-GoLive gate — MUST be PASSED BEFORE GoLive Day 1
#   GoLive Day 1 → AFTER QG4 is passed
#   QG5  → Programme Closure gate
TASKS = [
    # ── PHASE 0: Initialization  01 Apr 2026 – 30 Apr 2026 ──────────────────
    (1,  1, "Phase 0 - Initialization",                                         "22 days", "04/01/26", "04/30/26", "",          "Riyaz Ahmed + BGSW IT",            "",                                                                                                                 "No"),
    (2,  2, "0.1 Project Setup and Governance",                                 "8 days",  "04/01/26", "04/10/26", "",          "Riyaz Ahmed",                      "",                                                                                                                 "No"),
    (3,  3, "Appoint Bravo Programme Leads (BGSW and JV-Tata)",                 "3 days",  "04/01/26", "04/03/26", "",          "BGSW CIO + Riyaz Ahmed",           "",                                                                                                                 "No"),
    (4,  3, "Establish Bravo Steering Committee",                               "3 days",  "04/01/26", "04/03/26", "3",         "BGSW + Tata Senior Mgmt",          "",                                                                                                                 "No"),
    (5,  3, "Define RACI and Governance Model",                                 "5 days",  "04/06/26", "04/10/26", "3",         "Riyaz Ahmed",                      "",                                                                                                                 "No"),
    (6,  3, "Set Up PMO Tools and Collaboration Site",                          "5 days",  "04/06/26", "04/10/26", "4",         "PMO + Riyaz Ahmed",                "",                                                                                                                 "No"),
    (7,  2, "Signing - Frozen Zone Activated",                                  "0 days",  "04/07/26", "04/07/26", "3",         "BGSW + Tata Executive",            "No AI IT production changes without SteerCo approval from this date",                                              "Yes"),
    (8,  2, "0.2 BGSW AI IT Landscape Assessment (2 India Sites)",              "15 days", "04/06/26", "04/24/26", "3",         "BGSW IT + Riyaz Ahmed",            "2 India sites",                                                                                                    "No"),
    (9,  3, "AI IT Inventory (17 Apps / 70 Users / 2 Sites)",                   "5 days",  "04/06/26", "04/10/26", "5",         "BGSW IT PM",                       "17 AI apps; 70 users; 2 India sites; no ERP; no SAP",                                                             "No"),
    (10, 3, "Location Confirmation - 2 India Sites",                            "5 days",  "04/06/26", "04/10/26", "5",         "BGSW IT + Facilities",             "",                                                                                                                 "No"),
    (11, 3, "JV Legal Entity Preparation (India MCA Registration)",             "10 days", "04/13/26", "04/24/26", "9",         "Legal + Finance",                  "50/50 Tata-Bosch JV; Bosch leadership control; not antitrust relevant; no regulatory filing required",             "No"),
    (12, 3, "Identify AI IT Dependencies on BGSW Shared Services",             "10 days", "04/13/26", "04/24/26", "9",         "BGSW IT Architects",               "Minimal dependencies; Bosch-led JV remains closely aligned with BGSW operations post-GoLive",                     "No"),
    (13, 2, "0.3 Scope Definition and Charter",                                 "10 days", "04/13/26", "04/24/26", "8",         "Riyaz Ahmed",                      "",                                                                                                                 "No"),
    (14, 3, "Define Scope - 17 AI Apps, 70 Users, 2 India Sites",              "5 days",  "04/13/26", "04/17/26", "8",         "Riyaz Ahmed + BGSW IT",            "No ERP; No SAP; No TSA post-GoLive; Bosch-led JV; Combination model",                                             "No"),
    (15, 3, "JV Operating Framework Confirmation",                              "5 days",  "04/13/26", "04/17/26", "8",         "Riyaz Ahmed + Legal",              "Bosch leadership control; 50/50 JV with Tata; not antitrust relevant; JV continues under Bosch governance",        "No"),
    (16, 3, "Project Charter Draft and Approval",                               "5 days",  "04/20/26", "04/24/26", "14,15",     "Riyaz Ahmed",                      "",                                                                                                                 "No"),
    (17, 2, "QG0 - Initialization Quality Gate",                                "0 days",  "04/30/26", "04/30/26", "12,16",     "Steering Committee",               "Charter approved; inventory confirmed (17 apps, 70 users, 2 India sites); JV governance agreed; proceed to Concept", "Yes"),

    # ── PHASE 1: Concept  01 May 2026 – 15 May 2026 ──────────────────────────
    (18, 1, "Phase 1 - Concept",                                                "11 days", "05/01/26", "05/15/26", "17",        "BGSW IT + Riyaz Ahmed",            "",                                                                                                                 "No"),
    (19, 2, "1.1 As-Is Analysis",                                               "11 days", "05/01/26", "05/15/26", "17",        "BGSW IT + Riyaz Ahmed",            "",                                                                                                                 "No"),
    (20, 3, "Application Deep-Dive - 17 AI Apps",                               "8 days",  "05/01/26", "05/12/26", "17",        "BGSW IT Architects",               "Classify all 17 apps: retain in JV / retire / re-host; no ERP or SAP in scope",                                   "No"),
    (21, 3, "IT Infrastructure Mapping - 2 India Sites",                        "5 days",  "05/01/26", "05/07/26", "17",        "BGSW Infra Team",                  "",                                                                                                                 "No"),
    (22, 3, "Data Ownership and Separation Rules",                              "5 days",  "05/04/26", "05/08/26", "17",        "Riyaz Ahmed + Legal",              "AI data vs BGSW corporate data; simplified by Bosch JV leadership",                                               "No"),
    (23, 3, "Contract and Licence Review - 17 AI Apps",                        "8 days",  "05/01/26", "05/12/26", "17",        "BGSW Procurement + IT",            "Change-of-control clause review for all 17 AI applications",                                                       "No"),
    (24, 3, "HR IT Mapping - 70 AI Business Users",                             "5 days",  "05/01/26", "05/07/26", "17",        "BGSW HR IT",                       "India payroll and HR for 70 users across 2 India sites",                                                           "No"),

    # ── PHASE 2: Architecture and Design  11 May 2026 – 29 May 2026 ──────────
    (25, 1, "Phase 2 - Architecture and Design",                                "15 days", "05/11/26", "05/29/26", "20",        "BGSW IT Architects + Riyaz Ahmed", "",                                                                                                                 "No"),
    (26, 2, "2.1 JV IT Architecture Design",                                    "15 days", "05/11/26", "05/29/26", "20",        "BGSW IT Architects + Riyaz Ahmed", "",                                                                                                                 "No"),
    (27, 3, "JV Architecture Workshop (Combination Model)",                     "3 days",  "05/11/26", "05/13/26", "20",        "BGSW + Tata Architects",           "Confirms Combination JV model; Bosch leadership; JV IT target state defined",                                      "No"),
    (28, 3, "M365 and Azure Tenant Design - JV Entity (<70 Mailboxes)",         "7 days",  "05/14/26", "05/22/26", "27",        "BGSW Azure Team",                  "New M365 tenant; <70 mailboxes; Teams and OneDrive in scope",                                                      "No"),
    (29, 3, "Active Directory Design - JV Forest",                              "7 days",  "05/14/26", "05/22/26", "27",        "BGSW AD Team",                     "New AD forest for JV; fully separate from BGSW corporate AD",                                                     "No"),
    (30, 3, "Network Architecture - 2 India JV Sites",                          "7 days",  "05/14/26", "05/22/26", "27",        "BGSW Infra",                       "Domestic India connectivity across 2 JV sites; no international WAN required",                                     "No"),
    (31, 3, "Security Architecture - JV Entity",                                "5 days",  "05/25/26", "05/29/26", "29,30",     "BGSW CISO",                        "JV security independent of BGSW; ISO 27001 baseline; Bosch security standards retained",                          "No"),
    (32, 2, "2.2 Migration Strategy and Planning Baseline",                     "15 days", "05/11/26", "05/29/26", "20",        "Riyaz Ahmed + WS Leads",           "",                                                                                                                 "No"),
    (33, 3, "Application Migration Plan - 17 Apps (Single Wave)",               "7 days",  "05/11/26", "05/19/26", "20",        "Riyaz Ahmed + WS Leads",           "Single wave; all 17 apps; Bosch-led JV simplifies target endpoint configuration",                                  "No"),
    (34, 3, "Data Migration Plan",                                              "7 days",  "05/11/26", "05/19/26", "22",        "Riyaz Ahmed + BGSW IT",            "",                                                                                                                 "No"),
    (35, 3, "Detailed Project Plan and Budget Baseline",                        "5 days",  "05/25/26", "05/29/26", "33",        "Riyaz Ahmed + Finance",            "",                                                                                                                 "No"),
    (36, 3, "Risk Register Baseline",                                           "5 days",  "05/25/26", "05/29/26", "33",        "Riyaz Ahmed + WS Leads",           "",                                                                                                                 "No"),

    # ── PHASE 3: Build and Test  01 Jun 2026 – 26 Jun 2026 ───────────────────
    (37, 1, "Phase 3 - Build and Test",                                         "20 days", "06/01/26", "06/26/26", "35,36",     "All Workstreams",                  "",                                                                                                                 "No"),
    (38, 2, "3.1 JV IT Infrastructure Build",                                   "18 days", "06/01/26", "06/24/26", "36",        "BGSW Infra + Partners",            "",                                                                                                                 "No"),
    (39, 3, "M365 Tenant Provisioning (<70 Mailboxes)",                         "5 days",  "06/01/26", "06/05/26", "36",        "BGSW Azure Team",                  "",                                                                                                                 "No"),
    (40, 3, "Azure and Cloud Environment Setup - JV",                           "7 days",  "06/08/26", "06/16/26", "39",        "BGSW Cloud Team",                  "",                                                                                                                 "No"),
    (41, 3, "Active Directory Build - JV Forest",                               "10 days", "06/01/26", "06/12/26", "36",        "BGSW AD Team",                     "Foundation for all 70 JV users and 17 AI applications; no backout possible",                                      "No"),
    (42, 3, "Network Setup - 2 India JV Sites",                                 "10 days", "06/01/26", "06/12/26", "36",        "BGSW Infra",                       "",                                                                                                                 "No"),
    (43, 3, "Security and IAM Platform Setup - JV",                             "8 days",  "06/11/26", "06/22/26", "41,42",     "BGSW CISO",                        "",                                                                                                                 "No"),
    (44, 2, "3.2 Application Migration - 17 AI Apps (Single Wave)",             "18 days", "06/01/26", "06/24/26", "36",        "Dev Teams + BGSW IT",              "",                                                                                                                 "No"),
    (45, 3, "Application Reconfiguration for JV - 17 Apps",                    "12 days", "06/01/26", "06/16/26", "36",        "Dev Teams",                        "Re-point 17 apps from BGSW to JV AD / M365 / Azure; Bosch-led JV simplifies target configuration",                "No"),
    (46, 3, "Data Migration Execution",                                         "8 days",  "06/11/26", "06/22/26", "45",        "BGSW IT + Riyaz Ahmed",            "",                                                                                                                 "No"),
    (47, 2, "3.3 Device Migration - 70 Devices (2 India Sites)",                "10 days", "06/15/26", "06/26/26", "41",        "BGSW CWP Team",                    "",                                                                                                                 "No"),
    (48, 3, "Device Inventory and Assessment - 70 Devices",                     "3 days",  "06/15/26", "06/17/26", "41",        "BGSW IT + Asset Mgmt",             "",                                                                                                                 "No"),
    (49, 3, "Device Reimaging - JV Standard Image (70 Devices)",                "7 days",  "06/18/26", "06/26/26", "48",        "BGSW CWP",                         "JV domain image across 2 India sites",                                                                            "No"),
    (50, 2, "3.4 Testing and Validation",                                       "10 days", "06/15/26", "06/26/26", "44",        "Test Team + Business",             "",                                                                                                                 "No"),
    (51, 3, "Integration Testing - All 17 AI Applications",                     "7 days",  "06/15/26", "06/23/26", "45",        "App Teams + Test Team",            "",                                                                                                                 "No"),
    (52, 3, "User Acceptance Testing (UAT) - 70 Users",                        "5 days",  "06/18/26", "06/24/26", "51",        "Business Key Users",               "JV AI business users confirm Day 1 readiness across 2 India sites",                                               "No"),
    (53, 3, "Security and Penetration Testing - JV Environment",                "5 days",  "06/15/26", "06/19/26", "43",        "BGSW CISO",                        "",                                                                                                                 "No"),
    (54, 2, "UAT Sign-Off",                                                     "0 days",  "06/24/26", "06/24/26", "52",        "Business Leads + Riyaz Ahmed",     "",                                                                                                                 "Yes"),
    (55, 2, "QG1/2/3 - Combined Quality Gate (Concept + Architecture + Build)", "0 days",  "06/26/26", "06/26/26", "49,53,54",  "Steering Committee",               "Architecture approved; all 17 apps reconfigured and tested; 70 devices configured; UAT passed; proceed to GoLive preparation", "Yes"),

    # ── PHASE 4: GoLive and Closure  29 Jun 2026 – 30 Oct 2026 ──────────────
    # NOTE: QG4 (Pre-GoLive gate) comes FIRST — GoLive Day 1 only after QG4 passed
    (56, 1, "Phase 4 - GoLive and Closure",                                     "89 days", "06/29/26", "10/30/26", "55",        "All Workstreams",                  "",                                                                                                                 "No"),
    (57, 2, "4.1 Pre-GoLive Preparation",                                       "2 days",  "06/29/26", "06/30/26", "55",        "Riyaz Ahmed + WS Leads",           "",                                                                                                                 "No"),
    (58, 3, "Cutover Plan Finalisation and Go No-Go Assessment",                "1 days",  "06/29/26", "06/29/26", "55",        "All WS Leads + Riyaz Ahmed",       "",                                                                                                                 "No"),
    (59, 3, "Help Desk Activation and Training - JV",                           "2 days",  "06/29/26", "06/30/26", "55",        "IT Ops",                           "India support coverage for 70 JV AI users across 2 sites ready before GoLive",                                    "No"),
    (60, 3, "End-User Pre-GoLive Communication - 70 AI Business Users",         "1 days",  "06/29/26", "06/29/26", "55",        "Comms + BGSW IT",                  "Pre-GoLive communication to all 70 AI users across 2 India sites",                                               "No"),
    (61, 2, "QG4 - Pre-GoLive Quality Gate",                                    "0 days",  "06/30/26", "06/30/26", "58,59,60",  "Steering Committee",               "Cutover plan approved; all 17 apps ready; 70 users and devices ready; JV infra stable; DAY 1 GOLIVE AUTHORISED",   "Yes"),
    (62, 2, "4.2 Day 1 GoLive",                                                 "1 days",  "07/01/26", "07/01/26", "61",        "All Workstreams",                  "GoLive proceeds only after QG4 passed on 30 Jun 2026",                                                            "No"),
    (63, 3, "Day 1 - JV IT GoLive (Legal Separation Effective)",                "0 days",  "07/01/26", "07/01/26", "61",        "Executive Leadership",             "Bosch-led JV AI IT fully separated from BGSW; JV IT independent from this date; no TSA required",                 "Yes"),
    (64, 3, "Network Domain Cutover to JV Active Directory",                    "1 days",  "07/01/26", "07/01/26", "63",        "BGSW Infra",                       "",                                                                                                                 "No"),
    (65, 3, "Application Day 1 Activation - All 17 AI Apps",                   "1 days",  "07/01/26", "07/01/26", "63",        "App Teams",                        "",                                                                                                                 "No"),
    (66, 3, "User Go-Live Communication - 70 AI Business Users",                "1 days",  "07/01/26", "07/01/26", "63",        "Comms + BGSW IT",                  "",                                                                                                                 "No"),
    (67, 2, "4.3 Hypercare - 60 Working Days",                                  "60 days", "07/06/26", "09/25/26", "62",        "WS Leads + IT Ops",                "",                                                                                                                 "No"),
    (68, 3, "Daily Operations Review Meetings",                                 "60 days", "07/06/26", "09/25/26", "62",        "Riyaz Ahmed + WS Leads",           "India standup cadence; JV AI systems stability; Bosch-led escalation path available",                             "No"),
    (69, 3, "Incident Management and P1/P2 Resolution",                         "60 days", "07/06/26", "09/25/26", "62",        "IT Ops + Help Desk",               "",                                                                                                                 "No"),
    (70, 3, "Application Stability Monitoring - 17 Apps",                       "60 days", "07/06/26", "09/25/26", "62",        "App Teams + IT Ops",               "No TSA; JV Bosch-led; escalation via normal BGSW governance channels",                                            "No"),
    (71, 2, "Hypercare Close - Formal Sign-Off",                                "0 days",  "09/25/26", "09/25/26", "68,69,70",  "Steering Committee",               "",                                                                                                                 "Yes"),
    (72, 2, "4.4 Programme Closure",                                            "25 days", "09/28/26", "10/30/26", "71",        "Riyaz Ahmed + Bosch BD/MIL",       "",                                                                                                                 "No"),
    (73, 3, "Lessons Learned Workshop",                                         "3 days",  "09/28/26", "09/30/26", "71",        "All WS Leads",                     "",                                                                                                                 "No"),
    (74, 3, "Final Financial Reconciliation",                                   "5 days",  "09/28/26", "10/02/26", "71",        "Finance + Riyaz Ahmed",            "",                                                                                                                 "No"),
    (75, 3, "Programme Closure Report",                                         "10 days", "10/05/26", "10/16/26", "73",        "Riyaz Ahmed + Bosch BD/MIL",       "",                                                                                                                 "No"),
    (76, 3, "Final Programme Sign-Off",                                         "0 days",  "10/16/26", "10/16/26", "75,74",     "Executive Leadership",             "",                                                                                                                 "Yes"),
    (77, 2, "QG5 - Programme Closure Quality Gate",                             "0 days",  "10/30/26", "10/30/26", "76",        "Steering Committee",               "Hypercare stable; JV IT fully operational under Bosch governance; all 17 apps stable; Project Bravo formally closed", "Yes"),
]

HEADERS = ["ID", "Outline Level", "Name", "Duration", "Start", "Finish",
           "Predecessors", "Resource Names", "Notes", "Milestone"]


def write_csv(path: Path):
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS)
        for t in TASKS:
            writer.writerow(list(t))
    print(f"CSV written: {path}  ({len(TASKS)} tasks)")


if __name__ == "__main__":
    from datetime import datetime as _dt
    _t0 = _dt.now()
    print(f"Started : {_t0.strftime('%Y-%m-%d %H:%M:%S')}")
    try:
        write_csv(CSV_PATH)

        import subprocess
        result = subprocess.run(
            [sys.executable,
             str(HERE / "generate_msp_xml.py"),
             "--csv", str(CSV_PATH),
             "--out", str(XML_PATH),
             "--project", "Project Bravo"],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            print(result.stdout.strip())
            print(f"\nXML written: {XML_PATH}")
        else:
            print("generate_msp_xml.py error:")
            print(result.stderr)
            sys.exit(1)
    finally:
        _t1 = _dt.now()
        print(f"Finished: {_t1.strftime('%Y-%m-%d %H:%M:%S')}  ({(_t1-_t0).total_seconds():.1f}s elapsed)")
