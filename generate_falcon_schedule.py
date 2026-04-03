"""
generate_falcon_schedule.py
Generates Falcon_Project_Schedule.csv and Falcon_Project_Schedule.xml

Project Falcon: BGSW AI Business Carve-Out into 50/50 JV (Combination Model)
Seller: Robert Bosch India (BGSW) | Buyer: TBC (not disclosed)
9 months: 01 May 2026 -> 29 Jan 2027 | 3 India sites | <100 users | ~20 apps | No ERP
Lead: Bosch BD/MIL

Run:  python generate_falcon_schedule.py
Output: Falcon\Falcon_Project_Schedule.csv
        Falcon\Falcon_Project_Schedule.xml
"""

import csv
import os
import sys
from pathlib import Path

HERE = Path(__file__).parent
OUT_DIR = HERE / "Falcon"
OUT_DIR.mkdir(exist_ok=True)

CSV_PATH = OUT_DIR / "Falcon_Project_Schedule.csv"
XML_PATH = OUT_DIR / "Falcon_Project_Schedule.xml"

# Columns: ID, Outline Level, Name, Duration, Start, Finish,
#          Predecessors, Resource Names, Notes, Milestone
TASKS = [
    # ── PHASE 1: Initialization  01 May 2026 – 30 Jun 2026 ──────────────────
    (1,  1, "Phase 1 - Initialization",                                   "45 days",  "05/01/26", "06/30/26", "",              "Bosch BD/MIL + BGSW IT",                        "",                                                                                             "No"),
    (2,  2, "1.1 Project Setup and Governance",                           "15 days",  "05/01/26", "05/21/26", "",              "Bosch BD/MIL",                                  "",                                                                                             "No"),
    (3,  3, "Appoint Falcon Programme Leads (BGSW and JV)",               "5 days",   "05/01/26", "05/07/26", "",              "BGSW CIO + Bosch BD/MIL",                       "",                                                                                             "No"),
    (4,  3, "Establish Falcon Steering Committee",                         "5 days",   "05/05/26", "05/11/26", "3",             "BGSW + JV Senior Mgmt",                         "",                                                                                             "No"),
    (5,  3, "Define RACI and Governance Model",                            "7 days",   "05/05/26", "05/13/26", "3",             "Bosch BD/MIL",                                  "",                                                                                             "No"),
    (6,  3, "Set Up PMO Tools and Collaboration Site",                     "7 days",   "05/11/26", "05/19/26", "3",             "PMO + Bosch BD/MIL",                            "",                                                                                             "No"),
    (7,  2, "Signing - Frozen Zone Activated",                             "0 days",   "05/11/26", "05/11/26", "3",             "BGSW + JV Executive",                           "No AI IT production changes without SteerCo approval from this date",                          "Yes"),
    (8,  2, "1.2 BGSW AI IT Landscape Assessment (3 India Sites)",        "25 days",  "05/18/26", "06/19/26", "3",             "BGSW IT + Bosch BD/MIL",                        "3 India sites: Bangalore, Hyderabad, Pune",                                                    "No"),
    (9,  3, "AI IT Inventory Request (Apps / Infra / Devices)",            "7 days",   "05/18/26", "05/26/26", "5",             "BGSW IT PM",                                    "",                                                                                             "No"),
    (10, 3, "Location List - 3 India Sites",                               "7 days",   "05/25/26", "06/02/26", "9",             "BGSW IT + Facilities",                          "Bangalore, Hyderabad, Pune",                                                                   "No"),
    (11, 3, "Legal Entity Identification - JV Registration India",         "15 days",  "05/25/26", "06/12/26", "9",             "Legal + Finance",                               "JV entity registration in India; MCA filing; 50/50 shareholding structure",                    "No"),
    (12, 3, "Identify AI IT Dependencies on BGSW Shared Services",        "15 days",  "06/01/26", "06/19/26", "10",            "BGSW IT Architects",                            "Forms full TSA scope from BGSW shared services used by AI business",                           "No"),
    (13, 2, "1.3 Project Charter and Scope Definition",                    "15 days",  "06/08/26", "06/26/26", "8",             "Bosch BD/MIL",                                  "",                                                                                             "No"),
    (14, 3, "Draft Falcon Project Charter",                                "7 days",   "06/08/26", "06/16/26", "8",             "Bosch BD/MIL + BGSW IT PM",                     "",                                                                                             "No"),
    (15, 3, "Define Scope Boundaries - AI Business IT Only",               "7 days",   "06/08/26", "06/16/26", "10",            "Bosch BD/MIL + BGSW IT",                        "<100 users; ~20 apps; No ERP; Combination JV model; 3 India sites",                           "No"),
    (16, 3, "JV Operating Framework Workshop",                             "3 days",   "06/15/26", "06/17/26", "14",            "All WS Leads",                                  "",                                                                                             "No"),
    (17, 3, "JV IT Structure and Sponsorship Agreement",                   "5 days",   "06/15/26", "06/19/26", "14",            "Legal + Finance + BGSW",                        "Confirm 50/50 JV IT ownership and BGSW-as-carver model",                                       "No"),
    (18, 3, "Project Charter Approval",                                    "0 days",   "06/26/26", "06/26/26", "14,15,16,17",   "Steering Committee",                            "",                                                                                             "Yes"),
    (19, 2, "QG1 - Initialization Quality Gate",                           "0 days",   "06/30/26", "06/30/26", "18",            "Steering Committee",                            "Charter approved; full inventory confirmed; governance accepted; proceed to Concept Phase",     "Yes"),

    # ── PHASE 2: Concept  01 Jul 2026 – 28 Aug 2026 ─────────────────────────
    (20, 1, "Phase 2 - Concept",                                           "45 days",  "07/01/26", "08/28/26", "19",            "All Workstreams + Bosch BD/MIL",                "",                                                                                             "No"),
    (21, 2, "2.1 Detailed As-Is Analysis",                                 "35 days",  "07/01/26", "08/14/26", "19",            "BGSW IT + Bosch BD/MIL",                        "",                                                                                             "No"),
    (22, 3, "Application Landscape Inventory (~20 AI Apps)",               "20 days",  "07/01/26", "07/28/26", "19",            "BGSW IT Architects + Bosch BD/MIL",             "AI-specific apps only; flag BGSW shared platform dependencies",                                "No"),
    (23, 3, "IT Infrastructure As-Is Mapping - 3 India Sites",             "15 days",  "07/01/26", "07/21/26", "19",            "BGSW Infra Team",                               "Site surveys for Bangalore, Hyderabad, Pune",                                                  "No"),
    (24, 3, "Data Ownership and Separation Rules",                         "15 days",  "07/21/26", "08/10/26", "22",            "Bosch BD/MIL + Legal",                          "Identify AI data vs BGSW corporate data; confirm separation approach",                         "No"),
    (25, 3, "Contract and Licence Inventory (~20 AI Apps)",                "15 days",  "07/01/26", "07/21/26", "19",            "BGSW Procurement + IT",                         "Change of control clause review for all ~20 AI applications",                                  "No"),
    (26, 3, "HR IT Systems Mapping - AI Business Users",                   "10 days",  "07/01/26", "07/14/26", "19",            "BGSW HR IT",                                    "India payroll and HR systems in scope for <100 AI users",                                      "No"),
    (27, 2, "2.2 JV IT Architecture Design",                               "30 days",  "07/13/26", "08/21/26", "22,23",         "BGSW IT Architects + Bosch BD/MIL",             "",                                                                                             "No"),
    (28, 3, "JV Architecture Concept Workshop (Combination Model)",        "5 days",   "07/13/26", "07/17/26", "22",            "BGSW + JV Architects",                          "Confirm Combination model; define JV IT target state alongside BGSW carver setup",             "No"),
    (29, 3, "Network Architecture Design - 3 India JV Sites",              "12 days",  "07/20/26", "08/04/26", "28",            "BGSW Infra + Bosch BD/MIL",                     "JV LAN/WAN across India sites; domestic links sufficient; no international WAN required",     "No"),
    (30, 3, "Active Directory Design - JV Forest",                         "12 days",  "07/20/26", "08/04/26", "28",            "BGSW AD Team",                                  "New AD forest for JV entity; fully separate from BGSW corporate AD",                          "No"),
    (31, 3, "M365 and Azure Tenant Design - JV Entity",                    "12 days",  "07/20/26", "08/04/26", "28",            "BGSW Azure Team",                               "New M365 tenant for JV; <100 mailboxes; Teams and OneDrive in scope",                         "No"),
    (32, 3, "Security Architecture Design - JV Entity",                    "12 days",  "08/05/26", "08/20/26", "29,30",         "BGSW CISO",                                     "JV security architecture independent of BGSW; ISO 27001 baseline",                            "No"),
    (33, 2, "2.3 AI IT Migration Strategy",                                "25 days",  "07/28/26", "08/28/26", "22",            "Bosch BD/MIL + All WS Leads",                   "",                                                                                             "No"),
    (34, 3, "Application Categorisation (~20 AI Apps)",                    "7 days",   "07/28/26", "08/05/26", "22",            "Bosch BD/MIL + WS Leads",                       "All ~20 apps in single migration wave given small count",                                      "No"),
    (35, 3, "Data Separation Rules Finalised",                             "10 days",  "08/11/26", "08/24/26", "24",            "Bosch BD/MIL + Legal",                          "",                                                                                             "No"),
    (36, 3, "TSA Service Catalogue Definition",                            "15 days",  "07/28/26", "08/17/26", "28",            "BGSW IT + JV IT",                               "Services BGSW provides to JV under TSA; define exit criteria per service",                    "No"),
    (37, 3, "JV IT Operating Model (post-TSA)",                            "10 days",  "08/12/26", "08/25/26", "34",            "BGSW IT Architects",                            "JV standalone operating model after TSA exit; India-centric model",                           "No"),
    (38, 2, "2.4 Planning and Baseline",                                   "12 days",  "08/13/26", "08/28/26", "27,33",         "Bosch BD/MIL + IT PM",                          "",                                                                                             "No"),
    (39, 3, "Detailed Project Plan Development",                           "7 days",   "08/13/26", "08/21/26", "27,33",         "Bosch BD/MIL + IT PM",                          "",                                                                                             "No"),
    (40, 3, "Risk Register Baseline",                                      "7 days",   "08/13/26", "08/21/26", "33",            "Bosch BD/MIL + All WS Leads",                   "",                                                                                             "No"),
    (41, 3, "Resource Plan and Budget Baseline",                           "7 days",   "08/18/26", "08/26/26", "39",            "Finance + IT PM",                               "",                                                                                             "No"),
    (42, 2, "QG2 - Concept Quality Gate",                                  "0 days",   "08/28/26", "08/28/26", "38,37,36",      "Steering Committee",                            "Architecture approved; JV model agreed; TSA catalogue defined; migration strategy approved; proceed to Development", "Yes"),

    # ── PHASE 3: Development and Build  01 Sep 2026 – 30 Oct 2026 ───────────
    (43, 1, "Phase 3 - Development and Build",                             "45 days",  "09/01/26", "10/30/26", "42",            "All Workstreams",                               "",                                                                                             "No"),
    (44, 2, "3.1 JV IT Infrastructure Build",                              "35 days",  "09/01/26", "10/16/26", "42",            "BGSW Infra + Partners",                         "",                                                                                             "No"),
    (45, 3, "Network Setup - 3 India JV Sites",                            "20 days",  "09/01/26", "09/28/26", "42",            "BGSW Infra",                                    "JV LAN across Bangalore, Hyderabad, Pune; domestic India connectivity",                        "No"),
    (46, 3, "Server Infrastructure Build - JV (Cloud-first)",              "15 days",  "09/01/26", "09/21/26", "42",            "BGSW Infra",                                    "Cloud-first approach; minimal on-premise footprint",                                           "No"),
    (47, 3, "Active Directory Build - JV Forest",                          "20 days",  "09/14/26", "10/09/26", "45",            "BGSW AD Team",                                  "Foundation for all JV users and ~20 AI applications",                                         "No"),
    (48, 3, "M365 Tenant Provisioning - JV (<100 Mailboxes)",              "12 days",  "09/01/26", "09/16/26", "42",            "BGSW Azure Team",                               "",                                                                                             "No"),
    (49, 3, "Azure and Cloud Environment Setup - JV",                      "10 days",  "09/17/26", "09/30/26", "48",            "BGSW Cloud Team",                               "",                                                                                             "No"),
    (50, 3, "Security and IAM Platform Setup - JV",                        "15 days",  "10/01/26", "10/21/26", "47",            "BGSW CISO",                                     "",                                                                                             "No"),
    (51, 2, "3.2 Application Migration Preparation (~20 AI Apps)",         "35 days",  "09/01/26", "10/16/26", "42",            "BGSW IT + Bosch BD/MIL",                        "",                                                                                             "No"),
    (52, 3, "Application Reclassification and Assignment",                 "7 days",   "09/01/26", "09/09/26", "42",            "Bosch BD/MIL + BGSW IT",                        "Assign each of ~20 AI apps to JV or retain in BGSW carver",                                   "No"),
    (53, 3, "Application Adaptation for JV Environment (~20 Apps)",        "20 days",  "09/10/26", "10/07/26", "52",            "Dev Teams",                                     "Re-point app configurations from BGSW to JV domain / AD / M365",                              "No"),
    (54, 3, "Data Migration Preparation - AI Business Data",               "15 days",  "09/14/26", "10/02/26", "52",            "BGSW IT + Bosch BD/MIL",                        "",                                                                                             "No"),
    (55, 2, "3.3 Client Device Migration (<100 Devices)",                  "20 days",  "09/21/26", "10/16/26", "47",            "BGSW CWP Team",                                 "",                                                                                             "No"),
    (56, 3, "Device Inventory and Assessment - AI Users (<100 devices)",   "5 days",   "09/21/26", "09/25/26", "47",            "BGSW IT + Asset Mgmt",                          "",                                                                                             "No"),
    (57, 3, "Device Reimaging and Configuration - JV Standard Image",      "15 days",  "09/28/26", "10/16/26", "56",            "BGSW CWP",                                      "JV domain image for <100 devices across 3 India sites",                                        "No"),
    (58, 2, "3.4 TSA Framework Finalisation",                              "35 days",  "09/01/26", "10/16/26", "42",            "BGSW IT + Legal",                               "",                                                                                             "No"),
    (59, 3, "TSA Service Descriptions Finalised - 3 India Sites",          "15 days",  "09/01/26", "09/21/26", "42",            "BGSW IT + JV IT",                               "Each TSA service must have an exit date and acceptance criteria",                              "No"),
    (60, 3, "TSA SLA and KPI Framework",                                   "12 days",  "09/22/26", "10/07/26", "59",            "IT PM + Legal",                                 "",                                                                                             "No"),
    (61, 3, "TSA Contracts Legal Review and Finalisation",                 "15 days",  "09/22/26", "10/12/26", "60",            "Legal",                                         "",                                                                                             "No"),
    (62, 2, "QG3 - Development Quality Gate",                              "0 days",   "10/30/26", "10/30/26", "44,51,55,58",   "Steering Committee",                            "JV infra ready; all ~20 apps packaged; <100 devices configured; TSA contracts signed; proceed to Cutover", "Yes"),

    # ── PHASE 4: Testing and Cutover  02 Nov 2026 – 04 Dec 2026 ─────────────
    (63, 1, "Phase 4 - Testing and Cutover",                               "25 days",  "11/02/26", "12/04/26", "62",            "All Workstreams",                               "",                                                                                             "No"),
    (64, 2, "4.1 System Integration Testing",                              "15 days",  "11/02/26", "11/20/26", "62",            "BGSW IT + Bosch BD/MIL",                        "",                                                                                             "No"),
    (65, 3, "Integration Testing - All ~20 AI Applications",               "12 days",  "11/02/26", "11/17/26", "62",            "App Teams + Test Team",                         "",                                                                                             "No"),
    (66, 3, "User Acceptance Testing (UAT)",                               "10 days",  "11/09/26", "11/20/26", "65",            "Business Key Users",                            "JV IT and AI business users confirm Day 1 readiness",                                          "No"),
    (67, 3, "Security and Penetration Testing",                            "7 days",   "11/02/26", "11/10/26", "62",            "BGSW CISO",                                     "JV environment penetration test; India data residency confirmed",                               "No"),
    (68, 3, "UAT Sign-Off",                                                "0 days",   "11/20/26", "11/20/26", "66",            "Business Leads + IT PM",                        "",                                                                                             "Yes"),
    (69, 2, "4.2 Cutover Preparation",                                     "12 days",  "11/17/26", "12/02/26", "64",            "IT PM + All WS",                                "",                                                                                             "No"),
    (70, 3, "Cutover Plan Finalisation",                                   "7 days",   "11/17/26", "11/25/26", "64",            "IT PM + WS Leads",                              "",                                                                                             "No"),
    (71, 3, "End-User Communication - AI Business (<100 Users)",           "7 days",   "11/17/26", "11/25/26", "70",            "Comms + BGSW IT",                               "Communication to all AI users across 3 India sites",                                           "No"),
    (72, 3, "Dress Rehearsal and Cutover Simulation",                      "3 days",   "11/26/26", "11/30/26", "70",            "All WS Leads",                                  "",                                                                                             "No"),
    (73, 3, "Help Desk Activation and Training - JV",                      "5 days",   "11/26/26", "12/02/26", "70",            "IT Ops",                                        "India support coverage for JV AI users",                                                       "No"),
    (74, 3, "Go No-Go Readiness Assessment",                               "0 days",   "11/30/26", "11/30/26", "71,72,73",      "Steering Committee",                            "",                                                                                             "Yes"),
    (75, 2, "QG4 - Pre-Cutover Quality Gate",                              "0 days",   "12/04/26", "12/04/26", "74",            "Steering Committee",                            "UAT passed; <100 devices migrated; all AI apps live in JV; JV infra ready; cutover plan approved; Day 1 confirmed", "Yes"),

    # ── PHASE 5: GoLive and Programme Closure  07 Dec 2026 – 29 Jan 2027 ────
    (76, 1, "Phase 5 - GoLive and Programme Closure",                      "40 days",  "12/07/26", "01/29/27", "75",            "All Workstreams",                               "",                                                                                             "No"),
    (77, 2, "5.1 Day 1 GoLive Activities",                                 "3 days",   "12/07/26", "12/09/26", "75",            "IT PM + Business",                              "",                                                                                             "No"),
    (78, 3, "Day 1 - JV IT GoLive (Legal Separation Effective)",           "0 days",   "12/07/26", "12/07/26", "75",            "Executive Leadership",                          "JV AI IT fully separated from BGSW; JV IT independent from this date",                         "Yes"),
    (79, 3, "Network Domain Cutover to JV Active Directory",               "2 days",   "12/07/26", "12/08/26", "78",            "BGSW Infra",                                    "",                                                                                             "No"),
    (80, 3, "Application Day 1 Activation - All ~20 AI Apps",              "1 days",   "12/07/26", "12/07/26", "78",            "App Teams",                                     "",                                                                                             "No"),
    (81, 3, "BGSW TSA Residual Services Activated",                        "1 days",   "12/07/26", "12/07/26", "78",            "BGSW IT + JV IT",                               "TSA clock starts 07 Dec 2026; minimal scope BGSW-to-JV services; exit target Mar 2027",       "No"),
    (82, 3, "User Go-Live Communication (AI Business <100 Users)",         "1 days",   "12/07/26", "12/07/26", "78",            "Comms + BGSW IT",                               "",                                                                                             "No"),
    (83, 2, "5.2 Hypercare - 30 Working Days",                             "30 days",  "12/07/26", "01/18/27", "77",            "All WS Leads + IT Ops",                         "",                                                                                             "No"),
    (84, 3, "Daily Operations Review Meetings",                            "30 days",  "12/07/26", "01/18/27", "77",            "IT PM + WS Leads",                              "India standup cadence; focus on JV AI systems stability",                                      "No"),
    (85, 3, "Incident Management and P1/P2 Issue Resolution",              "30 days",  "12/07/26", "01/18/27", "77",            "IT Ops + Help Desk",                            "",                                                                                             "No"),
    (86, 3, "TSA SLA Compliance Monitoring",                               "30 days",  "12/07/26", "01/18/27", "77",            "IT Ops",                                        "BGSW residual TSA tracked from Day 1; planned full exit by Mar 2027",                         "No"),
    (87, 3, "Hypercare Close - Formal Sign-Off",                           "0 days",   "01/18/27", "01/18/27", "84,85,86",      "Steering Committee",                            "",                                                                                             "Yes"),
    (88, 2, "5.3 TSA Exit and BGSW Service Termination",                   "25 days",  "12/14/26", "01/16/27", "77",            "BGSW IT + JV IT",                               "",                                                                                             "No"),
    (89, 3, "TSA Exit Planning per Service",                               "10 days",  "12/14/26", "12/25/26", "77",            "IT PM + BGSW IT",                               "Define exit criteria and acceptance tests per remaining TSA service",                          "No"),
    (90, 3, "TSA Service Exit - All 3 India Sites",                        "15 days",  "01/05/27", "01/23/27", "89",            "BGSW IT + JV IT",                               "",                                                                                             "No"),
    (91, 3, "TSA Exit Confirmation (All BGSW Services Terminated)",        "0 days",   "01/23/27", "01/23/27", "90",            "BGSW + JV Executive",                           "",                                                                                             "Yes"),
    (92, 2, "5.4 Programme Closure",                                       "10 days",  "01/18/27", "01/29/27", "87",            "IT PM + Bosch BD/MIL",                          "",                                                                                             "No"),
    (93, 3, "Lessons Learned Workshop",                                    "3 days",   "01/18/27", "01/20/27", "87",            "All WS Leads",                                  "",                                                                                             "No"),
    (94, 3, "Programme Closure Report",                                    "7 days",   "01/21/27", "01/29/27", "93",            "IT PM + Bosch BD/MIL",                          "",                                                                                             "No"),
    (95, 3, "Final Programme Sign-Off",                                    "0 days",   "01/29/27", "01/29/27", "94",            "Executive Leadership",                          "",                                                                                             "Yes"),
    (96, 2, "QG5 - Programme Closure Quality Gate",                        "0 days",   "01/29/27", "01/29/27", "87,91,95",      "Steering Committee",                            "Hypercare stable; TSA fully exited; JV IT fully independent; Falcon programme formally closed", "Yes"),
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
             "--project", "Project Falcon"],
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
