"""
generate_hamburger_schedule.py
Generates Hamburger_Project_Schedule.csv and Hamburger_Project_Schedule.xml

Project Hamburger: Solar Energy Business Carve-Out (Stand Alone model)
Seller: Robert Bosch GmbH | Buyer: Undisclosed Buyer
14 months: 01 Apr 2026 -> 30 May 2027 | 17 worldwide sites | 2600 IT users
GoLive / Day 1 (D1): 01 Dec 2026  [same event in Bosch standard]
Hypercare: 90 calendar days (3 months) 01 Dec 2026 -> 01 Mar 2027
Quality Gates: QG0 (kick-off) -> QG1 -> QG2 -> QG3 -> QG4 (pre-GoLive) -> QG5 (closure)
Completion (QG5): 30 May 2027 — includes document archiving
Lead: Erik Ho (BD/MIL-ICC)

Run:  "C:/Program Files/px/python.exe" generate_hamburger_schedule.py
Output: Hamburger/Hamburger_Project_Schedule.csv
        Hamburger/Hamburger_Project_Schedule.xml
"""

import csv
import os
import sys
from pathlib import Path

HERE = Path(__file__).parent
OUT_DIR = HERE / "Hamburger"
OUT_DIR.mkdir(exist_ok=True)

CSV_PATH = OUT_DIR / "Hamburger_Project_Schedule.csv"
XML_PATH = OUT_DIR / "Hamburger_Project_Schedule.xml"

# Columns: ID, Outline Level, Name, Duration, Start, Finish,
#          Predecessors, Resource Names, Notes, Milestone
TASKS = [
    # ── PHASE 1: Initialization  01 Apr 2026 – 30 Jun 2026 ──────────────────
    # Bosch standard: QG0 = programme kick-off / initiation gate
    (1,   1, "Phase 1 - Initialization",                                      "65 days",  "04/01/26", "06/30/26", "",              "Erik Ho (BD/MIL-ICC) + Bosch IT",               "",                                                                                                                 "No"),
    (2,   2, "QG0 - Programme Kick-Off Gate",                                 "0 days",   "04/01/26", "04/01/26", "",              "Steering Committee",                            "Programme formally initiated; SteerCo constituted; all workstream leads confirmed; proceed to Phase 1",           "Yes"),
    (3,   2, "1.1 Project Setup and Governance",                               "20 days",  "04/01/26", "04/28/26", "",              "Erik Ho (BD/MIL-ICC)",                          "",                                                                                                                 "No"),
    (4,   3, "Appoint Programme Leads – Seller and Solar Business",            "5 days",   "04/01/26", "04/07/26", "",              "Bosch CIO + Erik Ho (BD/MIL-ICC)",              "",                                                                                                                 "No"),
    (5,   3, "Establish Steering Committee and Governance Model",              "7 days",   "04/01/26", "04/09/26", "4",             "Bosch + Solar Senior Mgmt",                     "",                                                                                                                 "No"),
    (6,   3, "Define RACI, Workstream Leads and Escalation Path",             "7 days",   "04/06/26", "04/14/26", "4",             "Erik Ho (BD/MIL-ICC)",                           "",                                                                                                                 "No"),
    (7,   3, "Set Up PMO Tools, Collaboration Site and Reporting Cadence",    "7 days",   "04/09/26", "04/17/26", "4",             "PMO + Erik Ho (BD/MIL-ICC)",                    "",                                                                                                                 "No"),
    (8,   2, "Signing - Frozen Zone Activated",                                "0 days",   "04/09/26", "04/09/26", "4",             "Bosch + Solar Executive",                       "No Solar IT production changes without SteerCo approval from this date",                                           "Yes"),
    (9,   2, "1.2 Solar IT Landscape Assessment (17 worldwide sites)",         "40 days",  "04/14/26", "06/09/26", "4",             "Bosch IT + Erik Ho (BD/MIL-ICC)",               "17 sites across EMEA, Americas and APAC",                                                                          "No"),
    (10,  3, "Solar IT Inventory Request (Apps / Infra / Devices / Data)",     "10 days",  "04/14/26", "04/27/26", "6",             "Solar IT PM",                                   "2600 users; global footprint",                                                                                     "No"),
    (11,  3, "Location List – 17 Worldwide Sites",                             "10 days",  "04/27/26", "05/12/26", "10",            "Solar IT + Facilities",                         "Sites across EMEA, Americas and APAC to be confirmed per region",                                                   "No"),
    (12,  3, "Legal Entity Identification – Buyer Entity Registration",        "20 days",  "04/27/26", "05/25/26", "10",            "Legal + Finance",                               "Buyer entity legal setup; change-of-control planning; undisclosed buyer engagement",                               "No"),
    (13,  3, "Identify Solar IT Dependencies on Bosch Shared Services",        "20 days",  "05/11/26", "06/09/26", "11",            "Bosch IT Architects",                           "Full TSA scope from Bosch shared services consumed by Solar business",                                              "No"),
    (14,  2, "1.3 Project Charter and Scope Definition",                       "20 days",  "05/18/26", "06/12/26", "9",             "Erik Ho (BD/MIL-ICC)",                          "",                                                                                                                 "No"),
    (15,  3, "Draft Hamburger Project Charter",                                "10 days",  "05/18/26", "05/29/26", "9",             "Erik Ho (BD/MIL-ICC) + Solar IT PM",            "",                                                                                                                 "No"),
    (16,  3, "Define Scope Boundaries – Solar IT Only, Stand Alone",           "10 days",  "05/18/26", "05/29/26", "11",            "Erik Ho (BD/MIL-ICC) + Bosch IT",               "17 sites; 2600 users; Stand Alone model; no Merger Zone",                                                          "No"),
    (17,  3, "Global IT Workstream Kickoff Workshop",                          "3 days",   "05/25/26", "05/27/26", "15",            "All WS Leads",                                  "",                                                                                                                 "No"),
    (18,  3, "Solar IT Standalone Operating Framework Agreement",              "5 days",   "06/01/26", "06/05/26", "15",            "Legal + Finance + Bosch IT",                    "Confirm Stand Alone model; no Merger Zone; direct Bosch IT hand-over to Buyer",                                   "No"),
    (19,  3, "Project Charter Approval",                                       "0 days",   "06/12/26", "06/12/26", "15,16,17,18",   "Steering Committee",                            "",                                                                                                                 "Yes"),
    (20,  2, "QG1 - Initialization Quality Gate",                              "0 days",   "06/30/26", "06/30/26", "19",            "Steering Committee",                            "Charter approved; full inventory confirmed; governance accepted; budget baseline submitted; proceed to Concept",    "Yes"),

    # ── PHASE 2: Concept  01 Jul 2026 – 31 Aug 2026 ─────────────────────────
    (21,  1, "Phase 2 - Concept",                                              "45 days",  "07/01/26", "08/31/26", "20",            "All Workstreams + Erik Ho (BD/MIL-ICC)",        "",                                                                                                                 "No"),
    (22,  2, "2.1 Detailed As-Is Analysis",                                    "35 days",  "07/01/26", "08/17/26", "20",            "Bosch IT + Erik Ho (BD/MIL-ICC)",               "",                                                                                                                 "No"),
    (23,  3, "Application Landscape Inventory – Solar Business",               "20 days",  "07/01/26", "07/28/26", "20",            "Bosch IT Architects + Erik Ho (BD/MIL-ICC)",    "Full app inventory across all 17 sites; flag Bosch shared platform dependencies",                                  "No"),
    (24,  3, "IT Infrastructure As-Is Mapping – 17 Global Sites",              "20 days",  "07/01/26", "07/28/26", "20",            "Bosch Infra Team",                              "Site surveys covering EMEA, Americas, APAC",                                                                       "No"),
    (25,  3, "ERP / SAP Landscape Analysis – Solar Business Scope",            "25 days",  "07/01/26", "08/04/26", "20",            "Bosch ERP / SAP Team",                          "Identify Solar-specific ERP instances and shared SAP mandants",                                                    "No"),
    (26,  3, "Data Ownership and Separation Rules",                            "15 days",  "07/21/26", "08/10/26", "23",            "Erik Ho (BD/MIL-ICC) + Legal",                  "Identify Solar data vs Bosch corporate data; confirm separation approach",                                          "No"),
    (27,  3, "Contract and Licence Inventory",                                 "15 days",  "07/01/26", "07/21/26", "20",            "Bosch Procurement + IT",                        "Change-of-control clause review for all Solar applications and licences",                                           "No"),
    (28,  3, "HR IT Systems Mapping – Solar Business Users",                   "10 days",  "07/01/26", "07/14/26", "20",            "Bosch HR IT",                                   "Global payroll and HR systems for 2600 Solar users",                                                               "No"),
    (29,  2, "2.2 Standalone IT Architecture Design",                          "30 days",  "07/13/26", "08/21/26", "23,24",         "Bosch IT Architects + Erik Ho (BD/MIL-ICC)",    "",                                                                                                                 "No"),
    (30,  3, "Architecture Concept Workshop – Stand Alone Model",              "5 days",   "07/13/26", "07/17/26", "23",            "Bosch + Buyer IT Architects",                   "Confirm Stand Alone model; define Solar IT target state; no Merger Zone",                                          "No"),
    (31,  3, "Network Architecture Design – 17 Global Sites",                  "15 days",  "07/20/26", "08/07/26", "30",            "Bosch Infra + Erik Ho (BD/MIL-ICC)",            "Solar LAN/WAN across EMEA, Americas, APAC; international links required",                                         "No"),
    (32,  3, "Active Directory Design – Solar Standalone Forest",               "12 days",  "07/20/26", "08/04/26", "30",            "Bosch AD Team",                                 "New AD forest for Solar; fully separate from Bosch corporate AD",                                                  "No"),
    (33,  3, "M365 and Azure Tenant Design – Solar Entity",                    "12 days",  "07/20/26", "08/04/26", "30",            "Bosch Azure Team",                              "New M365 tenant for Solar; 2600 mailboxes; Teams, SharePoint, OneDrive in scope",                                  "No"),
    (34,  3, "ERP / SAP Separation Design",                                    "20 days",  "07/20/26", "08/14/26", "25,30",         "Bosch ERP / SAP Team",                          "SAP client copy or greenfield decision; data migration scope for ERP",                                              "No"),
    (35,  3, "Security Architecture Design – Solar Standalone",                "12 days",  "08/10/26", "08/25/26", "31,32",         "Bosch CISO",                                    "Solar security architecture independent of Bosch; ISO 27001 baseline",                                             "No"),
    (36,  2, "2.3 Solar IT Migration Strategy",                                "30 days",  "07/28/26", "08/31/26", "23",            "Erik Ho (BD/MIL-ICC) + All WS Leads",           "",                                                                                                                 "No"),
    (37,  3, "Application Categorisation and Wave Planning",                   "10 days",  "07/28/26", "08/10/26", "23",            "Erik Ho (BD/MIL-ICC) + WS Leads",               "Multi-wave migration required given 17 sites and global scope",                                                    "No"),
    (38,  3, "Data Separation Rules Finalised",                                "10 days",  "08/11/26", "08/24/26", "26",            "Erik Ho (BD/MIL-ICC) + Legal",                  "",                                                                                                                 "No"),
    (39,  3, "TSA Service Catalogue Definition",                               "15 days",  "07/28/26", "08/17/26", "30",            "Bosch IT + Buyer IT",                           "Services Bosch provides to Solar/Buyer under TSA; define exit criteria per service",                                "No"),
    (40,  3, "Solar IT Operating Model (post-TSA)",                            "10 days",  "08/18/26", "08/31/26", "37",            "Bosch IT Architects",                           "Solar standalone operating model after TSA exit; global IT model",                                                 "No"),
    (41,  2, "2.4 Planning and Baseline",                                      "15 days",  "08/13/26", "08/31/26", "29,36",         "Erik Ho (BD/MIL-ICC) + IT PM",                  "",                                                                                                                 "No"),
    (42,  3, "Detailed Project Plan Development",                              "7 days",   "08/13/26", "08/21/26", "29,36",         "Erik Ho (BD/MIL-ICC) + IT PM",                  "",                                                                                                                 "No"),
    (43,  3, "Risk Register Baseline",                                         "7 days",   "08/13/26", "08/21/26", "36",            "Erik Ho (BD/MIL-ICC) + All WS Leads",           "",                                                                                                                 "No"),
    (44,  3, "Resource Plan and Budget Baseline",                              "7 days",   "08/20/26", "08/28/26", "42",            "Finance + IT PM",                               "",                                                                                                                 "No"),
    (45,  2, "QG2 - Concept Quality Gate",                                     "0 days",   "08/31/26", "08/31/26", "41,40,39",      "Steering Committee",                            "Architecture approved; Stand Alone model agreed; TSA catalogue defined; ERP strategy decided; proceed to Development", "Yes"),

    # ── PHASE 3: Development and Build  01 Sep 2026 – 30 Oct 2026 ───────────
    (46,  1, "Phase 3 - Development and Build",                                "45 days",  "09/01/26", "10/30/26", "45",            "All Workstreams",                               "",                                                                                                                 "No"),
    (47,  2, "3.1 Solar IT Infrastructure Build",                              "40 days",  "09/01/26", "10/27/26", "45",            "Bosch Infra + Partners",                        "",                                                                                                                 "No"),
    (48,  3, "Network Setup – 17 Global Sites",                                "25 days",  "09/01/26", "10/05/26", "45",            "Bosch Infra + WAN Provider",                    "Solar LAN/WAN across all 17 sites; MPLS or SD-WAN; international connectivity",                                   "No"),
    (49,  3, "Data Centre and Server Infrastructure Build – Solar",            "20 days",  "09/01/26", "09/28/26", "45",            "Bosch Infra",                                   "Primary and DR DC for Solar; cloud or co-lo per region",                                                           "No"),
    (50,  3, "Active Directory Build – Solar Standalone Forest",               "20 days",  "09/28/26", "10/27/26", "48",            "Bosch AD Team",                                 "Foundation for all 2600 Solar users and applications",                                                             "No"),
    (51,  3, "M365 Tenant Provisioning – Solar (2600 Mailboxes)",              "15 days",  "09/01/26", "09/21/26", "45",            "Bosch Azure Team",                              "",                                                                                                                 "No"),
    (52,  3, "Azure and Cloud Environment Setup – Solar",                      "12 days",  "09/22/26", "10/07/26", "51",            "Bosch Cloud Team",                              "",                                                                                                                 "No"),
    (53,  3, "Security and IAM Platform Setup – Solar Standalone",             "15 days",  "10/08/26", "10/28/26", "50",            "Bosch CISO",                                    "",                                                                                                                 "No"),
    (54,  2, "3.2 ERP / SAP Separation and Build",                             "40 days",  "09/01/26", "10/27/26", "45",            "Bosch ERP / SAP Team + Migration Partner",      "",                                                                                                                 "No"),
    (55,  3, "SAP System Separation – Solar Mandant Copy or Greenfield",       "25 days",  "09/01/26", "10/05/26", "45",            "Bosch ERP / SAP Team",                          "Critical path: SAP separation drives application migration dependencies",                                           "No"),
    (56,  3, "ERP Data Cleansing and Validation",                              "15 days",  "09/14/26", "10/02/26", "55",            "Bosch ERP / SAP Team + Migration Partner",      "",                                                                                                                 "No"),
    (57,  3, "ERP Integration Testing – Solar Standalone",                     "15 days",  "10/05/26", "10/23/26", "56",            "Bosch ERP / SAP Team",                          "",                                                                                                                 "No"),
    (58,  2, "3.3 Application Migration Preparation",                          "35 days",  "09/01/26", "10/16/26", "45",            "Bosch IT + Erik Ho (BD/MIL-ICC)",               "",                                                                                                                 "No"),
    (59,  3, "Application Reclassification and Assignment",                    "7 days",   "09/01/26", "09/09/26", "45",            "Erik Ho (BD/MIL-ICC) + Bosch IT",               "Assign each Solar app to Standalone or retain in Bosch",                                                           "No"),
    (60,  3, "Application Adaptation for Solar Standalone Environment",        "20 days",  "09/10/26", "10/07/26", "59",            "App Teams",                                     "Re-point app configurations from Bosch to Solar domain / AD / M365",                                               "No"),
    (61,  3, "Data Migration Preparation – Solar Business Data",               "15 days",  "09/14/26", "10/02/26", "59",            "Bosch IT + Erik Ho (BD/MIL-ICC)",               "",                                                                                                                 "No"),
    (62,  2, "3.4 Client Device Migration (2600 Devices)",                     "25 days",  "09/28/26", "10/30/26", "50",            "Bosch CWP Team",                                "",                                                                                                                 "No"),
    (63,  3, "Device Inventory and Assessment – 2600 Solar Users",             "7 days",   "09/28/26", "10/07/26", "50",            "Bosch IT + Asset Mgmt",                         "Global device inventory across 17 sites",                                                                          "No"),
    (64,  3, "Device Reimaging and Configuration – Solar Standard Image",      "18 days",  "10/08/26", "10/30/26", "63",            "Bosch CWP Team",                                "Solar domain image for 2600 devices; staged rollout across sites",                                                 "No"),
    (65,  2, "3.5 TSA Framework Finalisation",                                 "35 days",  "09/01/26", "10/16/26", "45",            "Bosch IT + Legal",                              "",                                                                                                                 "No"),
    (66,  3, "TSA Service Descriptions Finalised – Global Scope",              "15 days",  "09/01/26", "09/21/26", "45",            "Bosch IT + Buyer IT",                           "Each TSA service must have an exit date and acceptance criteria",                                                   "No"),
    (67,  3, "TSA SLA and KPI Framework",                                      "12 days",  "09/22/26", "10/07/26", "66",            "IT PM + Legal",                                 "",                                                                                                                 "No"),
    (68,  3, "TSA Contracts Legal Review and Finalisation",                    "15 days",  "09/22/26", "10/12/26", "67",            "Legal",                                         "",                                                                                                                 "No"),
    (69,  2, "QG3 - Development Quality Gate",                                 "0 days",   "10/30/26", "10/30/26", "47,54,58,62,65","Steering Committee",                            "Solar infra ready; all apps packaged; 2600 devices configured; ERP separated; TSA contracts signed; proceed to Testing", "Yes"),

    # ── PHASE 4: Testing and Cutover  02 Nov 2026 – 30 Nov 2026 ─────────────
    (70,  1, "Phase 4 - Testing and Cutover",                                  "22 days",  "11/02/26", "11/30/26", "69",            "All Workstreams",                               "",                                                                                                                 "No"),
    (71,  2, "4.1 System Integration Testing",                                 "15 days",  "11/02/26", "11/20/26", "69",            "Bosch IT + Erik Ho (BD/MIL-ICC)",               "",                                                                                                                 "No"),
    (72,  3, "Integration Testing – All Solar Applications",                   "12 days",  "11/02/26", "11/17/26", "69",            "App Teams + Test Team",                         "",                                                                                                                 "No"),
    (73,  3, "ERP / SAP End-to-End Regression Testing",                        "12 days",  "11/02/26", "11/17/26", "69",            "Bosch ERP / SAP Team",                          "",                                                                                                                 "No"),
    (74,  3, "User Acceptance Testing (UAT) – Global",                         "10 days",  "11/09/26", "11/20/26", "72,73",         "Business Key Users",                            "Solar IT and business users confirm Day 1 / GoLive readiness across all 17 sites",                                "No"),
    (75,  3, "Security and Penetration Testing",                               "7 days",   "11/02/26", "11/10/26", "69",            "Bosch CISO",                                    "Solar standalone environment penetration test; data residency confirmed per region",                               "No"),
    (76,  3, "UAT Sign-Off",                                                   "0 days",   "11/20/26", "11/20/26", "74",            "Business Leads + IT PM",                        "",                                                                                                                 "Yes"),
    (77,  2, "4.2 Cutover Preparation",                                        "15 days",  "11/13/26", "12/01/26", "71",            "IT PM + All WS",                                "",                                                                                                                 "No"),
    (78,  3, "Cutover Plan Finalisation",                                      "7 days",   "11/13/26", "11/21/26", "71",            "IT PM + WS Leads",                              "",                                                                                                                 "No"),
    (79,  3, "End-User Communication – Solar Business (2600 Users)",           "7 days",   "11/16/26", "11/24/26", "78",            "Comms + Bosch IT",                              "Communication to all Solar users across 17 global sites",                                                          "No"),
    (80,  3, "Dress Rehearsal and Cutover Simulation",                         "3 days",   "11/23/26", "11/25/26", "78",            "All WS Leads",                                  "",                                                                                                                 "No"),
    (81,  3, "Help Desk Activation and Training – Solar Global",               "5 days",   "11/24/26", "11/30/26", "78",            "IT Ops",                                        "Global support coverage for Solar users across 17 sites",                                                          "No"),
    (82,  3, "Go No-Go Readiness Assessment",                                  "0 days",   "11/27/26", "11/27/26", "79,80,81",      "Steering Committee",                            "",                                                                                                                 "Yes"),
    (83,  2, "QG4 - Pre-GoLive Quality Gate",                                  "0 days",   "11/30/26", "11/30/26", "82",            "Steering Committee",                            "UAT passed; all systems ready; cutover plan approved; Day 1 / GoLive confirmed for 01 Dec 2026",                  "Yes"),

    # ── PHASE 5: GoLive / Day 1 and Stabilisation  01 Dec 2026 – 01 Mar 2027 ─
    # Day 1 = GoLive = same event in Bosch standard (01 Dec 2026)
    # Hypercare = 90 calendar days (3 months): 01 Dec 2026 -> 01 Mar 2027
    (84,  1, "Phase 5 - GoLive and Stabilisation",                             "65 days",  "12/01/26", "03/01/27", "83",            "All Workstreams",                               "Phase 5 spans Day 1 (GoLive) through 90-day hypercare to 01 Mar 2027",                                            "No"),
    (85,  2, "5.1 Day 1 GoLive Activities",                                    "3 days",   "12/01/26", "12/03/26", "83",            "IT PM + Business",                              "",                                                                                                                 "No"),
    (86,  3, "Day 1 / GoLive – Solar IT Separation Effective (D1)",           "0 days",   "12/01/26", "12/01/26", "83",            "Executive Leadership",                          "Day 1 = GoLive (same event). Solar IT fully separated from Bosch; legal separation effective",                     "Yes"),
    (87,  3, "Network Domain Cutover to Solar Active Directory",               "2 days",   "12/01/26", "12/02/26", "86",            "Bosch Infra",                                   "",                                                                                                                 "No"),
    (88,  3, "Application Day 1 / GoLive Activation – All Solar Apps",        "1 days",   "12/01/26", "12/01/26", "86",            "App Teams",                                     "",                                                                                                                 "No"),
    (89,  3, "Bosch TSA Residual Services Activated",                          "1 days",   "12/01/26", "12/01/26", "86",            "Bosch IT + Buyer IT",                           "TSA clock starts Day 1 (01 Dec 2026); Bosch-to-Solar residual services; TSA exit target May 2027",               "No"),
    (90,  3, "User Go-Live Communication (Solar Business 2600 Users)",         "1 days",   "12/01/26", "12/01/26", "86",            "Comms + Bosch IT",                              "",                                                                                                                 "No"),
    (91,  2, "5.2 Hypercare – 90 Calendar Days (3 Months)",                   "65 days",  "12/01/26", "03/01/27", "85",            "All WS Leads + IT Ops",                         "Hypercare period: Day 1 (01 Dec 2026) to 01 Mar 2027 = 90 calendar days",                                          "No"),
    (92,  3, "Daily Operations Review Meetings",                               "65 days",  "12/01/26", "03/01/27", "85",            "IT PM + WS Leads",                              "Global standup cadence; focus on Solar systems stability during 90-day hypercare",                                 "No"),
    (93,  3, "Incident Management and P1/P2 Issue Resolution",                 "65 days",  "12/01/26", "03/01/27", "85",            "IT Ops + Help Desk",                            "",                                                                                                                 "No"),
    (94,  3, "TSA SLA Compliance Monitoring",                                  "65 days",  "12/01/26", "03/01/27", "85",            "IT Ops",                                        "Bosch residual TSA monitored throughout hypercare; full exit by QG5 (May 2027)",                                  "No"),
    (95,  3, "Hypercare Close – Formal Sign-Off (90 days from Day 1)",        "0 days",   "03/01/27", "03/01/27", "92,93,94",      "Steering Committee",                            "Hypercare complete; 90 calendar days from GoLive / Day 1 (01 Dec 2026)",                                           "Yes"),
    (96,  2, "5.3 TSA Exit and Bosch Service Termination",                     "40 days",  "12/14/26", "02/05/27", "85",            "Bosch IT + Buyer IT",                           "",                                                                                                                 "No"),
    (97,  3, "TSA Exit Planning per Service",                                  "10 days",  "12/14/26", "12/25/26", "85",            "IT PM + Bosch IT",                              "Define exit criteria and acceptance tests per remaining TSA service",                                               "No"),
    (98,  3, "TSA Service Exit – All 17 Sites",                                "20 days",  "01/05/27", "02/01/27", "97",            "Bosch IT + Buyer IT",                           "",                                                                                                                 "No"),
    (99,  3, "TSA Exit Confirmation (All Bosch Services Terminated)",          "0 days",   "02/05/27", "02/05/27", "98",            "Bosch + Buyer Executive",                       "",                                                                                                                 "Yes"),

    # ── PHASE 6: Programme Closure  01 Mar 2027 – 30 May 2027 ───────────────
    # QG5 = Programme Closure quality gate = completion date (includes document archiving)
    (100, 1, "Phase 6 - Programme Closure",                                    "61 days",  "03/01/27", "05/28/27", "95,99",         "IT PM + Erik Ho (BD/MIL-ICC)",                  "",                                                                                                                 "No"),
    (101, 2, "Stabilisation Milestone – Hypercare and TSA Complete",           "0 days",   "03/01/27", "03/01/27", "95,99",         "Steering Committee",                            "Hypercare 90 days complete; TSA fully exited; Solar IT fully independent; proceed to programme closure",           "Yes"),
    (102, 2, "6.1 Closure Activities",                                         "60 days",  "03/01/27", "05/26/27", "101",           "IT PM + All WS Leads",                          "",                                                                                                                 "No"),
    (103, 3, "Lessons Learned Workshop",                                       "3 days",   "03/01/27", "03/03/27", "101",           "All WS Leads",                                  "",                                                                                                                 "No"),
    (104, 3, "Knowledge Transfer to Buyer IT Organisation",                    "20 days",  "03/04/27", "03/31/27", "103",           "Bosch IT + Buyer IT",                           "Formal knowledge transfer for all Solar IT domains",                                                               "No"),
    (105, 3, "IT Contract and Asset Transfer Documentation",                   "20 days",  "03/04/27", "03/31/27", "103",           "Legal + Finance + IT PM",                       "",                                                                                                                 "No"),
    (106, 3, "Document Archiving – All Programme Records",                     "10 days",  "04/07/27", "04/20/27", "104,105",       "IT PM + Legal",                                 "All programme documents archived per Bosch standards; required before QG5 closure gate",                          "No"),
    (107, 3, "Programme Closure Report",                                       "10 days",  "04/21/27", "05/02/27", "106",           "IT PM + Erik Ho (BD/MIL-ICC)",                  "",                                                                                                                 "No"),
    (108, 3, "Final Programme Sign-Off",                                       "0 days",   "05/02/27", "05/02/27", "107",           "Executive Leadership",                          "",                                                                                                                 "Yes"),
    (109, 2, "QG5 - Programme Closure Quality Gate",                           "0 days",   "05/28/27", "05/28/27", "102,108",       "Steering Committee",                            "QG5 = completion gate. All docs archived; KT complete; contracts transferred; Solar IT handed over; Hamburger closed", "Yes"),
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
             "--project", "Project Hamburger"],
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
