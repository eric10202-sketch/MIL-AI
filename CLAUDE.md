# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Purpose

This repository is a **generic knowledge base and toolkit for IT carve-out project document generation**. It contains reference material from completed real-world M&A IT carve-outs plus methodology guides and generator scripts. Use it to generate standardised deliverables for any new carve-out engagement.

---

## ⚠️ CRITICAL: Reference Projects vs Active Engagements

### Reference Projects (Methodology Use Only)
Project folders and historical files in this repository (e.g. FRAME documents, completed project CSVs) are **reference only**. Use them solely for:
- Risk register patterns and probability/impact benchmarks
- Schedule structure and phase sequencing
- Workstream templates and dependency patterns
- Timeline benchmarks (WAN lead times, AD build, SAP migration, etc.)

**Never copy reference project facts** (parties, scope, dates, organisation names) into deliverables for a new engagement.

### Active Engagement — Always Confirm Before Generating
When the user describes a new project, confirm the buyer and seller before generating any deliverable. Every document — schedules, dashboards, charters, risk registers — must reflect the correct parties for that engagement only.

---

## Mandatory Inputs — Ask Before Generating Any Deliverable

Before generating any project deliverable (charter, schedule, risk register, dashboard, etc.), you **must** confirm the following with the user if not already provided. These fields determine the entire business relationship and must be correct in every document.

| Field | Question to ask | Why it matters |
|-------|----------------|----------------|
| **Project name** | "What is the project name or code name?" | Used in all file names and document headers |
| **Seller** | "Which company is selling / divesting the business?" | Seller currently owns and operates the IT; provides TSA services post-close |
| **Buyer** | "Which company is acquiring / receiving the business?" | Buyer is the sponsor customer; the carved-out IT will be integrated into the buyer's environment |
| **Business being carved out** | "Which business unit or division is being separated?" | Scopes all workstreams and inventory |
| **Carve-out model** | "Stand Alone, Integration with Buyer, or Combination?" | Determines architecture, Merger Zone, and TSA design |
| **PMO / methodology lead** | "Who is leading the PMO?" | Appears in governance section of all documents |
| **Number of worldwide sites** | "How many sites / locations are in scope for the carve-out?" | Must reflect actual carved-out entity sites only — not the seller's total global footprint |
| **Number of IT users** | "How many IT users (employees / contractors) are in scope?" | Must reflect the headcount of the carved-out business unit only — not the seller's total workforce |

> ⚠️ **Do NOT use estimates, benchmarks, or reference project figures for site count or user count.** These numbers must always come directly from the user or the engagement's confirmed scope documentation. They drive dashboard stats, schedule sizing, device wave planning, cost estimates, and resource loading — using wrong figures will invalidate all generated deliverables.

### Rules derived from buyer / seller
- **Sponsor Customer** = Buyer (benefits from the carve-out)
- **Sponsor Contractor** = Seller (owns and operates IT during TSA)
- **IT flow direction** = Seller's IT landscape → [Merger Zone if Integration model] → Buyer's IT environment
- **TSA** = Seller continues operating IT for the carved-out business until Buyer's infrastructure is ready
- For Stand Alone model: there is no Merger Zone; carved-out entity operates independently

---

## Project Schedule Inputs (for automated schedule generation)

The generator script uses Python and requires core project input fields. If any are missing, it prompts interactively.

- PROJECT_NAME
- DEAL_CLOSING_DATE (YYYY-MM-DD)
- SIGNING_DATE (YYYY-MM-DD)
- TSA_EXIT_DATE (YYYY-MM-DD)
- DAY1_DATE (YYYY-MM-DD)
- CLOSURE_DATE (YYYY-MM-DD)
- NUM_SITES
- NUM_USERS
- CARVEOUT_MODEL (Stand Alone / Integration / Combination)
- REGIONS (AP / AM / EMEA)
- ERP_COMPLEXITY (low / medium / high; Brazil / China / Mexico / India)

### Required milestone sequence

1. Deal closing / Programme kick-off
2. Signing / Frozen Zone activation
3. Quality Gates QG1 → QG5
4. GoLive / Day 1
5. Programme closure

## Repository File Reference

### Reference Material (methodology use only)

| File | Purpose |
|------|---------|
| `20240718_FRAME_IT Risk Assessment_completed.csv` | Historical risk register — risk categories, probability/impact ratings, mitigation patterns, full lifecycle |
| `FRAME_IT_OPL_vFINAL.csv` | Historical Open Points List — 220+ items across 8 IT sub-workstreams; dependency and issue patterns |
| `IT INFRA TIMELINE FRAME CURENT.pptx` | IT infrastructure milestone timeline — phase sequencing, critical dependencies, realistic durations |
| `Carveout best practices guide and template.docx` | IT carve-out methodology — carve-out models, guiding principles, full workstream templates |
| `Trinity_Project_Schedule.csv` | Reference schedule — 191 tasks across 6 phases; use as structural model for new project schedules |
| `Trinity_Project_Cost_Plan.csv` | Reference cost plan (Integration model, 6 phases) — structural reference for large engagements |
| `Risk_analysis_template.xlsx` | Blank risk register template for new engagements |
| `Project_charter_template.pptx` | Blank project charter template for new engagements |

### Generator Scripts

| File | Purpose |
|------|---------|
| `generate_alphax_schedule.py` | **Canonical project schedule generator** — copy and adapt for each new engagement; produces both CSV and MS Project XML |
| `generate_msp_xml.py` | Core XML converter — called by project schedule generators; do not modify |
| `generate_carveout_schedule_direct_xml.py` | Alternative generator — produces XML directly from CLI parameters without a task list CSV |
| `generate_risk_register.py` | Risk register generator — adapt for new engagements |
| `generate_project_charter.py` | Project charter generator — adapt for new engagements |
| `generate_alphax_monthly_report.py` | **Canonical monthly status report generator** — copy and adapt for each new engagement; produces a single-page A4 PDF; designed for Power Automate scheduled monthly runs |

### Project Folders

| Folder | Contents |
|--------|---------|
| `AlphaX/` | **Active reference engagement** — all 7 standard deliverables: schedule CSV+XML, cost plan, risk register, project charter, executive dashboard, KPI dashboard, monthly status report PDF + generator |

## ⚠️ New Project Setup — Mandatory Deliverables (Auto-Create All 7)

Whenever a new engagement is confirmed, **automatically create all 7 documents** in a new folder named after the project. Do not wait to be asked for each one individually.

```
{ProjectName}/
├── {ProjectName}_Project_Schedule.csv        ← generated by schedule script
├── {ProjectName}_Project_Schedule.xml        ← generated by schedule script (MS Project)
├── {ProjectName}_Cost_Plan.csv               ← derived from schedule resource assignments
├── {ProjectName}_Risk_Register_Template.xlsx ← generated by generate_risk_register.py
├── {ProjectName}_Project_Charter.html        ← generated by generate_project_charter.py
├── {ProjectName}_Executive_Dashboard.html    ← mirror AlphaX_Executive_Dashboard.html
├── {ProjectName}_Management_KPI_Dashboard.html ← mirror AlphaX_Management_KPI_Dashboard.html
└── {ProjectName}_Monthly_Status_Report_{MMM_YYYY}.pdf ← generate_{ProjectName}_monthly_report.py
```

**Creation order** (dependencies matter):
1. Confirm all mandatory inputs (see table above) before generating anything
2. **Schedule** first — all other documents derive data from it
3. **Cost plan** — derived from schedule resource assignments
4. **Risk register** — adapted from `generate_risk_register.py`
5. **Project charter** — adapted from `generate_project_charter.py`
6. **Executive dashboard** — mirror AlphaX template; populate with schedule + cost + risk data
7. **KPI dashboard** — mirror AlphaX template; populate with same data
8. **Monthly status report** — run `generate_{ProjectName}_monthly_report.py`

> **Generator scripts to create** for each new project:
> - `generate_{ProjectName}_schedule.py` (copy `generate_alphax_schedule.py`)
> - `generate_{ProjectName}_monthly_report.py` (copy `generate_alphax_monthly_report.py`)

---

## Document Generators

### Project Schedule — Canonical Approach (`generate_alphax_schedule.py` as template)

For each new engagement, **copy `generate_alphax_schedule.py` and adapt it**. This is the proven pattern:

1. **Copy the script** → rename to `generate_{ProjectName}_schedule.py`
2. **Replace the `TASKS` list** with the new project's task breakdown — follow the same column structure:
   `(ID, Outline Level, Name, Duration, Start, Finish, Predecessors, Resource Names, Notes, Milestone)`
3. **Update `OUT_DIR`** to point to the new project folder (e.g. `HERE / "ProjectName"`)
4. **Update `CSV_PATH` and `XML_PATH`** filenames
5. **Run the script** — it writes both `.csv` and `.xml` automatically by calling `generate_msp_xml.py`

```bash
python generate_{ProjectName}_schedule.py
# Outputs: {ProjectName}/{ProjectName}_Project_Schedule.csv
#          {ProjectName}/{ProjectName}_Project_Schedule.xml
```

**Task structure rules (mirror Trinity / AlphaX format):**
- Outline Level 1 = Phase summary
- Outline Level 2 = Sub-workstream or workstream group
- Outline Level 3 = Individual tasks
- Milestones: Duration = `"0 days"`, Milestone column = `"Yes"`
- Date format: `MM/DD/YY` (e.g. `06/01/26`)
- Predecessors: comma-separated task ID numbers (e.g. `"14,15,16"`)
- Resources: `+`-separated names (e.g. `"Bosch IT + KPMG"`)

**Standard phase structure for a new carve-out project:**

| Phase | Content | Gate |
|-------|---------|------|
| Phase 1 — Initialization | Governance setup · IT inventory · charter | QG1 |
| Phase 2 — Concept | As-is analysis · architecture design · migration strategy | QG2 |
| Phase 3 — Development & Build | Infrastructure build · ERP/SAP dev · app prep · device packaging | QG3 |
| Phase 4 — Implementation | Site cutovers · device migration · ERP go-live · app migration | QG4 |
| Phase 5 — GoLive & Closure | Day 1 · hypercare · TSA exit · programme closure | QG5 |

> For Stand Alone model: 5 phases (no Merger Zone / Stabilisation phase). For Integration model: add Phase 6 — Stabilisation & Merger Zone to Buyer Migration.

### Standard Outputs for a New Project

All 7 documents below are **mandatory** and must be created for every new engagement. Store all outputs in a project folder named after the project code:

| # | Output | File name pattern | How |
|---|--------|-------------------|-----|
| 1 | **Project schedule CSV** | `{ProjectName}_Project_Schedule.csv` | `generate_{ProjectName}_schedule.py` |
| 2 | **Project schedule XML** | `{ProjectName}_Project_Schedule.xml` | Same script (calls `generate_msp_xml.py`) — MS Project importable |
| 3 | **Cost plan** | `{ProjectName}_Cost_Plan.csv` | CSV derived from schedule — mirror `AlphaX/AlphaX_Cost_Plan.csv` |
| 4 | **Risk register** | `{ProjectName}_Risk_Register_Template.xlsx` | `generate_risk_register.py` |
| 5 | **Project charter** | `{ProjectName}_Project_Charter.html` | `generate_project_charter.py` |
| 6 | **Executive dashboard** | `{ProjectName}_Executive_Dashboard.html` | Mirror `AlphaX/AlphaX_Executive_Dashboard.html` |
| 7 | **KPI dashboard** | `{ProjectName}_Management_KPI_Dashboard.html` | Mirror `AlphaX/AlphaX_Management_KPI_Dashboard.html` |
| 8 | **Monthly status report** | `{ProjectName}_Monthly_Status_Report_{MMM_YYYY}.pdf` | `generate_{ProjectName}_monthly_report.py` |

> **Do not generate** `Open_Points_List.csv` unless explicitly requested — this requires detailed input data not available at project start.

---

## Project Cost Plan — Standard Template Specification

All cost plans must follow the structure established in `AlphaX/AlphaX_Cost_Plan.csv`. Use `Trinity_Project_Cost_Plan.csv` as an additional structural reference for larger Integration-model engagements.

### File Header (4 lines before data)

```
{ProjectName} - Draft IT Cost Plan
Generated: {date} | Based on: {ProjectName}_Project_Schedule.csv ({N} tasks / {N} phases)
Scope: {Seller} {Division} -> {Buyer} | {Carve-out Model} | {Merger Zone note}
Assumptions: 8 hrs/day; costs are fully-loaded labour only (excl. hardware / co-lo lease / WAN circuits / software licences / travel); shared tasks charged to each assigned resource in full; 0-day milestones = 0 hrs
```

### Cost Categories (rows in fixed order)

| Category | Who | Rate guidance |
|---|---|---|
| **Governance / PMO** | IT PM · PMO team · Steering Committee (active tasks only) | EUR 180/hr |
| **Seller IT Team** | Seller IT architects, infra, ERP, IAM, security, CWP, ops, procurement | EUR 120/hr |
| **Buyer / NewCo Team** | Buyer IT ramp-up, architects, management (onboarding only until org is confirmed) | EUR 100–120/hr |
| **External Partners** | PMO firm (e.g. KPMG) · Migration partner (ERP/SAP) · WAN provider | EUR 90/hr blended |
| **Cross-Functional** | App teams · Integration · Dev · OT Security · Legal · Finance · Comms · Test | EUR 75/hr |
| **Regional Teams** | EMEA · AP · AM regional IT leads and on-site teams | EMEA 180 · AP 99 · AM 165 |
| **Executive** | Seller + Buyer executive leadership (active governance tasks only) | EUR 180/hr |

Each category block ends with a `SUBTOTAL` row. After all categories, include:
- A double `====================` separator row
- `OVERALL PROJECT TOTAL` row
- Another `====================` separator row

### Summary Sections (after totals)

1. **COST BREAKDOWN BY CATEGORY** — one row per category with EUR total
2. **COST BREAKDOWN BY PHASE** — one row per programme phase with EUR total
3. **CAPEX / ADDITIONAL COSTS (NOT INCLUDED IN LABOUR TOTAL)** — itemised list of hardware, WAN, co-lo, licences, MSP/ITO contracts — all marked TBC unless confirmed
4. **NOTES** — numbered list covering: methodology, model-specific adjustments, rate rationale, TSA scope, budget approval gate

### Key Rules

- **Labour only** — hardware, software licences, WAN circuits, co-lo leases, and travel are always excluded from the labour total and listed separately under CAPEX
- **Stand Alone model**: no seller-integration phase; buyer team costs are ramp-up only; TSA is short and minimal scope
- **Integration model**: add buyer integration team; include Merger Zone build and TSA exit phases; TSA typically 12–18 months
- **Buyer team costs** are estimates only until the buyer's IT organisation is confirmed — flag this clearly in notes
- **Budget baseline** must be formally approved at QG1 — mark as "TBC, approved at QG1" in all documents until then
- **Rate benchmarks**: EMEA EUR 180/hr · AP EUR 99/hr · AM EUR 165/hr · seller/buyer IT EUR 120/hr · external partners EUR 90/hr · cross-functional EUR 75/hr
- Derive day totals from the project schedule resource assignments; do not invent figures

### Columns

`CATEGORY, RESOURCE, TOTAL DAYS, TOTAL HRS, HOURLY RATE (EUR), TOTAL COST (EUR)`

---

## Executive Dashboard — Standard Template Specification

All executive dashboards must follow the structure and style established in `AlphaX/AlphaX_Executive_Dashboard.html`. This is the canonical reference implementation — open it as a base and replace all AlphaX-specific data with the new project's data.

### Design Rules (non-negotiable)

| Rule | Requirement |
|------|-------------|
| **Self-contained** | No external CDN, no Google Fonts, no remote URLs of any kind — 100% inline HTML+CSS |
| **Fonts** | `'Segoe UI', Arial, sans-serif` system stack only. Do not add any web font `<link>` tags |
| **Brand colors** | **Bosch Digital palette** (from `New Bosch Digital color theme.png`): primary red `#ED0007`, dark navy `#004975`, blue `#007BC0`, dark teal `#0A4F4B`, dark green `#00512A`, charcoal `#43464A`, mid-gray `#71767C`, light bg `#F5F5F5`, white `#FFFFFF`; extended: purple `#9E2896`, teal `#18837E` |
| **Logo** | Use actual `Bosch-Logo.png` (root folder) — embed as base64 data URI in HTML (`<img src="data:image/png;base64,...">`); load via `reportlab.lib.utils.ImageReader` in PDF scripts. Never use inline SVG text approximation. |
| **Print** | Include `@media print` CSS with `page-break-before: always` on `.page-break` divs |
| **SharePoint** | Must open and render correctly when uploaded to SharePoint with no internet connectivity |

### Layout Structure (3 Sections — mirror exactly)

#### Section 1 — Overview & Timeline
1. **Header bar** — Dark navy (`#004975`) background · `Bosch-Logo.png` embedded (left) · Programme title + seller/buyer/model subtitle (centre) · Date range + dashboard date (right)
2. **Countdown strip** — Bosch red (`#ED0007`) background · 3 countdown figures (days to Kick-Off, days to Day 1, days to Closure) · Overall programme health badge (AMBER/GREEN/RED) on the right
3. **Section header label** — `PROGRAMME OVERVIEW` (dark bar, white uppercase label)
4. **Project overview** — Two-column layout: left = 3-paragraph free text description; right = 4 metadata boxes (Carve-Out Model, Key Parties with colour badges, Programme Budget, Governance)
5. **Stats row** — 6 equal-width tiles across full width, each with: large icon, bold number/value, uppercase label, sub-label. Top border: 3px Bosch red. Metrics: Sites · Employees · Devices · Applications · Duration · TSA
6. **Phase timeline** — Horizontal segmented bar (5 colour segments, one per phase) with phase name + date range in each segment; date labels below; legend below that
7. **Milestones + Budget** — Two-column: left = milestone table (icon/name/description/date/days-offset/status pill); right = budget total box (red) + breakdown rows with mini progress bars

#### Section 2 — Workstreams, Quality Gates & Risks
8. **Workstream grid** — 3×3 grid of 9 IT sub-workstreams (WS1–WS9), each card showing: title, sub-topics, confidence % badge with colour (🟢/🟡/🔴)
9. **Quality Gate Tracker** — Vertical list of QG1–QG5 rows; each row: date + offset + status pill (left), QG title + entry/exit criteria description (right)
10. **Regional scope** — Horizontal bar chart showing site/device count per region (EMEA / AP / AM) with proportional fill bars
11. **Risk indicators** — Stack of risk cards, each with colour-coded left border (red=HIGH, amber=MEDIUM, green=LOW), risk title, description, and P·I·Rating footer

#### Section 3 — Apps, Countries & Critical Path
12. **Application migration waves** — 3 rows (Wave 1/2/3), each with: label + sub-description, proportional fill bar, target date
13. **Country complexity hotspots** — 4-column grid: Germany / China / India / Mexico — flag emoji + country name heading + 2–3 lines of specific regulatory/technical risk
14. **Stats strip** — Dark background, 4 spotlight numbers (QGs, Workstreams, Sites, Co-Lo Hubs)
15. **Critical path + principles** — 4 columns: Infrastructure path · ERP path · Client Devices path · Guiding Principles (with red arrow bullets)
16. **Footer** — Dark background · Project name + parties (left) · Classification + date + confidentiality notice (right)

### CSS Patterns to Reuse

```css
/* Bosch Digital design tokens */
:root {
  --bosch-red:   #ED0007;  /* Bosch Digital primary red   */
  --bosch-dark:  #004975;  /* Bosch Digital dark navy     */
  --bd-blue:     #007BC0;  /* Bosch Digital blue          */
  --bd-teal-dk:  #0A4F4B;  /* Bosch Digital dark teal     */
  --bd-green-dk: #00512A;  /* Bosch Digital dark green    */
  --bosch-mid:   #43464A;  /* body text                   */
  --bosch-muted: #71767C;  /* muted / labels              */
  --bosch-bg:    #F5F5F5;
  --good:        #00884A;  /* Bosch Digital success green */
}

/* Section header */
.sh { background: #004975; color: #fff; padding: 5px 14px; font-size: 9.5px; font-weight: 700; letter-spacing: 1.5px; text-transform: uppercase; }

/* Status pills */
.pill { display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 9px; font-weight: 700; }
.p-upcoming { background: #E8F4FD; color: #1565C0; }
.p-critical  { background: #FEE8EA; color: #9B0006; }
.p-planned   { background: #F3F3F3; color: #666; }
.p-closure   { background: #E6F4EC; color: #00512A; }

/* Risk cards */
.risk-card.hi  { background: #FEE8EA; border-left: 4px solid #9B0006; }
.risk-card.med { background: #FFF3E0; border-left: 4px solid #C66000; }
.risk-card.lo  { background: #E6F4EC; border-left: 4px solid #00512A; }

/* Page break (print) */
.page-break { border-top: 4px solid #ED0007; margin-top: 2px; }
@media print { .page-break { page-break-before: always; border-top: none; } }
```

### Data to Populate Per Project

Before generating the dashboard, ensure you have all of the following from the user or from existing project documents:

| Field | Source |
|-------|--------|
| Seller / Buyer / PMO | User input (mandatory) |
| Carve-out model (Stand Alone / Integration / Combination) | User input |
| Sites, Users, Devices, Apps | User input or estimate |
| Duration and key dates (Kick-Off, Signing, QG1–QG5, Day1, Closure) | User input |
| Days-offset for each milestone (from "Dashboard as of" date) | Calculate from dates |
| Phase names and date ranges | Derive from QG dates |
| Top 3–5 risks with P / I / Rating | Risk register or user input |
| App wave breakdown (Wave 1/2/3 counts and targets) | User input or estimate |
| Country complexity list | Derive from site regions |
| Budget (or "TBC at QGx") | User input |
| Workstream confidence levels | User input or default 🟡 if unknown |

---

## Monthly Status Report — Canonical Approach (`generate_alphax_monthly_report.py` as template)

For each new engagement, **copy `generate_alphax_monthly_report.py` and adapt it**. The script is designed for unattended scheduled execution (Power Automate or any scheduler) — no arguments, no interactive prompts.

### How It Works

- All "days to gate" values auto-calculate from `datetime.date.today()` at runtime
- Output filename auto-includes the current month/year: `{ProjectName}_Monthly_Status_Report_{MMM_YYYY}.pdf`
- Running it again in May produces `..._May_2026.pdf` with refreshed countdown numbers automatically

### Adapting for a New Project

1. **Copy the script** → rename to `generate_{ProjectName}_monthly_report.py`
2. **Update the PROJECT CONFIGURATION block** (top of file):
   - `PROJECT_CODE`, `PROJECT_TITLE`, `PROJECT_SUBTITLE`
   - `SELLER`, `BUYER`, `PMO`, `MODEL`
   - `SCOPE_TEXT`, `LABOUR_BUDGET`, `DURATION`
   - `DATES` dict (ISO dates for kickoff, signing, QG1–QG5, day1, closure)
   - `PHASE_ROWS`, `TOP_RISKS`, `BUDGET_ROWS`, `PROGRAMME_NOTES`
3. **Update `OUT_DIR`** to point to the new project folder
4. **Ensure `Bosch-Logo.png`** is present in the root workspace folder (loaded via `ImageReader`)

### Report Layout (single A4 page — mirrors `Sample_Monthly_Status_Report.pdf`)

| Section | Content |
|---|---|
| **Header** | Bosch logo + project title + subtitle + report date / PMO / next review date |
| **Metadata bar** | Red strip — 6 auto-updating fields: Status · Duration · Day 1 · Scope · Budget · Days to Day 1 |
| **Phase Status & Timeline** | 5-phase table with auto-coloured status pills (Upcoming / Planned / CRITICAL / Complete) |
| **Key Risks & Immediate Actions** | Two-column: top 5 risks with P×I rating + colour (left) · decision actions for next 2 weeks (right) |
| **Budget Snapshot** | All resource categories + TOTAL row from cost plan |
| **Critical Milestones & QGs** | 8 rows (Kickoff → QG1–QG5 → Day 1 → Closure) with auto-calculated days-to-gate |
| **Programme Notes** | 6 action items / key messages for the current period |
| **Footer** | Distribution list · Confidentiality · Generated date |

### Bosch Digital Colours in PDF (reportlab)

```python
C_RED   = colors.HexColor("#ED0007")   # primary red
C_NAVY  = colors.HexColor("#004975")   # header / section bars
C_TEAL  = colors.HexColor("#0A4F4B")   # risk / actions section
C_GREEN = colors.HexColor("#00512A")   # closure / complete states
C_BLUE  = colors.HexColor("#007BC0")   # upcoming states
C_MID   = colors.HexColor("#43464A")   # body text
C_MUTED = colors.HexColor("#71767C")   # labels
```

### Power Automate Setup

- Action type: **Run Python script** (or HTTP trigger → Run script)
- Script path: `generate_{ProjectName}_monthly_report.py` (full absolute path)
- Python executable: `C:/Program Files/px/python.exe` (Bosch embedded Python)
- Schedule: **1st of each month** (or last working day)
- Output: PDF written to `{ProjectName}/` subfolder automatically
- No parameters, no secrets, no environment variables required


## IT Sub-Workstream Structure (from FRAME — use as default template)

1. IT Infrastructure (WAN/LAN, AD, servers, M365/Azure, telephony)
2. Commercial IT incl. ERP (SAP migration, CRM, FSM, BPO)
3. Other Applications (~500 apps; SharePoint, Confluence/Jira, etc.)
4. Engineering IT (PLM/Windchill, FOSS compliance, developer network)
5. Production IT (OT Security, MES, DOT, plant telephony)
6. HR IT
7. IT Organization & Processes (TOM, ITO contracting, ITSM/ServiceNow)
8. IT Contracts & Licenses (change of control, SAM, FOSS)
9. IT Security (IAM/Saviynt, CISO, ISO 27001, BCM, GDPR)

## Key Carve-Out Concepts to Apply

- **"First make it work, then make it better"** — pragmatism over optimization; Day-1 readiness is the goal
- **Carve-out models**: Stand Alone (full independence), Integration with Buyer (buyer leads), or Combination
- **TSA (Transitional Service Agreement)**: Parent provides temporary services post-closing; minimize scope and duration
- **Big Bang vs. Staggered go-live**: FRAME used Big Bang due to tightly coupled logistics/ERP processes
- **Point of No Return**: An implemented production change that cannot be reversed without significant cost/compliance impact
- **Risk rating** = Probability × Impact (1–5 scale); categories: ScR=Schedule, SR=Scope, RR=Resource, BtR=Budget, QR=Quality, BR=Business, LR=Legality, CR=Customer Satisfaction
- **Country-specific complexity**: Brazil ERP (extreme tax law complexity), China (local FTS/customs), Mexico (legal entity delays), India (local IT solutions like RBIN)

## Typical Timeline Benchmarks (from FRAME)

- Total project duration: ~18–24 months from kick-off to closing
- WAN ordering lead time: 4–6 months minimum
- Active Directory build and go-live: ~6 months
- SAP migration (shell copy + testing): 9–12 months
- M365 tenant build + cutover: ~6 months
- Phone number migration: 4–8 weeks per operator (state-owned operators may take longer)
- IAM/IdM implementation: 6–9 months; Saviynt SC2 connector in Bosch-governed env restricted to 10 days before closing
- Infrastructure hub setup (co-locator): decision by month 4, implementation by month 9
- "Frozen Zone" begins at Signing — minimize changes to production environment
