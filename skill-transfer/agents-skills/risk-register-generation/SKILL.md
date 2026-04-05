---
name: risk-register-generation
description: "Use when generating carve-out risk registers with probability x impact scoring, risk categories, mitigation plans, ownership, and top-risk prioritization."
---

# Risk Register Generation

## Purpose
Generate a project-specific risk register using carve-out risk patterns and templates.

## Canonical Sources
- `BD_Risk-Register_template_en_V1.0_Dec2023.xlsx` — **the authoritative template for all new projects** (V1.0 Dec 2023); use for structure, lookup ranges, and formatting
- Existing project risk register generators (e.g. `generate_bravo_risk_register.py`) — **template-handling code reference only**; never copy their risk descriptions, categories, owners, dates, or mitigation plans
- `20240718_FRAME_IT Risk Assessment_completed.csv` — FRAME reference data (methodology patterns only; never copy specific risk entries)

## Template: BD_Risk-Register_template_en_V1.0_Dec2023.xlsx

### Sheet names (exact, case-sensitive)
| Sheet | Purpose |
|---|---|
| `Info` | Project metadata (document ID, file name, project name, owner) |
| `Risk Register` | Risk data table |
| `Matrix ` | Probability/Impact matrix — do not modify |

### Info Sheet cell mapping
| Cell | Label |
|---|---|
| `C4` | Document ID |
| `C5` | File Name |
| `C6` | Project Name |
| `C7` | Project ID |
| `C8` | Document Owner |

### Risk Register Sheet column layout
- **Header row**: 4 (group headers at row 3)
- **First data row**: 5

| Col | Letter | Field |
|---|---|---|
| 2 | B | Risk ID (sequential) |
| 3 | C | Creation date |
| 4 | D | Risk category (see categories below) |
| 5 | E | Cause(s) |
| 6 | F | Event (risk description) |
| 7 | G | Effect(s) |
| 8 | H | Risk event date |
| 9 | I | Owner |
| 10 | J | Source |
| 12 | L | Impact (text: `Very Low` / `Low` / `Moderate` / `High` / `Very High`) |
| 13 | M | VLOOKUP: impact numeric — write formula explicitly |
| 14 | N | Probability (text: `10%` / `30%` / `50%` / `70%` / `90%`) |
| 15 | O | VLOOKUP: probability numeric — write formula explicitly |
| 16 | P | `threat` or `opportunity` |
| 17 | Q | Matrix score = `=M{row}*O{row}` |
| 18 | R | Qualitative impact description |
| 19 | S | Monetary impact EUR current year |
| 20 | T | Monetary impact EUR 3 subsequent years |
| 22 | V | Risk response strategy (`Avoid` / `Transfer` / `Mitigate` / `Accept` / `Exploit` / `Enhance` / `Share`) |
| 23 | W | Measure (mitigation actions) |
| 24 | X | Due date |
| 26 | Z | Status (`not started` / `in progress` / `on hold` / `implemented` / `cancelled`) |
| 27 | AA | Reporting date |
| 28 | AB | Impact actual (same as L initially) |
| 29 | AC | VLOOKUP: actual impact numeric |
| 30 | AD | Probability actual (same as N initially) |
| 31 | AE | VLOOKUP: actual probability numeric |
| 32 | AF | `=AC{row}` |
| 33 | AG | `=AE{row}` |
| 34 | AH | `=AF{row}*AG{row}` (actual matrix score) |
| 35 | AI | Notes |

### Template-controlled lookup ranges
- Risk category list: `D140:D156` (17 values)
- Risk source list: `D171:D175`
- Impact values: `D182:E186`
- Probability values: `D189:E193`
- Traffic-light thresholds: `D196:E199`
- Risk strategy list: `D202:D208`
- Status list: `D213:D217`

These ranges are template-owned. Do not move or rewrite them in generated workbooks.

### VLOOKUP formula pattern (write for ALL data rows)
```python
ws.cell(r, 13).value = f'=_xlfn.IFNA(VLOOKUP(L{r},$D$182:$E$186,2,FALSE),"")'  # M - impact numeric
ws.cell(r, 15).value = f'=_xlfn.IFNA(VLOOKUP(N{r},$D$189:$E$193,2,FALSE),"")'  # O - probability numeric
ws.cell(r, 17).value = f'=M{r}*O{r}'                                       # Q - matrix score
ws.cell(r, 29).value = f'=_xlfn.IFNA(VLOOKUP(AB{r},$D$182:$E$186,2,FALSE),"")'  # AC
ws.cell(r, 31).value = f'=_xlfn.IFNA(VLOOKUP(AD{r},$D$189:$E$193,2,FALSE),"")'  # AE
ws.cell(r, 32).value = f'=AC{r}'
ws.cell(r, 33).value = f'=AE{r}'
ws.cell(r, 34).value = f'=AF{r}*AG{r}'
```

### Matrix formatting rule
- The `Matrix ` sheet is template-owned and should not be structurally modified.
- When saving with `openpyxl`, explicitly set black font on yellow-filled matrix cells (`FFFFFF00`, `FFFFFFCC`). Theme-inherited font colour may otherwise be lost on save, making yellow cells unreadable in Excel.

### Risk Category values (new template taxonomy)
`Technology, R&D`, `Engineering`, `Manufacturing`, `Quality`, `Strategy & Portfolio`, `Budget`, `Schedule`, `Resources`, `Supply Chain`, `Market & Competitors`, `Customers`, `Raw Materials`, `Stakeholder Relations & Public Affairs`, `Intellectual Property`, `Legal & Compliance`, `Ecosystems & Ethics`, `Security & Data Protection`

### Risk Source values
`Formal Risk Review`, `SQA Audit/Review`, `Status Meeting`, `Stakeholder`, `Other`

### Probability/Impact mapping from 1–5 scale
| 1–5 | Impact label | Probability |
|---|---|---|
| 1 | Very Low | 10% |
| 2 | Low | 30% |
| 3 | Moderate | 50% |
| 4 | High | 70% |
| 5 | Very High | 90% |

## Core Method
- Risk rating = Probability × Impact (1–5 scale each); threshold for high priority = 12.
- Use category taxonomy: ScR, SR, RR, BtR, QR, BR, LR, CR.
- Include mitigation, owner, target date, and status per risk.
- Use plain ASCII punctuation in generated narrative text where practical. Prefer ` - ` over an em dash in scripted workbook text to avoid encoding corruption in Excel-bound outputs.

## Risk Identification: Comprehensive Category Evaluation
**Critical requirement:** Risk identification must be *systematic and comprehensive*, evaluating **all 17 risk categories** and identifying realistic threats *and opportunities* specific to the project. A register with fewer than 15-20 risks (for standard IT carve-outs) is likely under-identified and will not meet M&A governance standards.

**Process:**
1. For each of the 17 categories, assess applicability to the project scope:
   - **Not applicable categories** (e.g. `Market & Competitors`, `Raw Materials`, `Intellectual Property` for IT-only carve-outs): Document and exclude with clear rationale
   - **Applicable categories**: Identify 1–3 realistic threats *and 0–2 opportunities* per category tied to project-specific factors (carve-out model, timeline, technology landscape, number of sites/users, TSA scope, manufacturing cutover, etc.)
2. Generate 20–30 risks minimum for complex carve-outs (>500 users, >15 applications, multi-site, manufacturing involvement, integration model)
   - Include **at least 2 positive risks (opportunities)** where carve-out enables strategic advantage or cost optimization (e.g., standalone infrastructure enables cloud migration, vendor consolidation OpEx savings, independent licensing negotiation leverage)
3. Ensure risk distribution across categories reflects true exposure: do not over-weight Schedule or under-weight Technology/Security
4. Risks must be *project-specific* — grounded in current engagement scope, not generic M&A templates
5. **Positive risks (Opportunities):** Where carve-out structure enables strategic advantage, must include response strategy of `Exploit` (maximize upside) or `Enhance` (strengthen enablement)

## Risk Description Requirements: Detailed Causes and Events
**Critical requirement:** Risk `Cause` (column E) and `Event` (column F) must be *detailed and context-specific*, not abbreviated or generic.

**Cause Column (E) — Root Factors with Project Quantification:**
- Include specific scope metrics: number of sites, users, applications, systems, timelines
- Example INADEQUATE: "ISP leads"
- Example ADEQUATE: "ISP lead times for 3 Singapore sites 8-12 weeks; provider allocation constraints; late RFQ submission by Apr 20"
- Embed root drivers: supply chain delays, shared infrastructure, compressed UAT windows, change-of-control triggers, shared data boundaries, legacy system interdependencies
- Reference specific project dates/milestones: GoLive Nov 7, hypercare Nov 8 - Feb 6, Phase 3 UAT Oct 5-31, QG4 gate Nov 1

**Event Column (F) — Specific Failure Scenarios:**
- Describe the specific failure mechanism (not generic "delayed")
- Tie to schedule impact: which phase affected, which milestone at risk, how many days delayed
- Example INADEQUATE: "Network unavail"
- Example ADEQUATE: "Network infrastructure not provisioned by Aug 1 (Phase 3 UAT target); redundant WAN circuits not available; Go-Live Nov 7 at risk; parallel run capacity insufficient"
- Include concrete trigger conditions and impact pathways (GL posting failures → financial close delays, inventory accuracy drifts → revenue impact, support ticket volume 80-120/day vs 40 projected)
- Quantify business impact where possible (EUR amounts, operational disruption hours, user population affected)

**Best Reference:** Examine `active-projects/Zebra/Zebra_Risk_Register.xlsx` rows 5-14 for professional elaboration pattern showing project-specific risk narratives.

## Python Implementation Rules
- Always add `sys.path.insert(0, os.path.join(os.path.expanduser("~"), "py_packages"))` before importing `openpyxl`.
- The correct Python interpreter on this machine is `C:/Program Files/px/python.exe`.
- **Do not** attempt to read existing `.xls` files as source data — they may be XML-format XLS which openpyxl cannot parse. Embed risk data directly in the generation script as a Python list/dict structure.
- Use `load_workbook(template_path)` to open `BD_Risk-Register_template_en_V1.0_Dec2023.xlsx`; populate cells, then `wb.save(output_path)`.
- Output file must be `.xlsx` — never `.xls` or `.csv` only.

## Generation Rules
- **Never copy risk entries, descriptions, causes, effects, owners, or mitigation plans from any existing project** (Bravo, AlphaX, Falcon, Trinity, Hamburger, or any other).
- Generate all risks fresh based on the current project's specific scope, complexity, timeline, carve-out model, and technical landscape (e.g. number of applications, SAP involvement, TSA, number of sites/users).
- Keep risks specific to the current engagement parties and scope only.
- Do not copy historical project facts (parties, dates, scope) from reference files or other project folders.
- **Risk register must be comprehensive:** Systematically evaluate all 17 risk categories (see "Risk Identification: Comprehensive Category Evaluation" section above); identify 20–30 risks minimum for complex carve-outs; ensure causes and events are detailed and project-specific (not abbreviated).
- Prioritize risks by P×I rating and delivery impact.
- Ensure top risks align with schedule critical path and quality gates.
- Always generate a `.xlsx` file using `BD_Risk-Register_template_en_V1.0_Dec2023.xlsx` as the base — not a blank workbook, not the old `Risk_analysis_template.xlsx`.
- The output file must be saved to `<ProjectName>/<ProjectName>_Risk_Register.xlsx`.
- Do not reference `AlphaX/AlphaX_Risk_Register_Template.xlsx` or `Risk_analysis_template.xlsx` as the template — they are deprecated.
