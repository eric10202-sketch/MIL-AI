---
name: risk-register-generation
description: "Use when generating carve-out risk registers with probability x impact scoring, risk categories, mitigation plans, ownership, and top-risk prioritization."
---

# Risk Register Generation

## Purpose
Generate a project-specific risk register using carve-out risk patterns and templates.

## Canonical Sources
- `generate_risk_register.py` — reference implementation (Trinity project)
- `Risk_analysis_template.xlsx` — **the only authoritative Excel template** for all projects
- `20240718_FRAME_IT Risk Assessment_completed.csv` — FRAME reference data (methodology only)

## Template: Risk_analysis_template.xlsx

### Sheet names (exact, case-sensitive)
| Sheet | Purpose |
|---|---|
| `Cover sheet` | Project metadata (manager names) |
| `Analysis of project risks` | Risk data table |
| `Explanation_History` | Reference — do not modify |

### Cover Sheet cell mapping
| Cell | Label |
|---|---|
| `D1` | IT-as-is Manager name |
| `D2` | IT-Project Manager name |
| `D3` | PMI-Lead name |

### Analysis Sheet column layout
- **Header row**: 7
- **First data row**: 9
- **Pre-built formula rows**: template already contains `=$H{row}*$I{row}` in column J from row 9.
  For any rows added beyond the template's pre-built range, set column J explicitly: `ws.cell(row_num, 10).value = f"=$H{row_num}*$I{row_num}"`

| Col | Letter | Field |
|---|---|---|
| 1 | A | No. (sequential row index) |
| 2 | B | Sub-project (optional) |
| 3 | C | Entry Date of Risk |
| 4 | D | Risk category (e.g. ScR - Schedule) |
| 5 | E | Risk — concrete, understandable description |
| 6 | F | **Effects** (severity significance — maps to T/Impact column) |
| 7 | G | **Causes** |
| 8 | H | W — Probability (1–5) |
| 9 | I | T — Impact (1–5) |
| 10 | J | RZ = W × T formula (pre-built in template; do not overwrite for in-range rows) |
| 11 | K | Actions — preventive and counter measures |
| 12 | L | Responsible |
| 13 | M | Deadline |
| 14 | N | Status of action |
| 15 | O | Remarks / notes (optional) |

### Column F vs O distinction
- **Column F (Effects)**: what will happen if the risk materialises — business or schedule consequences.
- **Column O (Remarks)**: supplementary notes, accepted residual risk statements, or response strategy label.
- Do **not** swap these. Notes/remarks go in O; effects/consequences go in F.

## Core Method
- Risk rating = Probability × Impact (1–5 scale each); threshold for high priority = 12.
- Use category taxonomy: ScR, SR, RR, BtR, QR, BR, LR, CR.
- Include mitigation, owner, target date, and status per risk.

## Python Implementation Rules
- Always add `sys.path.insert(0, os.path.join(os.path.expanduser("~"), "py_packages"))` before importing `openpyxl`.
- The correct Python interpreter on this machine is `C:/Program Files/px/python.exe`.
- **Do not** attempt to read existing `.xls` files as source data — they may be XML-format XLS which openpyxl cannot parse. Embed risk data directly in the generation script as a Python list/dict structure.
- Use `load_workbook(template_path)` to open `Risk_analysis_template.xlsx`; populate cells, then `wb.save(output_path)`.
- Output file must be `.xlsx` — never `.xls` or `.csv` only.

## Generation Rules
- Keep risks specific to the current engagement parties and scope only.
- Do not copy historical project facts (parties, dates, scope) from reference files or other project folders.
- Prioritize risks by P×I rating and delivery impact.
- Ensure top risks align with schedule critical path and quality gates.
- Always generate a `.xlsx` file using `Risk_analysis_template.xlsx` as the base — not a blank workbook.
- The output file must be saved to `<ProjectName>/<ProjectName>_Risk_Register.xlsx`.
- Do not reference `AlphaX/AlphaX_Risk_Register_Template.xlsx` as the template — it is not the authoritative source.
