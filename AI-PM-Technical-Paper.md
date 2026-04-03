# AI-Assisted Project Management for IT Carve-Outs  
## A Technical Paper on the Carveout AI Toolkit

**Author:** Erik Ho (BD/MIL-ICC), Bosch Global Business Services  
**Date:** April 2026  
**Classification:** Internal — Bosch Group  
**Repository:** https://github.com/erichokh/MIL-Carveout-project

---

## Abstract

This paper describes the design, architecture, and operational practice of an AI-assisted project-management toolkit developed within Bosch's MIL (Mergers, Integrations, and Liquidations) practice. The toolkit uses GitHub Copilot (powered by Claude Sonnet) running inside Visual Studio Code in agent mode to automate the generation of all mandatory IT carve-out deliverables — project schedule, risk register, cost plan, project charter, executive dashboard, management KPI dashboard, and monthly status reports. The system combines a structured repository on Bosch DevCloud/GitHub, a domain-specific instruction layer (`copilot-instructions.md`), modular skill files that encode methodology logic, Python generator scripts, and a memory framework that persists project-level knowledge across sessions. The result is a significant reduction in manual document-production effort while enforcing methodological consistency across all engagements.

---

## 1. Background and Motivation

IT carve-outs are among the most document-intensive activities in M&A. A typical engagement requires a project manager to produce and keep in sync seven major artifacts — schedule, risk register, cost plan, project charter, two HTML dashboards, and a monthly report — each with interdependencies, complex formatting requirements, and data derived from the same core scope facts (buyer, seller, sites, users, applications). Producing these documents manually is time-consuming, error-prone, and inconsistent across projects. Small differences in assumptions — a phase date that drifts, a resource label that changes, a cost line not reconciled with the risk register — compound into governance problems at quality gates.

The initiative described here arose from real project experience on engagements such as Project Trinity (reference methodology project), Project AlphaX, Project Falcon, Project Bravo, and Project Hamburger. The goal was to build a toolkit where a project manager could supply the ten mandatory engagement parameters and receive a complete, consistent, methodologically correct set of deliverables within minutes — with all cross-checks enforced automatically, not by human review.

---

## 2. Platform Architecture

### 2.1 Tool Stack

The toolkit runs on a combination of commercial and enterprise-grade tools:

| Layer | Tool | Role |
|---|---|---|
| AI inference | GitHub Copilot (Claude Sonnet 4.6) | Reasoning, code generation, content synthesis |
| IDE | Visual Studio Code | Primary workspace and chat interface |
| Version control | Git + GitHub (Bosch DevCloud / github.com) | Repository storage, history, team sharing |
| Python runtime | CPython 3.x (`C:/Program Files/px/python.exe`) | Script execution |
| Document generation | Python stdlib + `openpyxl` | CSV, XML, HTML, PDF output |
| Collaboration | Microsoft SharePoint / OneDrive | File distribution and offline HTML rendering |

### 2.2 Repository Structure

The repository follows a four-zone content model:

```
Carveout/
├── .claude/
│   └── skills/                    ← Domain-skill instruction files
│       ├── intake-compliance-gate/SKILL.md
│       ├── schedule-generation/SKILL.md
│       ├── cost-plan-generation/SKILL.md
│       ├── risk-register-generation/SKILL.md
│       ├── executive-dashboard-generation/SKILL.md
│       ├── management-kpi-dashboard-generation/SKILL.md
│       ├── monthly-status-report-generation/SKILL.md
│       └── repository-governance-updates/SKILL.md
├── .github/
│   └── copilot-instructions.md    ← Auto-loaded workspace instructions
├── references/                    ← Methodology-only reference assets
├── templates/                     ← Blank project-agnostic templates
│   └── Risk_analysis_template.xlsx
├── active-projects/               ← Current engagement metadata
├── archive/                       ← Closed engagement artifacts
│
├── AlphaX/                        ← Project-specific output folder
├── Bravo/
├── Falcon/
├── Hamburger/
│
├── generate_msp_xml.py            ← Canonical CSV→XML converter
├── generate_{project}_schedule.py ← Per-project schedule generators
├── generate_{project}_risk_register.py
├── generate_{project}_charter.py
├── generate_{project}_monthly_report.py
├── Bosch.png                      ← Embedded logo asset
└── CLAUDE.md                      ← Always-on AI guardrails
```

**Content model rules (strictly enforced):**
- `references/` and `archive/` are methodology-only. The AI is instructed never to copy reference party names, dates, or scope into active engagement deliverables.
- Every project output lives in its own named folder (`AlphaX/`, `Bravo/`, etc.) — isolated from all other engagements.
- Generator scripts are project-specific (one script, one engagement) so that parameters are hard-coded and auditable, not interpolated at runtime from unvalidated input.

---

## 3. The AI Instruction Layer

### 3.1 copilot-instructions.md — Workspace-Level Guardrails

Visual Studio Code's GitHub Copilot extension automatically loads `.github/copilot-instructions.md` for any workspace that contains it. This file is the **always-on constitution** of the AI's behaviour in this repository. It contains:

- **Mandatory intake fields:** The AI is blocked from generating any deliverable unless all eleven engagement parameters are confirmed (project name, seller, buyer, business carved out, carve-out model, PMO lead, worldwide sites, IT users, project start date, GoLive date, completion date).
- **Derivation rules:** Sponsor Customer = Buyer; Sponsor Contractor = Seller; IT flow direction = Seller IT → Merger Zone (if Integration model) → Buyer IT; TSA = Seller operates services until buyer-side readiness.
- **Deliverable orchestration sequence:** Schedule → Risk Register → Cost Plan → Project Charter → Executive Dashboard → Management KPI Dashboard → Monthly Status Report. Each deliverable depends on all preceding ones; the AI cannot skip steps.
- **Output format standards:** Schedule always in both CSV and MSPDI XML; risk register always in `.xlsx` using the authoritative template; HTML deliverables always self-contained (no external CDN links); Bosch logo always embedded as base64 PNG.
- **Skill routing table:** A lookup table maps each task type to the corresponding skill file under `.claude/skills/`, instructing the AI to read the full skill before acting.

### 3.2 CLAUDE.md — Redundant Guardrail

`CLAUDE.md` at the workspace root is a second copy of the global guardrails, formatted for the Claude AI's native instruction format. Having both files ensures the rules are loaded regardless of which entry point is used (native Copilot chat vs. direct Claude API). The two files are kept in sync as part of repository governance.

### 3.3 Skill Files — Domain Knowledge Encoding

The `.claude/skills/` directory contains eight SKILL.md files, each encoding the complete methodology for one deliverable type. A skill file is not a prompt template — it is a structured procedural specification that the AI reads and executes deterministically. Each skill specifies:

- **Purpose** — what the deliverable achieves
- **Canonical sources** — which existing files are the authoritative references
- **Step-by-step generation procedure** — numbered, unambiguous
- **Schema and column definitions** — exact field names, data types, and allowed values
- **Cross-check rules** — mandatory consistency checks between deliverables
- **Format non-negotiables** — visual standards, colour palette, logo embedding rules
- **Python implementation rules** — interpreter path, library imports, output paths

This separation of *knowledge* (skill files) from *instructions* (copilot-instructions.md) is a deliberate architectural choice. It keeps the global instruction file short enough to be reliably loaded, while allowing rich domain detail in skill files that are loaded on demand.

---

## 4. The Compliance Gate

The intake-compliance-gate skill operationalises a strict validation step before any document generation begins.

**Validation contract:**
- If all eleven mandatory fields are present → emit a validation-passed summary listing all inputs, then proceed to deliverable generation.
- If any field is missing → block all generation, list every missing field in a single response, and wait for the user to supply them. No fielding incomplete data; no substitutions from reference projects.

**Budget exception:** If the user explicitly states that budget is unknown or pending approval, the AI substitutes the literal string `TBC - to be approved at QG1` rather than guessing. This is the only permitted exception to the no-substitution rule.

**Buyer/seller derivation:** Once seller and buyer are confirmed, several downstream fields are automatically derived. The AI does not ask for them separately; it applies the derivation rules from the instruction layer. This keeps the user interaction lean — eleven fields in, full engagement context out.

The compliance gate design reflects a real project risk: in practice, AI tools can generate plausible-looking content from incomplete inputs. The gate pattern forces the AI to surface incompleteness explicitly rather than fill gaps silently.

---

## 5. Schedule Generation

### 5.1 Two-File Output Contract

Every schedule generation produces exactly two files:
- `{ProjectName}_Project_Schedule.csv` — human-readable, directly importable to Excel
- `{ProjectName}_Project_Schedule.xml` — Microsoft Project Input format (MSPDI)

The CSV is the source of truth. The XML is always derived from the CSV via `generate_msp_xml.py`, never hand-written.

### 5.2 The Canonical XML Generator

`generate_msp_xml.py` is the central piece of tooling infrastructure. It is a pure Python stdlib script (no non-standard dependencies) that reads any project's CSV schedule and emits a standards-conformant MSPDI XML file. Its significance is in encoding a set of XML rules that are not documented clearly in the Microsoft Project specification, but which are critical for correct import behaviour:

| Rule | Consequence if violated |
|---|---|
| `<TaskMode>1</TaskMode>` must appear immediately after `<Name>` in each Task element | MS Project silently ignores the manual-task flag; all dates recalculate from predecessors on import |
| `<ManualStart>` and `<ManualFinish>` must accompany `<Start>` and `<Finish>` | MS Project displays auto-calculated dates rather than the authored dates |
| `<ConstraintType>2</ConstraintType>` + `<ConstraintDate>` on every non-summary task | Without this "Must Start On" pin, task dates drift when the file is reopened |
| Project header must include `<NewTasksAreManual>1</NewTasksAreManual>` and `<DefaultTaskType>1</DefaultTaskType>` | New tasks added in MS Project default to auto-scheduling, undermining date integrity |
| Duration format `PT{n*8}H0M0S` (working hours), not ISO `P{n}D` | ISO calendar-day duration causes MS Project to miscalculate working-day task lengths |

These rules were discovered empirically through iteration on real project imports. They are now captured both in the script logic and in the `schedule-generation` SKILL.md and the user's persistent memory file — ensuring that any future AI session does not regress to broken XML.

### 5.3 Per-Project Schedule Generator Pattern

Each project has its own `generate_{project}_schedule.py`. This file contains the complete TASKS list (a Python list of tuples with schema: ID, OutlineLevel, Name, Duration, Start, Finish, Predecessors, ResourceNames, Notes, Milestone) plus a `__main__` block that:

1. Writes the CSV output
2. Calls `generate_msp_xml.py` as a subprocess with `--csv`, `--out`, and `--project` arguments
3. Reports start time, finish time, and elapsed seconds to the console

The per-project pattern — rather than a single parametric script — is deliberate. Hardcoded task lists are auditable. A project manager can read `generate_bravo_schedule.py` and see exactly what was generated and why. There is no runtime variable substitution that could produce unexpected output.

---

## 6. Risk Register Generation

Risk registers are generated as `.xlsx` files using `Risk_analysis_template.xlsx` as the base workbook. The template already contains RAG-formula cells (`=$H{row}*$I{row}` in column J) and the authoritative column structure used across all Bosch MIL engagements.

**Column schema:**

| Column | Field | Notes |
|---|---|---|
| A | Risk number | Sequential |
| B | Sub-project | Optional |
| C | Entry date | |
| D | Risk category | ScR, SR, RR, BtR, QR, BR, LR, CR taxonomy |
| E | Risk description | Concrete, engagement-specific |
| F | Effects | Business/schedule consequence if risk materialises |
| G | Causes | Root cause or trigger condition |
| H | Probability (W) | 1–5 scale |
| I | Impact (T) | 1–5 scale |
| J | Risk rating (RZ) | Formula: W × T; threshold 12 = high priority |
| K | Mitigation actions | Preventive and corrective measures |
| L | Owner | Named individual |
| M | Deadline | Target resolution date |
| N | Action status | |
| O | Remarks | Supplementary notes; NOT the same as Effects |

High-priority risks (RZ ≥ 12) are surfaced at script end for immediate review. The cost plan generation skill mandates a cross-check against these high-priority risks: any risk whose mitigation involves external cost must have a corresponding contingency line in the cost plan.

---

## 7. Cost Plan Generation

The cost plan is the third deliverable in the mandatory sequence and can only be generated after both the schedule and risk register are complete. The generation skill enforces three pre-generation cross-checks:

1. **Schedule alignment:** Phase names and date ranges in the cost plan must match the schedule exactly. Resource names in cost lines must map to resource names in the schedule's `Resource Names` column.
2. **Risk register alignment:** Every Amber/Red or RZ ≥ 12 risk whose mitigation involves external spend must have a corresponding CAPEX/contingency line in the cost plan.
3. **Resource name consistency:** Each `+`-separated token in the schedule's resource column must be traceable to at least one cost plan line. No invented resource labels.

The output is a structured CSV with mandatory sections: category blocks with subtotals, an overall project total, cost breakdowns by category and by phase, and a CAPEX/additional costs section (excluded from the labour total) containing risk-driven contingency lines.

---

## 8. HTML Deliverables — Executive Dashboard and Management KPI Dashboard

### 8.1 Design Principles

Both HTML deliverables follow a strict set of non-negotiable design rules:

- **Self-contained HTML** — no external CDN links, no web font requests. The file must render correctly offline (SharePoint, local file system, email).
- **Bosch logo embedded as base64 PNG** — `Bosch.png` from the workspace root is read at generation time, base64-encoded, and injected as a `data:` URI in the `<img>` tag. This eliminates broken-image problems when the file is moved.
- **Blue as primary colour** — `#003b6e` (deep navy) for headers/hero bands, `#0066CC` (Bosch mid-blue) for accents. Bosch Red (`#E20015`) is reserved exclusively for RAG badges and critical-path indicators.
- **Print-safe rules** — `page-break-before:always` on section breaks enables clean PDF printing from browser.

### 8.2 Executive Dashboard — Three-Page A4 Layout

Page 1 contains the programme overview, countdown to key events, phase timeline, key milestones and quality gates table, and budget distribution. Page 2 adds workstream confidence (9 workstreams in a 3×3 card grid), quality gate tracker, regional site distribution, and key risk indicators. Page 3 covers application migration waves, country complexity hotspots, resource statistics, and the critical path.

The canonical layout reference is the Project Trinity Executive Dashboard PDF. All new dashboards replicate this structure with project-specific data.

### 8.3 Management KPI Dashboard

The KPI dashboard is the operational steering view for the PMO and SteerCo. It uses a 12-column CSS grid card layout and renders key performance indicators: Schedule Performance Index (SPI), Cost Performance Index (CPI), Day-1 readiness, Stand Alone / TSA confidence, workstream confidence bars, milestone gate control timeline, top risk table, carve-out model key differences, and a 90-day action forecast.

### 8.4 CSS Container Rule for Logo

A specific CSS bug was discovered and documented during development: using `display:grid` or fixed `width`/`height` on the `.bosch-logo` container clips the Bosch PNG at certain viewport sizes. The correct rule is `display:flex; align-items:center;` — this is enforced in both skill files and the user memory system.

---

## 9. Monthly Status Report Generation

Monthly status reports are PDF documents generated by per-project Python scripts using the `reportlab` library. Each script produces a single-page A4 report with: programme overview band, days-to-gate countdown, phase RAG statuses, risk summary, budget burn, and upcoming actions. The output filename includes the runtime month/year (`{ProjectName}_Monthly_Status_Report_{MMM_YYYY}.pdf`), so the script auto-dates itself on each run — no manual filename editing.

---

## 10. Memory Architecture

### 10.1 Three-Tier Memory

The toolkit uses a three-tier memory system that persists knowledge beyond individual conversation sessions:

| Tier | Scope | Contents |
|---|---|---|
| User memory (`/memories/`) | Cross-workspace, permanent | User preferences, common patterns, learned rules |
| Session memory (`/memories/session/`) | Current conversation | Task-specific context, in-progress state |
| Repository memory (`/memories/repo/`) | This repository | Repository facts, interpreter path, GitHub remote URL |

### 10.2 What Is Stored

The user memory currently captures:
- **Output format standards:** logo embedding rules, colour theme enforcement, schedule format requirements
- **MS Project XML critical rules:** the `<TaskMode>` element ordering rule and all related constraints — the hard-won XML knowledge that took multiple debugging iterations to establish

The repository memory captures:
- **GitHub remote URL** (`https://github.com/erichokh/MIL-Carveout-project`)
- **Canonical branch** (`main`)
- **Python interpreter path** (`C:/Program Files/px/python.exe`)

This memory architecture means that when a new conversation begins — even weeks later — the AI does not rediscover known gotchas. It reads the memory files at the start of each session and applies known-good rules from the outset.

### 10.3 Memory as Institutional Knowledge

The memory system serves a function analogous to onboarding documentation in a traditional team. When a new AI session starts, reading the memory is equivalent to a new team member reading the project handbook. The difference is that the AI's memory is:

- **Precise:** stored as structured rules, not narrative prose
- **Always consulted:** not skipped because the new person is in a hurry
- **Always current:** updated immediately when a new lesson is learned

---

## 11. Generator Script Patterns

### 11.1 Timing Wrapper

All fourteen generator scripts now follow a consistent pattern that records and displays start time, finish time, and elapsed duration:

```python
if __name__ == "__main__":
    from datetime import datetime as _dt
    _t0 = _dt.now()
    print(f"Started : {_t0.strftime('%Y-%m-%d %H:%M:%S')}")
    try:
        main()
    finally:
        _t1 = _dt.now()
        print(f"Finished: {_t1.strftime('%Y-%m-%d %H:%M:%S')}  ({(_t1-_t0).total_seconds():.1f}s elapsed)")
```

This pattern was added uniformly across all scripts in one AI-assisted operation. The `try/finally` ensures that the timing line always prints, even if the generation step raises an exception — so the project manager always knows how long the run took.

### 11.2 Subprocess Chaining

Schedule generators call `generate_msp_xml.py` as a subprocess, passing `--csv`, `--out`, and `--project` arguments:

```python
result = subprocess.run(
    [sys.executable,
     str(HERE / "generate_msp_xml.py"),
     "--csv", str(CSV_PATH),
     "--out", str(XML_PATH),
     "--project", "Project Bravo"],
    capture_output=True, text=True
)
```

This design keeps the XML generation logic centralised in one place (`generate_msp_xml.py`) while allowing each project's schedule script to call it. When the XML generator is improved (e.g. when the `<TaskMode>` ordering rule was identified), the fix is applied once and all project scripts immediately benefit.

---

## 12. Deliverable Orchestration

The deliverable generation sequence is not merely a recommendation — it is enforced at the AI level:

```
1. Schedule (CSV + XML)
2. Risk Register (.xlsx)
3. Cost Plan (.csv)
4. Project Charter (.html)
5. Executive Dashboard (.html)
6. Management KPI Dashboard (.html)
7. Monthly Status Report (.pdf)
```

The ordering exists because each deliverable depends on its predecessors:
- The risk register references schedule phases and quality gates.
- The cost plan cross-references both the schedule (resource assignments) and the risk register (external-cost mitigations).
- The project charter presents the schedule, budget, and risks in a single governance document.
- The dashboards visualise data from the schedule, cost plan, and risk register — they cannot be accurate if those inputs are incomplete.
- The monthly report is a snapshot of all above.

If any mandatory predecessor is missing, the AI blocks generation and reports the dependency gap explicitly.

---

## 13. Version Control and Collaboration

### 13.1 Git Workflow

Generated deliverables and generator scripts are both tracked in Git. The repository lives on both Bosch DevCloud (`github.boschdevcloud.com/kho1sgp/SUaaS`) and the public shared repository (`github.com/erichokh/MIL-Carveout-project`). The standard workflow is:

1. Open workspace in VS Code.
2. Chat with Copilot in Agent mode to generate or update deliverables.
3. Review generated files.
4. `git add . ; git commit -m "<description>" ; git push` to persist the work.

Git provides a complete audit trail: every generated file has a commit history showing when it was created or modified, by whom, and with what AI assistance. For carve-out engagements — which have regulatory and legal implications — this auditability is a significant governance benefit.

### 13.2 Repository Governance

The `repository-governance-updates` skill captures the maintenance responsibility: whenever a new project folder is created, a new generator script is added, a template changes, or a project status changes from active to closed, the repository metadata documents must be updated. This includes the `CLAUDE.md` "Last reviewed" date, the `README.md` project inventory, and the `REPOSITORY_CONTENT_STRUCTURE.md` zone map.

---

## 14. Observed Outcomes and Lessons

### 14.1 Speed

A complete deliverable set (schedule + risk register + cost plan + charter + two dashboards + status report) for a new engagement — which previously required several days of manual work — is now generated in minutes. The primary time cost is now in reviewing the output for engagement-specific accuracy, not in document construction.

### 14.2 Consistency

Because the AI always reads the skill files before generating, and the skill files encode the canonical layouts and schemas, outputs are structurally identical across projects. A stakeholder familiar with the Project AlphaX charter immediately understands the Project Bravo charter — same sections, same visual language, same quality gate structure.

### 14.3 Cross-Check Enforcement

The mandatory cross-checks (schedule ↔ cost plan ↔ risk register) are enforced by the AI's instruction layer, not by human memory. In practice, this has caught real inconsistencies: resource names that changed between schedule and cost plan, phase dates that were updated in the schedule but not reflected in the dashboard.

### 14.4 Institutional Memory Accumulation

Each time a new lesson is learned — the `<TaskMode>` XML rule, the `.bosch-logo` CSS grid clipping bug, the `.xlsx` vs `.xls` openpyxl restriction — it is written into the memory system and enforced from that point forward. The system gets more reliable with each engagement, rather than regressing to known bugs.

### 14.5 Limitations

- **Data accuracy is the PM's responsibility:** The AI generates structure and formatting from given facts. If the scope facts supplied in the intake are wrong (wrong site count, wrong user number), the deliverables will be internally consistent but externally incorrect.
- **Complex Excel formulas:** The `openpyxl` library writes cell values but cannot execute Excel formulas at generation time. The risk register formula column (`=$H{row}*$I{row}`) is written as a string; Excel calculates it on file open.
- **AI session restarts:** Each Copilot chat session begins fresh. The memory architecture mitigates this for known rules, but a very long multi-step session may benefit from periodic memory reads to refresh working context.

---

## 15. Future Directions

Several extensions to the toolkit are under consideration:

- **Dashboard auto-update:** A script that reads actual progress data (from a project's status source) and regenerates the dashboards automatically, without AI intervention, for routine weekly refreshes.
- **Deliverable diff reporting:** A tool that compares two versions of a schedule or cost plan and summarises what changed — useful for change-control presentations.
- **Multi-project view:** An aggregated portfolio dashboard showing status across all active engagements, drawing from the individual project folders.
- **Natural language status entry:** A Copilot chat flow where the PM types a brief status update in natural language and the AI updates the relevant fields in the monthly report and dashboards.
- **Template versioning:** Formal versioning of `Risk_analysis_template.xlsx` and the dashboard layout spec, with the AI instructed to use the specific version tagged at engagement start.

---

## 16. Conclusion

The Carveout AI Toolkit demonstrates that AI-assisted project management, when properly structured, can go well beyond text generation. By combining a domain-specific instruction layer, modular skill files that encode methodology, an enforced deliverable sequence, cross-check rules, and a memory architecture that accumulates institutional knowledge, the system produces a complete set of governance-quality deliverables with high structural consistency and low manual effort.

The key insight is architectural: the value of the AI is not in the sophistication of a single prompt but in the layer of structured knowledge (skills, instructions, memory) that wraps and guides its inference. Without that layer, the AI produces plausible output. With it, the AI produces correct, consistent, cross-checked, audit-ready output.

For IT carve-out practitioners, this translates to a practical shift: project managers can focus their energy on engagement-specific judgment — which risks matter, what the critical path really is, how to communicate with the SteerCo — rather than on document construction. The construction is automated. The judgment remains human.

---

## Appendix A — Mandatory Engagement Intake Fields

| Field | Description | Example |
|---|---|---|
| Project name | Engagement codename | Project Hamburger |
| Seller | Entity divesting the business | Robert Bosch GmbH |
| Buyer | Entity acquiring the business | Undisclosed Buyer |
| Business carved out | The specific business/division | Solar Energy Business Unit |
| Carve-out model | Stand Alone / Integration / Combination | Stand Alone |
| PMO lead | Named methodology lead | Erik Ho (BD/MIL-ICC) |
| Worldwide sites | Number of locations in scope | 17 |
| IT users | Target population count | 2,600 |
| Project start date | Phase 0 kickoff date | 01 April 2026 |
| GoLive / Day-1 date | Operational cutover date | 01 December 2026 |
| Project completion date | QG5 / programme closure | 30 May 2027 |

---

## Appendix B — Deliverable Generation Sequence

```
Intake Compliance Gate
        │
        ▼
  1. Schedule  ─────────────────────┐
        │                           │
        ▼                           │
  2. Risk Register ─────────────────┤
        │                           │
        ▼                           ▼
  3. Cost Plan  (requires 1 + 2) ───┤
        │                           │
        ▼                           │
  4. Project Charter (requires 3)   │
        │                           ▼
        ▼                    5. Executive Dashboard
  6. Management KPI Dashboard  (requires 1+2+3)
        │
        ▼
  7. Monthly Status Report (requires 1–6)
```

---

## Appendix C — Generator Scripts Inventory

| Script | Project | Produces |
|---|---|---|
| `generate_msp_xml.py` | All (shared) | `{project}.xml` from any CSV |
| `generate_alphax_schedule.py` | AlphaX | CSV + XML |
| `generate_bravo_schedule.py` | Bravo | CSV + XML |
| `generate_falcon_schedule.py` | Falcon | CSV + XML |
| `generate_hamburger_schedule.py` | Hamburger | CSV + XML |
| `generate_bravo_risk_register.py` | Bravo | Risk register `.xlsx` |
| `generate_falcon_risk_register.py` | Falcon | Risk register `.xlsx` |
| `generate_hamburger_risk_register.py` | Hamburger | Risk register `.xlsx` |
| `generate_risk_register.py` | Trinity (reference) | Risk register `.xlsx` |
| `generate_bravo_charter.py` | Bravo | Project charter `.html` |
| `generate_project_charter.py` | Multi-project | Project charter `.html` |
| `generate_alphax_monthly_report.py` | AlphaX | Status report `.pdf` |
| `generate_hamburger_monthly_report.py` | Hamburger | Status report `.pdf` |
| `generate_carveout_schedule_direct_xml.py` | Legacy | Direct XML generation |

---

*End of paper.*
