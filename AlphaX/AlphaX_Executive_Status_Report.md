# Project AlphaX — Executive Status Report
**Reporting Period:** April 2026 (Programme Initiation)
**Document Classification:** CONFIDENTIAL — SteerCo Distribution Only
**Report Date:** 1 April 2026

---

## 1. Programme Identity

| Field | Value |
|---|---|
| **Programme Name** | Project AlphaX |
| **Programme Description** | Full IT carve-out of Bosch Battery Division into standalone NewCo legal entity |
| **Seller / Sponsor Contractor** | Robert Bosch GmbH |
| **Buyer / Sponsor Customer** | NewCo Battery Division (permanent buyer in sourcing) |
| **Carve-Out Model** | Stand Alone — Full Separation (NO Merger Zone) |
| **Delivery Lead** | KPMG (Methodology) |
| **Sites in Scope** | 35 Battery sites |
| **Users in Scope** | ~3,000 |
| **Applications in Scope** | ~200 |
| **Programme Start** | 1 June 2026 |
| **Day 1 / GoLive** | 1 October 2027 |
| **Programme Closure** | 31 December 2027 |
| **Total Duration** | 19 months |

---

## 2. Overall Programme Health

| Dimension | Status | Trend | Comment |
|---|---|---|---|
| **Overall Health** | 🟡 AMBER | ↓ | Legal entity uncertainty and compressed timeline are primary concerns |
| **Schedule** | 🟡 AMBER | → | SPI 0.91; Phase 1 concept documents 6 days behind |
| **Cost** | 🟢 GREEN | ↑ | CPI 1.04; running within approved cost envelope |
| **Scope** | 🟡 AMBER | → | Battery app inventory not yet finalised; OT scope under review |
| **Quality** | 🟢 GREEN | → | No P1/P2 defects; quality gates not yet active |
| **Risks & Issues** | 🔴 RED | ↓ | 7 High risks open; R1 (legal entity) is critical path blocker |
| **Resources** | 🟡 AMBER | → | Bosch IT team onboarding; KPMG fully mobilised |
| **TSA / Separation** | 🟡 AMBER | → | Bosch TSA framework draft in progress; exit criteria TBD |

---

## 3. Key Achievements This Period

- Programme kick-off governance structure agreed; Steering Committee members confirmed from both Bosch and KPMG sides.
- Phase 1 workplan issued across all 9 IT sub-workstreams with workstream leads assigned.
- Battery Division IT landscape initial scoping analysis commenced — CMDB discovery run initiated for 10 priority sites.
- NewCo legal entity interim pathway defined; parallel registration approach agreed with Legal.
- Technology options assessment for ERP (greenfield vs shell copy) scoped — decision workshop scheduled for June 2026.
- AlphaX project repository, SharePoint site, and ITSM project space established.

---

## 4. Phase Status Summary

| Phase | Name | Dates | Status | % Complete |
|---|---|---|---|---|
| Phase 1 | Initialization | Jun – Aug 2026 | 🟡 In Progress | 8% |
| Phase 2 | Concept & Architecture | Sep – Nov 2026 | ⚪ Not Started | 0% |
| Phase 3 | Development & Build | Dec 2026 – Jun 2027 | ⚪ Not Started | 0% |
| Phase 4 | Implementation & Testing | Jul – Sep 2027 | ⚪ Not Started | 0% |
| Phase 5 | GoLive & Hypercare | Oct – Dec 2027 | ⚪ Not Started | 0% |

---

## 5. Workstream Health Snapshot

| # | Workstream | Lead | Confidence | Status | Priority Actions |
|---|---|---|---|---|---|
| WS1 | IT Infrastructure | Bosch Infra | 66% 🟡 | On Track | WAN vendor pre-qualification by end Phase 1; co-lo RFP issued |
| WS2 | ERP / Commercial IT | Bosch SAP + KPMG | 55% 🔴 | At Risk | ERP strategy decision required by Jun 2026 — critical gate |
| WS3 | Applications (~200) | App Leads + KPMG | 74% 🟡 | On Track | Full inventory freeze mandate; LeanIX scan in progress |
| WS4 | Engineering IT | Eng IT Lead | 78% 🟢 | On Track | PLM separation design started |
| WS5 | Production IT / OT | OT Lead | 64% 🟡 | On Track | OT site assessment for top 12 Battery manufacturing sites |
| WS6 | HR IT | HR IT Lead | 80% 🟢 | On Track | Personnel data separation design in progress |
| WS7 | IT Org & Processes | IT PM + KPMG | 72% 🟡 | On Track | NewCo ITSM platform selection underway |
| WS8 | Contracts & Licenses | Legal + SAM | 60% 🟡 | At Risk | SAM audit of ~200 app licences not yet started; 890+ contracts to review |
| WS9 | IT Security / IAM | CISO + IAM Team | 57% 🔴 | At Risk | NewCo IAM platform selection not started; ISO 27001 baseline pre-assessment pending |

---

## 6. Top Risks

| ID | Risk | Prob | Impact | RZ | Owner | Action Required |
|---|---|---|---|---|---|---|
| R1 | NewCo legal entity setup delayed — permanent buyer not confirmed | 4 | 4 | **16** 🔴 | Legal + Finance | Confirm interim NewCo registration pathway; brief SteerCo Q2 |
| R2 | WAN ordering for 35 sites delayed — 4-6 month lead time | 3 | 5 | **15** 🔴 | Bosch Procurement | Pre-qualify WAN vendors immediately; mandate QG2 ordering |
| R3 | Battery IT application inventory incomplete (~200 apps) | 4 | 3 | **12** 🟡 | KPMG + IT Architects | LeanIX scan; app owner mandate by Week 8 |
| R4 | SAP ERP strategy (greenfield vs shell copy) undecided | 3 | 4 | **12** 🟡 | ERP Architects | Decision workshop Jun 2026; KPMG benchmarks issued |
| R5 | 19-month programme timeline compressed for Stand Alone model | 4 | 3 | **12** 🟡 | IT PM | Parallel workstream execution; resource loading confirmed each QG |
| R6 | NewCo IAM / Identity provider not operational on Day 1 | 3 | 4 | **12** 🟡 | Bosch CISO + NewCo | Platform selection by QG2; no dependency on Bosch IAM from Day 1 |
| R7 | Bosch TSA staff attrition post-separation announcement | 3 | 3 | **9** 🟡 | HR + IT PM | Retention clauses; KT programme from Day 1 |

---

## 7. Open Issues Requiring SteerCo Decision

| # | Issue | Raised | Decision Needed By | Impacted Workstreams |
|---|---|---|---|---|
| I-01 | **ERP strategy unresolved** — greenfield vs SAP shell copy: cost delta is significant; both options require vendor selection | 1 Apr 2026 | 30 Jun 2026 (QG1) | WS2, WS7 |
| I-02 | **NewCo legal entity** — interim entity path approved, but permanent buyer identification must be tracked; architecture decisions may need revisiting | 1 Apr 2026 | 31 Aug 2026 (QG1) | All workstreams |
| I-03 | **WAN vendor pre-qualification** — must begin now to hit QG2 order trigger; procurement approval required | 1 Apr 2026 | 15 May 2026 | WS1 |
| I-04 | **IT Security CISO appointment for NewCo** — no interim CISO confirmed; ISO 27001 baseline is blocked | 1 Apr 2026 | 30 Jun 2026 | WS9 |

---

## 8. Upcoming Milestones (Next 90 Days)

| Date | Milestone | Gate |
|---|---|---|
| 15 May 2026 | WAN vendor pre-qualification complete; procurement approval for RFP | — |
| 1 Jun 2026 | Programme formal start — Deal Closing | P0 |
| 15 Jun 2026 | Signing; Frozen Zone activated for all Battery IT systems | QG0 |
| 30 Jun 2026 | ERP strategy decision (greenfield vs shell copy) confirmed by SteerCo | — |
| 31 Aug 2026 | QG1 — Initialization complete; full app/site/user inventory; TSA draft; governance approved | **QG1** |

---

## 9. Key Decisions Made

| Date | Decision | Made By |
|---|---|---|
| Apr 2026 | Carve-out model confirmed: Stand Alone (no Merger Zone) — Battery IT to NewCo direct, no hybrid transition | SteerCo |
| Apr 2026 | Programme duration: Jun 2026 – Dec 2027 (19 months) — within Board mandate | SteerCo |
| Apr 2026 | Project name confirmed: AlphaX | Programme Sponsor |
| Apr 2026 | NewCo interim parallel legal entity pathway approved while buyer sourcing continues | Legal |

---

## 10. Financial Snapshot

| Category | Approved Budget | Committed to Date | Forecast Cost at Completion | Variance |
|---|---|---|---|---|
| KPMG Consulting (Phases 1–4) | TBC at QG1 | — | TBC | — |
| Infrastructure (WAN, co-lo, servers) | TBC at QG2 | — | TBC | — |
| ERP / SAP Migration | TBC at QG2 | — | TBC | — |
| Applications & Licensing | TBC at QG1 | — | TBC | — |
| Programme PMO | TBC at QG1 | — | TBC | — |
| **Total Programme** | **TBC – QG1 gate** | **Pre-spend minimal** | **TBC** | **TBC** |

> Note: Full programme budget baseline will be approved at QG1 (31 Aug 2026) following completion of the Phase 1 scope and architecture assessment.

---

## 11. TSA Framework Summary

Since this is a Stand Alone model with no Merger Zone, Bosch's TSA obligations are the sole bridge for NewCo operations until full separation:

| TSA Coverage Area | Responsible Provider | Planned Exit |
|---|---|---|
| WAN / network access to Battery sites | Bosch IT (Infra) | Per site cutover in Phase 4 waves |
| SAP / ERP transactional services | Bosch SAP team | Day 1 (Oct 2027) or per-system plan |
| M365 / email / Teams | Bosch Digital Workplace | Day 1 new NewCo M365 tenant |
| Security operations (SOC/SIEM) | Bosch CISO | Handover to NewCo SOC by Day 1 |
| ServiceDesk / ITSM | Bosch IT Ops | NewCo MSP in place by QG3 |
| Application support (~200 apps) | Bosch App teams | Per-app exit per wave plan |

> **Key TSA principle for AlphaX:** Bosch TSA is the **only** backstop — there is no Merger Zone safety net. Every TSA service must have a defined exit date and exit criteria, agreed and documented before QG2.

---

*Project AlphaX · Executive Status Report v1.0 · April 2026 · Prepared by AlphaX PMO / KPMG*
*CONFIDENTIAL – For Steering Committee Distribution Only*
