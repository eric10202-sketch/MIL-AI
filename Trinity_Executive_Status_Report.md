# PROJECT TRINITY — EXECUTIVE STATUS REPORT
## IT Carve-Out: Bosch → JCI (Keenfinity)

**Report Date:** March 28, 2026  
**Project Status:** ⚠️ **PRE-LAUNCH (Awaiting Go-Live Authorization)**  
**Overall Health:** ON TRACK  

---

## EXECUTIVE SUMMARY

Project Trinity is a full IT carve-out separating Bosch's independent operations business to Keenfinity (JCI). The project scope encompasses **8,000 employees** across **180 sites** in 3 regions (AP/AM/EMEA), with ~**500 applications** and **6,000 client devices** to separate and migrate.

**Project Timeline:**  
- **Kick-Off:** July 1, 2026  
- **Closing Target:** Q2 2028 (18-month TSA + 3-month hypercare)  
- **Total Duration:** 24 months (Initialization → Stabilization)

---

## KEY METRICS AT A GLANCE

| Metric | Current Status | Target | Variance |
|--------|---|---|---|
| **Project Phases Planned** | 6 phases, 192 tasks | ✓ | 0 days |
| **Governance Model** | Steering Committee + PMO draft | Approved at QG1 | On schedule |
| **Budget (Labour Only)** | €5.1M (draft cost plan) | Awaiting CFO sign-off | Under review |
| **Resource Capacity** | 45+ FTE allocated | Per schedule | Confirmed |
| **Risk Register** | 25 legacy risks (FRAME baseline) | Mitigation plans in progress | Manageable |
| **Contracts & TSA** | TSA framework drafted (18 months) | Legal review underway | On track |

---

## PHASE STATUS

### ✅ Phase 1 — INITIALIZATION (Jul 2026 – Sep 2026)
**Objective:** Establish governance, define scope, and baseline IT landscape.
- **Key Deliverables:** Project Charter, Governance Model, IT Inventory, Site Mapping (180 sites)
- **Quality Gate 1:** Sep 30, 2026 (PENDING: All governance and initial assessment sign-offs)

### ⏳ Phase 2 — CONCEPT (Oct 2026 – Mar 2027)
**Objective:** Design Merger Zone architecture, finalize migration strategy, and plan cutover.
- **Key Deliverables:** IT architecture design, application wave plan, ERP migration strategy, TSA service catalog
- **Quality Gate 2:** Mar 31, 2027 (Detailed project plan approved)

### 📋 Phase 3 — DEVELOPMENT (Apr 2027 – Jul 2027)
**Objective:** Build Merger Zone environments and prepare migration tooling.
- **Key Deliverables:** Infrastructure build (3 regional DCs), AD setup, M365 tenant, SAP shell copy, CMDB configuration

### 🚀 Phase 4 — IMPLEMENTATION (Aug 2027 – Oct 2027)
**Objective:** Execute migrations in waves; activate Merger Zone services.
- **Key Deliverables:** WAN cutover, file share migration, application migration, client re-imaging

### 🎯 Phase 5 — GO-LIVE (Nov 2027)
**Objective:** Big Bang cutover to independent carve-out; Day 1 operations.
- **Critical Path:** SPA/APA signing → TSA activation → Day 1 go-live → Hypercare (3 months)

### ✓ Phase 6 — STABILIZATION (Dec 2027 – Feb 2028)
**Objective:** Hypercare operations, performance monitoring, TSA exit planning.

---

## TOP 5 RISKS & MITIGATIONS

| # | Risk | Probability | Impact | Rating | Mitigation Strategy |
|---|------|---|---|---|---|
| **R1** | **Schedule Compression** – WAN lead time (4–6 mo) + SAP complexity causes critical path delay | **HIGH** | **HIGH** | **25** | Order WAN by Month 4; run SAP shell copy in parallel; agile sprint approach for UAT |
| **R2** | **Scope Creep** – 500+ applications difficult to categorize; app owners request exceptions | **MEDIUM** | **HIGH** | **20** | Lock application inventory at QG1; wave-based prioritization; freeze zone from Month 12 |
| **R3** | **Resource Availability** – Bosch IT and JCI IT competing for same resources during TSA handover | **HIGH** | **MEDIUM** | **15** | Hybrid staffing model; cross-training from Month 4; hire external SME contractors early |
| **R4** | **Country-Specific Complexity** – Brazil tax ERP rules, China FTS/customs, India local IdM policies slow deployment | **MEDIUM** | **HIGH** | **20** | Engage regional legal/finance teams Month 2; pre-stage regional governance boards; India RBIN alternatives identified |
| **R5** | **Data Separation Errors** – Non-selective migration approach risks mixing Bosch/JCI operational data | **MEDIUM** | **CRITICAL** | **25** | Finalize separation rules at QG1; third-party audit pre-cutover; manual spot checks on 10% of records |

---

## BUDGET & RESOURCE SNAPSHOT

### Estimated Labour Cost (Full Project)
- **Total:** €5.1M (fully-loaded labour; excludes hardware, software, travel, vendor fees)
- **Breakdown:**
  - Governance/PMO: €633.6K
  - Bosch IT Team: €1.95M
  - JCI IT Team: €1.2M
  - KPMG Consulting: €1.3M
  
### Staffing Model
- **Peak FTE:** ~45 FTE during Phases 3–4 (Development & Implementation)
- **Teams Leading:** Bosch IT (Infra, ERP, Security) + JCI IT (TSA operations) + KPMG (methodology & workstream lead)

---

## CRITICAL DEPENDENCIES & MILESTONES

| Date | Milestone | Dependencies / Gate |
|------|-----------|---|
| **Jul 1, 2026** | Project Kick-Off + Team Mobilization | Executive sign-off on charter |
| **Sep 30, 2026** | ⚠️ **QG1 — Initialization Sign-Off** | Governance approved; 180-site inventory complete; scope frozen |
| **Dec 1, 2026** | Merger Zone Architecture Design Complete | As-is analysis finalized; decision on Stand Alone carve-out model |
| **Jan 13, 2027** | Co-Locator / Data Center Selection Finalized | Regional hub model (AP/AM/EMEA) confirmed; ordering begins |
| **Feb 19, 2027** | Architecture Design Complete (WAN/AD/M365/IAM) | All design reviews and approvals completed |
| **Mar 31, 2027** | **QG2 — Concept Phase Sign-Off** | Migration strategy approved; wave plan locked; cutover plan drafted |
| **Aug 1, 2027** | Infrastructure Build Complete (3 Regional DCs) | WAN ordered (critical: 4–6 month lead) |
| **Oct 31, 2027** | All Migrations Complete to Merger Zone | Applications migrated; AD users converged; file shares live |
| **Nov 1, 2027** | **🎯 DAY 1 GO-LIVE** | SPA/APA signed; TSA activated; Bosch operates independently |
| **Feb 28, 2028** | **Project Closure** | Hypercare complete; TSA exit initiated; lessons learned documented |

---

## STEERING COMMITTEE DECISIONS REQUIRED (NEXT 30 DAYS)

### 🔴 IMMEDIATE (Apr 2026)
1. **Approve Project Charter & Governance Model** – Draft completed; awaiting steering sign-off
2. **Confirm Budget & Resource Allocation** – €5.1M labour cost + hardware/software procurement budget TBD
3. **Authorize KPMG Engagement** – Consulting partner onboarding; contract finalization

### 🟡 WITHIN 60 DAYS (May 2026)
4. **Carve-Out Model Decision** – Stand Alone vs. Integration with Buyer (assumed Stand Alone; confirm)
5. **TSA Duration & Service Scope** – 18-month TSA framework drafted; finalize scope and exit criteria
6. **Regional Hub Architecture** – Approve 3-hub model (AP/AM/EMEA co-locators)

---

## SUCCESS CRITERIA (PROJECT CLOSURE — FEB 2028)

- ✓ All 8,000 employees transitioned to independent JCI operations on Day 1
- ✓ $0 (zero) unplanned downtime on critical systems (ERP, email, network) during go-live
- ✓ 95%+ of applications successfully migrated or decommissioned
- ✓ All contracts & licenses transferred or renegotiated (no legal blockers)
- ✓ 18-month TSA exited with <2% Bosch-to-JCI escalations over 3-month hypercare
- ✓ Project closure: on schedule, within budget, all quality gates passed

---

## NEXT STEPS

| Owner | Action | Due Date |
|-------|--------|------|
| **Steering Committee** | Approve Project Charter, Budget, Governance | Apr 15, 2026 |
| **Bosch IT PM + KPMG** | Finalize TSA Framework & Service Catalog | Apr 30, 2026 |
| **Procurement** | Issue RFP for WAN, co-locators, consulting | May 1, 2026 |
| **Legal** | Finalize carve-out contracts & SPA structure | May 15, 2026 |
| **All IT Workstreams** | Mobilize teams; begin detailed planning | Jul 1, 2026 (Kick-Off) |

---

**Report Prepared By:** Project Trinity Leadership Team  
**Distribution:** Bosch Executive Leadership, JCI Board, Steering Committee, IT Leadership  
**Confidentiality:** Internal Use Only
