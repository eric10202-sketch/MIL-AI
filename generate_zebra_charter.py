#!/usr/bin/env python3
"""
Generate Zebra_Project_Charter.html

Project: Zebra (Packaging carve-out)
Seller: Robert Bosch GmbH
Buyer: Undisclosed (legal hold)
Business: Packaging
Model: Stand Alone
Scope: 37 sites, 3500+ users, 208 applications, SAP + TSA
Timeline: 1 April 2026 - 31 October 2027
PM: Gill Amandeep Singh (BD/MIL-PSM1)
"""

import base64
from datetime import datetime as _dt
from pathlib import Path

_t0 = _dt.now()
print(f"Started : {_t0.strftime('%Y-%m-%d %H:%M:%S')}")

HERE = Path(__file__).parent
OUT  = HERE / "active-projects" / "Zebra" / "Zebra_Project_Charter.html"

logo_b64 = base64.b64encode((HERE / "Bosch.png").read_bytes()).decode()

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Project Zebra – Project Charter</title>
  <style>
    :root {{
      --ink: #0f1923; --muted: #5a6478; --line: #d8dde8;
      --bg: #f0f3f8;  --card: #ffffff;
      --brand: #003b6e; --brand-lt: #e4edf9;
      --good: #007A33; --warn: #E8A000; --risk: #CC0000;
    }}
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
      font-family: "Segoe UI", system-ui, -apple-system, Arial, sans-serif;
      background: var(--bg); color: var(--ink);
      font-size: 13.5px; line-height: 1.65;
    }}
    .page {{ max-width: 1020px; margin: 0 auto; padding: 32px 28px 64px; }}
    .cover {{
      background: linear-gradient(145deg, #001f45 0%, #003b6e 60%, #0066CC 100%);
      border-radius: 18px; color: #fff;
      padding: 40px 44px 36px; margin-bottom: 28px;
    }}
    .bosch-logo {{
      display: flex; align-items: center;
      background: #fff; padding: 4px 8px; border-radius: 4px;
      width: fit-content; margin-bottom: 14px;
    }}
    .cover h1 {{ font-size: 34px; font-weight: 700; margin-bottom: 8px; }}
    .cover p  {{ font-size: 15px; opacity: 0.9; margin-bottom: 20px; }}
    .cover-meta {{
      display: grid; grid-template-columns: 1fr 1fr; gap: 24px;
      margin-top: 28px; padding-top: 24px;
      border-top: 1px solid rgba(255,255,255,0.2);
    }}
    .meta-item  {{ font-size: 12px; opacity: 0.85; }}
    .meta-label {{ font-weight: 700; display: block; margin-bottom: 4px; }}
    .section {{
      background: var(--card); border-radius: 12px;
      padding: 24px; margin-bottom: 16px;
    }}
    .section h2 {{
      font-size: 18px; font-weight: 700; color: var(--brand);
      margin-bottom: 14px;
      border-bottom: 2px solid var(--brand-lt); padding-bottom: 10px;
    }}
    .section h3 {{
      font-size: 14px; font-weight: 700;
      color: var(--ink); margin-top: 16px; margin-bottom: 8px;
    }}
    ul, ol {{ margin-left: 20px; margin-bottom: 12px; }}
    li {{ margin-bottom: 6px; }}
    .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }}
    .box {{
      background: var(--brand-lt);
      border-left: 4px solid var(--brand);
      padding: 14px; border-radius: 6px; font-size: 12px;
    }}
    .box-title {{ font-weight: 700; color: var(--brand); margin-bottom: 6px; }}
    .highlight {{
      background: #fff8e1;
      border-left: 4px solid var(--warn);
      padding: 12px 14px; border-radius: 6px;
      font-size: 12px; margin-top: 12px;
    }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 12px; }}
    th, td {{ text-align: left; padding: 10px; border-bottom: 1px solid var(--line); font-size: 12px; }}
    th {{ background: var(--brand-lt); font-weight: 700; color: var(--brand); }}
    .pill {{
      display: inline-block; padding: 2px 10px; border-radius: 12px;
      font-size: 11px; font-weight: 700; color: #fff;
    }}
    .green  {{ background: var(--good); }}
    .amber  {{ background: var(--warn); color: #1a1a1a; }}
    .red    {{ background: var(--risk); }}
    .footer {{
      background: var(--brand-lt); border-radius: 12px;
      padding: 16px; margin-top: 24px;
      font-size: 11px; color: var(--muted); text-align: center;
    }}
    p {{ margin-bottom: 10px; }}
  </style>
</head>
<body>
  <div class="page">
    <!-- COVER -->
    <div class="cover">
      <div class="bosch-logo">
        <img src="data:image/png;base64,{logo_b64}" alt="Bosch — Invented for Life" style="height:36px;display:block;" />
      </div>
      <div style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;margin-bottom:12px;opacity:0.9;">Project Charter</div>
      <h1>Project Zebra</h1>
      <p>Packaging Business Carve-Out | Global Multi-Site IT Separation | Stand Alone Model with TSA</p>
      <div class="cover-meta">
        <div class="meta-item">
          <span class="meta-label">Project Duration</span>
          01 Apr 2026 – 31 Oct 2027 (20 months)
        </div>
        <div class="meta-item">
          <span class="meta-label">Scope</span>
          37 Worldwide Sites · 3500+ Users · 208 Apps · SAP Included · TSA Required
        </div>
        <div class="meta-item">
          <span class="meta-label">Carve-Out Model</span>
          Stand Alone (Separation + Handover to Buyer)
        </div>
        <div class="meta-item">
          <span class="meta-label">Budget Baseline</span>
          EUR 2.34M (Labour + CAPEX) – to be approved at QG0
        </div>
      </div>
    </div>

    <!-- PROJECT OVERVIEW -->
    <div class="section">
      <h2>Project Overview</h2>
      <p><strong>Objective:</strong> Execute the complete separation and carve-out of the Packaging business from Robert Bosch GmbH's global IT infrastructure and hand over full operational control to the Buyer (identity confidential pending legal clearance). The Packaging business operates across 37 worldwide sites with 3500+ IT users, 208 integrated applications including SAP ERP, complex data dependencies, and multi-geography supply-chain systems. Full separation, data migration, system integration, and operational handover must be completed by 1 June 2027 (GoLive Day 1), followed by a 4-month post-GoLive TSA and 4-month hypercare closure phase.</p>
      <p><strong>Strategic Context:</strong> This is a large-scale Stand Alone carve-out involving complex multi-geography operations, significant ERP and application portfolio scope, and mandatory TSA-period separation phasing. The separation model requires complete IT independence from Seller systems by GoLive, with Seller providing time-bounded Transition Services Agreement (TSA) support for up to 6 months post-GoLive. This is a high-complexity project due to global site distribution, application portfolio scale (208 apps), SAP landscape complexity, and buyer infrastructure readiness challenges.</p>
      <p><strong>Sponsor Customer (Buyer):</strong> [Confidential — legal hold pending disclosure approval]</p>
      <p><strong>Sponsor Contractor (Seller):</strong> Robert Bosch GmbH</p>
      <p><strong>Programme Manager:</strong> Gill Amandeep Singh (BD/MIL-PSM1)</p>
      <p><strong>PMO Partner:</strong> KPMG</p>
    </div>

    <!-- SCOPE & MODEL -->
    <div class="section">
      <h2>Scope &amp; Model</h2>
      <div class="grid-2">
        <div>
          <h3>In Scope</h3>
          <ul>
            <li>208 integrated applications (retain, retire, or transition)</li>
            <li>SAP ERP system (Packaging-specific instance or separation)</li>
            <li>3500+ IT users across 37 worldwide sites</li>
            <li>Global network separation and internet interconnection</li>
            <li>Application data migration and cutover (15+ years historical data)</li>
            <li>Packaging master data (materials, orders, supply chains)</li>
            <li>Security and IAM platform separation</li>
            <li>Infrastructure (servers, storage, backup) capacity planning</li>
            <li>Global data residency and GDPR compliance mapping</li>
            <li>Transition Services Agreement (TSA) definition and execution</li>
          </ul>
        </div>
        <div>
          <h3>Out of Scope</h3>
          <ul>
            <li>Seller's other business units or global operations</li>
            <li>Manufacturing / OT systems (unless Packaging-critical)</li>
            <li>Post-TSA vendor management (buyer responsibility)</li>
            <li>Buyer's internal IT integration beyond Packaging handover</li>
            <li>Financial restatement or legal entity setup (finance/legal handles)</li>
            <li>Customer communications or brand transition</li>
            <li>Antitrust or M&amp;A advisory (buyer/seller legal handles)</li>
          </ul>
        </div>
      </div>
      <div class="box" style="margin-top:16px;">
        <div class="box-title">Carve-Out Model: Stand Alone with TSA</div>
        <strong>Structure:</strong> Complete Packaging IT separation from Bosch. Buyer assumes full operational control at GoLive. <strong>Stand Alone Characteristics:</strong> No merger zone, no shared infrastructure retained post-GoLive; Buyer operates independently after TSA closure. <strong>TSA Scope:</strong> Seller provides transition IT services (e.g., legacy system access, incident support, knowledge transfer) for up to 6 months post-GoLive or until buyer achieves operational readiness; exit criteria and service levels formally defined in TSA contract.
      </div>
      <div class="highlight">
        <strong>Complexity Drivers:</strong> 37-site global footprint, 208 applications, SAP integration, data residency regulations, and buyer infrastructure readiness create significant execution risk. Parallel cutover across multiple geographies must be coordinated, tested, and backed by robust rollback plans.
      </div>
    </div>

    <!-- TIMELINE & MILESTONES -->
    <div class="section">
      <h2>Timeline &amp; Key Milestones</h2>
      <table>
        <tr>
          <th>Phase</th>
          <th>Window</th>
          <th>Key Deliverable</th>
          <th>Gate</th>
          <th>Date</th>
        </tr>
        <tr>
          <td>0 – Initialization</td>
          <td>01 Apr – 17 Apr 2026</td>
          <td>Charter signed; governance active; risk baseline; legal alignment</td>
          <td><span class="pill green">QG0</span></td>
          <td>17 Apr 2026</td>
        </tr>
        <tr>
          <td>1 – Concept</td>
          <td>18 Apr – 18 May 2026</td>
          <td>Requirements; AS-IS app landscape; TSA terms; SAP roadmap; cutover strategy draft</td>
          <td><span class="pill green">QG1</span></td>
          <td>18 May 2026</td>
        </tr>
        <tr>
          <td>2 – Design &amp; Architecture</td>
          <td>19 May – 27 Jul 2026</td>
          <td>SAP separation design; data migration plan; infrastructure design; testing strategy; cutover playbook</td>
          <td><span class="pill green">QG2/3</span></td>
          <td>27 Jul 2026</td>
        </tr>
        <tr>
          <td>3 – Development &amp; Build &amp; Test</td>
          <td>28 Jul 2026 – 27 May 2027</td>
          <td>SAP system copy &amp; build; data migration (3 dry runs); app integration; UAT across 37 sites; dress rehearsal</td>
          <td><em>QG4 gate review</em></td>
          <td>27 May 2027</td>
        </tr>
        <tr>
          <td>4 – GoLive &amp; Closure</td>
          <td>28 May – 31 Oct 2027</td>
          <td>Day 1 GoLive (1 Jun 2027); parallel cutover; hypercare (4 months); TSA exit (6 months); programme closure</td>
          <td><span class="pill green">QG4 + QG5</span></td>
          <td>31 Oct 2027</td>
        </tr>
      </table>
      <p style="margin-top:12px;"><strong>Key Milestone:</strong> <span class="pill green">GoLive Day 1 = 1 June 2027</span> — All 3500+ users and 208 applications transitioned to buyer infrastructure; Seller TSA support commences. <span class="pill green">QG4</span> (pre-GoLive gate) must pass 3 days before GoLive to authorize cutover.</p>
    </div>

    <!-- KEY SUCCESS CRITERIA -->
    <div class="section">
      <h2>Key Success Criteria</h2>
      <ul>
        <li><strong>GoLive Day 1 (1 June 2027):</strong> All 3500+ users and 208 applications fully operational on buyer infrastructure; zero critical P1 open issues blocking user access.</li>
        <li><strong>Data Integrity:</strong> 100% Packaging master data (15+ years) migrated with validation; no data loss; post-cutover reconciliation closed by Day 10 ("hypercare checkpoint").</li>
        <li><strong>SAP Functionality:</strong> Packaging SAP business processes (procurement, inventory, orders, costing) execute correctly on buyer systems with zero financial impact.</li>
        <li><strong>Global Site Readiness:</strong> All 37 sites pass pre-cutover readiness checks; network connectivity validated; local teams certified on buyer systems by Day 1 GoLive.</li>
        <li><strong>User Acceptance:</strong> UAT by buyer business representatives across geographies completed by 20 May 2027 with sign-off; post-GoLive support SLAs met for 30 consecutive days.</li>
        <li><strong>Regulatory Compliance:</strong> GDPR, data residency, export-control, and local data protection rules honored; zero compliance violations post-GoLive.</li>
        <li><strong>TSA Execution:</strong> Seller TSA support active until buyer achieves operational readiness (estimated 4-6 months); formal exit criteria met and approved by both parties.</li>
        <li><strong>Budget Compliance:</strong> Project delivers within EUR 2.34M baseline (labour + approved CAPEX); contingency reserve deployed only for risk-driven scope changes.</li>
        <li><strong>Programme Closure by 31 October 2027</strong> with QG5 passed; all artifacts archived; knowledge transfer completed; lessons learned documented.</li>
      </ul>
    </div>

    <!-- RISKS & CONSTRAINTS -->
    <div class="section">
      <h2>Key Risks &amp; Constraints</h2>
      <h3>Top Risks (from Risk Register)</h3>
      <table>
        <tr><th>Risk</th><th>Rating</th><th>Status</th></tr>
        <tr>
          <td>Multi-Site Parallel Cutover (37 sites) — coordination breakdown or single-site failure</td>
          <td>P5×I5 = 25</td>
          <td><span class="pill red">Red</span></td>
        </tr>
        <tr>
          <td>SAP Data Separation Incomplete — data leakage post-GoLive</td>
          <td>P4×I4 = 16</td>
          <td><span class="pill amber">Amber</span></td>
        </tr>
        <tr>
          <td>Buyer Infrastructure Readiness — buyer IT not prepared by GoLive</td>
          <td>P4×I4 = 16</td>
          <td><span class="pill amber">Amber</span></td>
        </tr>
        <tr>
          <td>208-Application Portfolio Transition — licensing delays, vendor lock-in</td>
          <td>P4×I4 = 16</td>
          <td><span class="pill amber">Amber</span></td>
        </tr>
        <tr>
          <td>Regulatory &amp; Data Residency Compliance — breach post-GoLive</td>
          <td>P3×I5 = 15</td>
          <td><span class="pill amber">Amber</span></td>
        </tr>
      </table>
      <h3>Key Constraints</h3>
      <ul>
        <li><strong>Immovable GoLive Date (1 June 2027):</strong> Business and legal deadlines tie to this date; no material postponement allowed beyond 2-week buffer.</li>
        <li><strong>Buyer Confidentiality:</strong> Buyer identity is currently confidential (legal hold). Limited buyer IT team engagement until disclosure cleared; access constraints impact early architecture alignment.</li>
        <li><strong>Multi-Geography Complexity:</strong> 37 sites span EMEA, Americas, APAC. Time zone constraints, local network dependencies, and varied IT maturity levels all require coordinated planning.</li>
        <li><strong>SAP Landscape Complexity:</strong> Packaging uses core SAP modules (MM, SD, CA, CO); separation or instance replication must preserve financial integrity and auditability.</li>
        <li><strong>Budget TBC on CAPEX:</strong> Hardware, network WAN upgrades, cloud capacity, and regulatory audit costs require QG0 board approval before vendor commitments.</li>
        <li><strong>TSA Service Level Ambiguity:</strong> Seller IT and Buyer IT not yet aligned on TSA scope, SLAs, and hand-off criteria; must be locked by end of Phase 1 (18 May 2026).</li>
      </ul>
    </div>

    <!-- ORGANISATION -->
    <div class="section">
      <h2>Organisation &amp; Governance</h2>
      <table>
        <tr><th>Role</th><th>Name / Entity</th><th>Responsibility</th></tr>
        <tr>
          <td>Executive Sponsor</td>
          <td>Bosch Board + Buyer Executive</td>
          <td>Overall carve-out strategy; budget/resource approval; escalation; deal governance</td>
        </tr>
        <tr>
          <td>Programme Manager</td>
          <td>Gill Amandeep Singh (BD/MIL-PSM1)</td>
          <td>End-to-end programme delivery; PMO governance; schedule/budget; SteerCo chair</td>
        </tr>
        <tr>
          <td>PMO / Methodology</td>
          <td>KPMG</td>
          <td>Project orchestration; risk management; quality gates; deliverable standards</td>
        </tr>
        <tr>
          <td>Seller IT Lead</td>
          <td>Bosch Global IT / Packaging CIO</td>
          <td>Seller IT execution; system separation; cutover leadership; TSA scope definition</td>
        </tr>
        <tr>
          <td>Buyer IT Lead</td>
          <td>[Confidential until disclosure]</td>
          <td>Buyer IT design input; infrastructure planning; acceptance testing; post-GoLive readiness</td>
        </tr>
        <tr>
          <td>Steering Committee</td>
          <td>PM + Seller CIO + Buyer IT Lead + Finance + Legal + KPMG PMO</td>
          <td>Weekly reviews; QG gate decisions; risk escalation; change control; legal alignment</td>
        </tr>
        <tr>
          <td>Workstream Leads</td>
          <td>KPMG (SAP, Apps, Infrastructure, Testing, Change Mgmt, Data); Seller IT; Buyer IT</td>
          <td>Domain-specific execution; daily coordination; risk reporting; status dashboards</td>
        </tr>
      </table>
    </div>

    <!-- ASSUMPTIONS -->
    <div class="section">
      <h2>Assumptions</h2>
      <ul>
        <li>Buyer's IT identity and governance structure will be disclosed and engaged by 1 May 2026 to enable architecture co-design.</li>
        <li>Seller commits dedicated Packaging CIO and IT leadership for full 20-month project duration.</li>
        <li>Packaging business operations are stable (no major restructuring, exit of businesses, or site closures) during carve-out period.</li>
        <li>Budget baseline (EUR 2.34M labour + TBC CAPEX) is approved at QG0 Steering Committee by latest 17 April 2026.</li>
        <li>Buyer infrastructure capacity (network, compute, storage) roadmap is committed by end of Phase 2 (27 July 2026).</li>
        <li>SAP ERP scope is frozen at 208 applications and major modules (MM, SD, CA, CO) post-QG1; material additions require change control.</li>
        <li>No competing major IT transformation or carve-out initiatives in Seller organization during Phase 3 & 4 that would compete for SAP/ERP resources.</li>
        <li>TSA contract terms are agreed and signed by end of Phase 1 (18 May 2026); formal service catalogue is published by end of Phase 2.</li>
        <li>Buyer IT staffing plan and organizational structure are finalized by 1 May 2026 to enable joint design sprints.</li>
        <li>Global Packaging data (15+ years) is retained per business records retention policy; no major data purges required during carve-out.</li>
      </ul>
    </div>

    <div class="footer">
      <strong>Project Zebra | Project Charter</strong><br/>
      Issued: 01 April 2026 &nbsp;|&nbsp; Seller: Robert Bosch GmbH &nbsp;|&nbsp; Buyer: [Confidential] &nbsp;|&nbsp;
      PM: Gill Amandeep Singh (BD/MIL-PSM1) &nbsp;|&nbsp; PMO: KPMG &nbsp;|&nbsp;
      Next Review: QG0 (17 April 2026) &nbsp;|&nbsp; Confidential — Internal Use Only
    </div>
  </div>
</body>
</html>"""

OUT.write_text(html, encoding="utf-8")
print(f"Charter written: {OUT}")
_t1 = _dt.now()
print(f"Finished: {_t1.strftime('%Y-%m-%d %H:%M:%S')}  ({(_t1-_t0).total_seconds():.1f}s elapsed)")
