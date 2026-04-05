#!/usr/bin/env python3
"""
Generate Trinity-CAM Monthly Status Report (HTML).
Follows monthly-status-report-generation SKILL.md.
Report month: July 2026 (Programme Initiation — QG0 month).
All content derived from Trinity-CAM schedule/cost/risk data.
"""

import base64
from datetime import date
from pathlib import Path

HERE   = Path(__file__).parent
LOGO   = HERE / "Bosch.png"
TODAY  = date.today()
MONTH  = TODAY.strftime("%b_%Y")
OUTPUT = HERE / "active-projects" / "Trinity-CAM" / f"Trinity-CAM_Monthly_Status_Report_{MONTH}.html"
OUTPUT.parent.mkdir(parents=True, exist_ok=True)

logo_b64 = base64.b64encode(LOGO.read_bytes()).decode() if LOGO.exists() else ""
logo_tag  = f'<img src="data:image/png;base64,{logo_b64}" alt="Bosch" style="height:36px;display:block;" />' if logo_b64 else ""

QG1    = date(2026, 10, 1)
QG23   = date(2027, 7, 31)
QG4    = date(2027, 12, 8)
GOLIVE = date(2028, 1, 1)
QG5    = date(2028, 4, 1)

def ddays(d):
    return (d - TODAY).days

HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Trinity-CAM Monthly Status Report — {TODAY.strftime('%B %Y')}</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:Calibri,'Segoe UI',Arial,sans-serif;background:#f4f6f9;color:#1a1a1a;font-size:12px;line-height:1.6;}}
.page{{max-width:900px;margin:0 auto;padding:20px;background:#fff;box-shadow:0 2px 8px rgba(0,0,0,.1);}}

/* --- HEADER --- */
.hdr{{background:linear-gradient(135deg,#003b6e 0%,#005199 100%);color:#fff;padding:16px 20px;display:flex;align-items:center;gap:16px;border-radius:4px 4px 0 0;}}
.bosch-logo{{display:flex;align-items:center;background:#fff;padding:4px 8px;border-radius:4px;flex-shrink:0;}}
.hdr-info h1{{font-size:17px;font-weight:700;}}
.hdr-info h2{{font-size:11px;font-weight:400;opacity:.85;margin-top:3px;}}
.hdr-right{{margin-left:auto;text-align:right;font-size:10px;opacity:.85;}}
.hdr-right strong{{font-size:13px;display:block;}}

/* --- RAG SUMMARY BAR --- */
.rag-bar{{display:flex;background:#eef2f7;border:1px solid #dde3ee;border-radius:0 0 0 0;overflow:hidden;}}
.rag-item{{flex:1;padding:7px 10px;border-right:1px solid #dde3ee;text-align:center;}}
.rag-item:last-child{{border-right:none;}}
.rag-label{{font-size:9px;color:#666;text-transform:uppercase;letter-spacing:.4px;}}
.rag-val{{font-size:13px;font-weight:700;margin-top:1px;}}
.rv-g{{color:#007A33;}} .rv-a{{color:#E8A000;}} .rv-r{{color:#CC0000;}} .rv-b{{color:#005199;}}

/* --- SECTION --- */
.section{{margin-top:14px;}}
.sec-title{{background:#005199;color:#fff;font-size:11px;font-weight:700;padding:5px 10px;letter-spacing:.4px;text-transform:uppercase;border-radius:3px 3px 0 0;}}
.sec-body{{background:#fff;border:1px solid #dde3ee;border-top:none;padding:10px 12px;border-radius:0 0 3px 3px;}}

/* --- KEY FACTS GRID --- */
.kf-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;}}
.kf-cell{{background:#EFF4FB;border-radius:3px;padding:6px 10px;}}
.kf-label{{font-size:9px;color:#555;text-transform:uppercase;letter-spacing:.4px;}}
.kf-val{{font-weight:700;font-size:12px;color:#003b6e;}}

/* --- STATUS TABLE --- */
table.st{{width:100%;border-collapse:collapse;font-size:11px;}}
.st th{{background:#003b6e;color:#fff;padding:5px 8px;text-align:left;font-size:10px;}}
.st td{{padding:5px 8px;border-bottom:1px solid #e8ecf2;vertical-align:top;}}
.st tr:nth-child(even){{background:#EFF4FB;}}

/* --- RAG PILL --- */
.rpill{{display:inline-block;padding:1px 7px;border-radius:10px;font-size:9px;font-weight:700;}}
.rp-g{{background:#e9f7ef;color:#007A33;border:1px solid #b7e4c7;}}
.rp-a{{background:#fef9e7;color:#b7770d;border:1px solid #f8e58e;}}
.rp-r{{background:#fde;color:#CC0000;border:1px solid #f5b7b1;}}
.rp-b{{background:#EFF4FB;color:#005199;border:1px solid #aed6f1;}}

/* --- RISK TABLE --- */
.score{{display:inline-block;width:20px;height:20px;border-radius:50%;font-size:9px;font-weight:700;text-align:center;line-height:20px;color:#fff;}}
.sr{{background:#e74c3c;}} .sa{{background:#f39c12;}} .sg{{background:#27ae60;}}

/* --- MILESTONES --- */
.ms-row{{display:flex;gap:8px;padding:4px 0;border-bottom:1px solid #eee;font-size:11px;align-items:center;}}
.ms-row:last-child{{border-bottom:none;}}
.ms-gate{{min-width:55px;font-weight:700;color:#003b6e;}}
.ms-date{{min-width:90px;color:#555;}}
.ms-cd{{min-width:70px;font-size:10px;color:#0066CC;font-weight:600;}}
.ms-desc{{flex:1;}}

/* --- PROGRESS BAR --- */
.prog-wrap{{height:16px;background:#e8ecf2;border-radius:8px;overflow:hidden;margin:3px 0;}}
.prog-bar{{height:100%;border-radius:8px;background:#0066CC;display:flex;align-items:center;padding-left:6px;color:#fff;font-size:8px;font-weight:700;}}

/* --- INFO BOX --- */
.info-box{{background:#EFF4FB;border-left:4px solid #0066CC;padding:7px 10px;margin:6px 0;border-radius:0 3px 3px 0;font-size:11px;}}

/* --- FOOTER --- */
.footer{{background:#003b6e;color:#fff;text-align:center;font-size:9px;padding:7px;border-radius:0 0 4px 4px;opacity:.85;margin-top:14px;}}

p{{margin:4px 0;}}
ul{{margin:4px 0 4px 16px;}}
li{{margin:2px 0;font-size:11px;}}
</style>
</head>
<body>
<div class="page">

<!-- HEADER -->
<div class="hdr">
  <div class="bosch-logo">{logo_tag}</div>
  <div class="hdr-info">
    <h1>Trinity-CAM — Monthly Status Report <span style="font-size:13px;font-weight:400;opacity:.8;">[Pre-Initiation]</span></h1>
    <h2>JCI Aircondition &rarr; Merger Zone (Infosys) &rarr; Robert Bosch GmbH &nbsp;|&nbsp; Integration Model</h2>
  </div>
  <div class="hdr-right">
    <strong>{TODAY.strftime('%B %Y')}</strong>
    Report Date: {TODAY.strftime('%d %b %Y')}
  </div>
</div>

<!-- PRE-INITIATION NOTICE -->
<div style="background:#fff8e1;border-left:5px solid #E8A000;padding:8px 14px;font-size:11px;margin-bottom:2px;">
  <strong>Pre-Initiation Status Report — {TODAY.strftime('%B %Y')}</strong> &nbsp;|
  Formal programme baseline: <strong>1 July 2026 (QG0)</strong>. This report covers pre-mobilisation activities undertaken prior to contractual programme start.
  SPI / CPI metrics will be set to 1.00 (QG0 baseline) from July 2026. Schedule and cost baselines become active at QG0.
</div>

<!-- RAG BAR -->
<div class="rag-bar">
  <div class="rag-item"><div class="rag-label">Overall RAG</div><div class="rag-val rv-a">AMBER</div></div>
  <div class="rag-item"><div class="rag-label">Schedule SPI</div><div class="rag-val rv-a">1.00</div></div>
  <div class="rag-item"><div class="rag-label">Cost CPI</div><div class="rag-val rv-a">1.00</div></div>
  <div class="rag-item"><div class="rag-label">Readiness</div><div class="rag-val rv-b">12%</div></div>
  <div class="rag-item"><div class="rag-label">Top Risk</div><div class="rag-val rv-r">20 (R001)</div></div>
  <div class="rag-item"><div class="rag-label">Days to GoLive</div><div class="rag-val rv-b">{ddays(GOLIVE)}</div></div>
</div>

<!-- SECTION 1: EXECUTIVE SUMMARY -->
<div class="section">
  <div class="sec-title">1. Executive Summary</div>
  <div class="sec-body">
    <p><strong>Report context:</strong> This is the pre-initiation status report prepared in <strong>{TODAY.strftime('%B %Y')}</strong>, ahead of the formal programme start on <strong>1 July 2026 (QG0)</strong>. Pre-mobilisation activities are underway: KPMG PMO is being staffed, Infosys Statement of Work is in final negotiation, and Bosch / JCI Steering Committee membership has been agreed. SPI, CPI, and readiness metrics will be formally baselined at QG0.</p>
    <p>Project Trinity-CAM formally commences on <strong>1 July 2026</strong> following expiry of the JCI Transitional Service Agreement (TSA). The programme is in <strong>Phase 0 — Initiation &amp; Governance</strong>. KPMG PMO is being mobilised. Infosys Statement of Work is in final sign-off. The Steering Committee has been constituted with Bosch (Sponsor Customer) and JCI (Sponsor Contractor) representation.</p>
    <p>Overall programme RAG status is <strong>AMBER</strong>. The programme is on plan at QG0 baseline with SPI/CPI of 1.00. The AMBER rating reflects the open risk on SAP landscape complexity (R001, score 20) which requires an early response plan before QG1, and the GDPR clearance process which has commenced but not yet concluded across all 48 jurisdictions.</p>
    <p>The next major gate is <strong>QG1 on 1 October 2026</strong> ({ddays(QG1)} days from today). Key deliverables for QG1 are: application inventory complete (1,800+), Merger Zone architecture approved, wave plan baselined, and CAPEX budget approved by Bosch Finance.</p>
    <div class="info-box"><strong>No schedule slip. No cost overrun. Key risk R001 (SAP complexity) response plan to be tabled at August 2026 SteerCo.</strong></div>
  </div>
</div>

<!-- SECTION 2: KEY PROJECT FACTS -->
<div class="section">
  <div class="sec-title">2. Key Project Facts</div>
  <div class="sec-body">
    <div class="kf-grid">
      <div class="kf-cell"><div class="kf-label">Project</div><div class="kf-val">Trinity-CAM</div></div>
      <div class="kf-cell"><div class="kf-label">Seller</div><div class="kf-val">Johnson Controls Int. (JCI)</div></div>
      <div class="kf-cell"><div class="kf-label">Buyer</div><div class="kf-val">Robert Bosch GmbH</div></div>
      <div class="kf-cell"><div class="kf-label">Business</div><div class="kf-val">Aircondition Division</div></div>
      <div class="kf-cell"><div class="kf-label">Model</div><div class="kf-val">Integration (via Merger Zone)</div></div>
      <div class="kf-cell"><div class="kf-label">PMO</div><div class="kf-val">KPMG</div></div>
      <div class="kf-cell"><div class="kf-label">Delivery Partner</div><div class="kf-val">Infosys (MZ &amp; Migration)</div></div>
      <div class="kf-cell"><div class="kf-label">Sites / Users</div><div class="kf-val">48 Sites | 12,000 Users</div></div>
      <div class="kf-cell"><div class="kf-label">Applications</div><div class="kf-val">1,800+ (incl. SAP)</div></div>
      <div class="kf-cell"><div class="kf-label">Programme Start</div><div class="kf-val">1 July 2026</div></div>
      <div class="kf-cell"><div class="kf-label">GoLive</div><div class="kf-val">1 January 2028</div></div>
      <div class="kf-cell"><div class="kf-label">Completion (QG5)</div><div class="kf-val">1 April 2028</div></div>
    </div>
  </div>
</div>

<!-- SECTION 3: PHASE STATUS -->
<div class="section">
  <div class="sec-title">3. Phase &amp; Workstream Status</div>
  <div class="sec-body">
    <table class="st">
      <tr><th style="width:20%">Workstream</th><th style="width:8%">RAG</th><th style="width:15%">Progress</th><th>Status Summary</th></tr>
      <tr>
        <td><strong>Programme Control</strong></td>
        <td><span class="rpill rp-g">GREEN</span></td>
        <td><div class="prog-wrap"><div class="prog-bar" style="width:15%;">15%</div></div></td>
        <td>KPMG PMO mobilised; SteerCo constituted; programme plan v1 baselined. Fortnightly health reviews scheduled. Dashboard tooling provisioned.</td>
      </tr>
      <tr>
        <td><strong>MZ Infrastructure</strong></td>
        <td><span class="rpill rp-a">AMBER</span></td>
        <td><div class="prog-wrap"><div class="prog-bar" style="width:5%;">5%</div></div></td>
        <td>MZ architecture scoping in progress. Infosys preparing DC-vs-cloud recommendation. SD-WAN topology for 48 sites under review. Architecture sign-off target: QG1.</td>
      </tr>
      <tr>
        <td><strong>SAP Migration</strong></td>
        <td><span class="rpill rp-a">AMBER</span></td>
        <td><div class="prog-wrap"><div class="prog-bar" style="width:3%;">3%</div></div></td>
        <td>SAP landscape discovery commenced. Initial count: 47+ SAP systems, 120+ interfaces, significant Z-program estate. R001 SAP complexity response plan drafted; awaiting Infosys input for Aug SteerCo.</td>
      </tr>
      <tr>
        <td><strong>App Migration (1,800+)</strong></td>
        <td><span class="rpill rp-b">PLANNING</span></td>
        <td><div class="prog-wrap"><div class="prog-bar" style="width:2%;">2%</div></div></td>
        <td>Application inventory in progress. Infosys catalogue team embedded at JCI UK site. Target: 100% inventory complete 15 Aug 2026. Wave strategy to be defined at QG1.</td>
      </tr>
      <tr>
        <td><strong>End-User Workplace</strong></td>
        <td><span class="rpill rp-b">PLANNING</span></td>
        <td><div class="prog-wrap"><div class="prog-bar" style="width:2%;">2%</div></div></td>
        <td>Site cluster plan drafted. 48 sites grouped into 6 migration clusters. M365 MZ tenant provisioning scheduled for Q4 2026 (Phase 2). No blockers identified.</td>
      </tr>
      <tr>
        <td><strong>Identity &amp; Access</strong></td>
        <td><span class="rpill rp-g">GREEN</span></td>
        <td><div class="prog-wrap"><div class="prog-bar" style="width:5%;">5%</div></div></td>
        <td>IAM architecture defined. MZ AD forest design approved in principle. PAM tooling selection in progress. MFA rollout plan drafted for Phase 2 execution.</td>
      </tr>
      <tr>
        <td><strong>Security &amp; GDPR</strong></td>
        <td><span class="rpill rp-a">AMBER</span></td>
        <td><div class="prog-wrap"><div class="prog-bar" style="width:8%;">8%</div></div></td>
        <td>DPIA framework initiated. External GDPR counsel engaged in Germany, France, China, Japan, USA. Legal basis scoping underway. All transfer mechanisms to be assessed before data lands in MZ (Phase 2).</td>
      </tr>
      <tr>
        <td><strong>Data Migration</strong></td>
        <td><span class="rpill rp-b">PLANNING</span></td>
        <td><div class="prog-wrap"><div class="prog-bar" style="width:2%;">2%</div></div></td>
        <td>Data classification framework being defined. Infosys data migration lead appointed. ETL tooling selection to be made at QG1. 48-jurisdiction data flow mapping to start Aug 2026.</td>
      </tr>
      <tr>
        <td><strong>TSA Exit &amp; HR/Legal</strong></td>
        <td><span class="rpill rp-a">AMBER</span></td>
        <td><div class="prog-wrap"><div class="prog-bar" style="width:5%;">5%</div></div></td>
        <td>TSA service dependency audit completed: 214 service lines catalogued. TUPE consultation initiated in Germany (6 sites, works council notification filed). France EWC notification scheduled Aug 2026.</td>
      </tr>
    </table>
  </div>
</div>

<!-- SECTION 4: MILESTONE STATUS -->
<div class="section">
  <div class="sec-title">4. Milestone &amp; Gate Status</div>
  <div class="sec-body">
    <div class="ms-row"><div class="ms-gate">QG0</div><div class="ms-date">1 Jul 2026</div><div class="ms-cd" style="color:#007A33;">ACTIVE</div><div class="ms-desc">Programme kickoff complete. All workstream leads appointed. Steering Committee constituted.</div><span class="rpill rp-g">ON TRACK</span></div>
    <div class="ms-row"><div class="ms-gate">QG1</div><div class="ms-date">1 Oct 2026</div><div class="ms-cd">{ddays(QG1)}d</div><div class="ms-desc">App inventory; MZ architecture; CAPEX budget; wave plan. In preparation — on target.</div><span class="rpill rp-a">IN PREP</span></div>
    <div class="ms-row"><div class="ms-gate">QG2&amp;3</div><div class="ms-date">31 Jul 2027</div><div class="ms-cd">{ddays(QG23)}d</div><div class="ms-desc">MZ DC live; SAP copy complete; Wave 1 apps validated.</div><span class="rpill rp-b">PLANNED</span></div>
    <div class="ms-row"><div class="ms-gate">QG4</div><div class="ms-date">8 Dec 2027</div><div class="ms-cd">{ddays(QG4)}d</div><div class="ms-desc">All 12,000 users + 1,800+ apps migrated; SAP Mock 2 passed.</div><span class="rpill rp-b">PLANNED</span></div>
    <div class="ms-row"><div class="ms-gate">GoLive</div><div class="ms-date">1 Jan 2028</div><div class="ms-cd">{ddays(GOLIVE)}d</div><div class="ms-desc">MZ Day 1 cutover; hypercare active.</div><span class="rpill rp-b">PLANNED</span></div>
    <div class="ms-row"><div class="ms-gate">QG5</div><div class="ms-date">1 Apr 2028</div><div class="ms-cd">{ddays(QG5)}d</div><div class="ms-desc">90-day hypercare complete; TSA exit confirmed; Bosch handover done.</div><span class="rpill rp-b">PLANNED</span></div>
  </div>
</div>

<!-- SECTION 5: RISK HIGHLIGHTS -->
<div class="section">
  <div class="sec-title">5. Risk Highlights (Top 5)</div>
  <div class="sec-body" style="padding:8px 10px;">
    <table class="st">
      <tr><th>ID</th><th>Description</th><th>P</th><th>I</th><th>Score</th><th>Owner</th><th>Action</th></tr>
      <tr>
        <td><b>R001</b></td>
        <td>SAP landscape complexity (47+ systems, 120+ interfaces) delays QG2&amp;3</td>
        <td>70%</td><td>VH</td>
        <td><span class="score sr">20</span></td>
        <td>KPMG SAP</td>
        <td>Response plan at Aug SteerCo; Infosys SAP offshore lead appointed</td>
      </tr>
      <tr>
        <td><b>R003</b></td>
        <td>SAP Mock Cutover 2 defects not cleared before QG4; GoLive slips</td>
        <td>50%</td><td>VH</td>
        <td><span class="score sr">15</span></td>
        <td>KPMG PMO</td>
        <td>Mock 1 scheduled Jul 2027 to validate remediation timeline</td>
      </tr>
      <tr>
        <td><b>R002</b></td>
        <td>Infosys MZ DC delivery delay (procurement lead times); Phase 3 start risk</td>
        <td>50%</td><td>H</td>
        <td><span class="score sa">12</span></td>
        <td>Infosys PM</td>
        <td>DC procurement to start at QG1 approval; cloud fallback option being assessed</td>
      </tr>
      <tr>
        <td><b>R006</b></td>
        <td>GDPR breach during 18-month migration; 12,000 users PII, 48 jurisdictions</td>
        <td>30%</td><td>VH</td>
        <td><span class="score sa">10</span></td>
        <td>Infosys Sec</td>
        <td>DPIA in progress; external counsel engaged; security architecture reviewed before data enters MZ</td>
      </tr>
      <tr>
        <td><b>R005</b></td>
        <td>TUPE non-compliance in Germany/France blocks Wave 1/2 user migrations</td>
        <td>30%</td><td>VH</td>
        <td><span class="score sa">10</span></td>
        <td>JCI Legal</td>
        <td>Works council Germany notified Jul 2026; EWC notification France Aug 2026; HR lead assigned</td>
      </tr>
    </table>
  </div>
</div>

<!-- SECTION 6: BUDGET STATUS -->
<div class="section">
  <div class="sec-title">6. Budget Status</div>
  <div class="sec-body">
    <div class="kf-grid">
      <div class="kf-cell"><div class="kf-label">Labour Budget</div><div class="kf-val">EUR 7,873,600</div></div>
      <div class="kf-cell"><div class="kf-label">Incl. Contingency</div><div class="kf-val">EUR 9,054,640</div></div>
      <div class="kf-cell"><div class="kf-label">CAPEX</div><div class="kf-val">TBC at QG1</div></div>
      <div class="kf-cell"><div class="kf-label">Actual Cost to Date</div><div class="kf-val">EUR 629,888 (Ph1)</div></div>
      <div class="kf-cell"><div class="kf-label">Cost Variance</div><div class="kf-val">EUR 0 (on budget)</div></div>
      <div class="kf-cell"><div class="kf-label">CPI</div><div class="kf-val">1.00 (baseline)</div></div>
    </div>
    <p style="margin-top:8px;font-size:11px;color:#555;">Phase 1 budget (EUR 629,888, 8%) is on track. CAPEX items (MZ infrastructure, M365 licences, SD-WAN) are pending Bosch Finance approval at QG1. Risk contingency reserve of EUR 1,181,040 is held centrally by KPMG PMO.</p>
  </div>
</div>

<!-- SECTION 7: NEXT STEPS -->
<div class="section">
  <div class="sec-title">7. Key Actions for Next 30 Days</div>
  <div class="sec-body">
    <ul>
      <li><strong>Application Inventory:</strong> Infosys to complete 100% catalogue of all 1,800+ JCI Aircon applications by 15 August 2026, including SAP system map and interface register.</li>
      <li><strong>MZ Architecture:</strong> Infosys to submit MZ architecture document (DC/cloud topology, SD-WAN design, security baseline) to KPMG for review by 20 August 2026.</li>
      <li><strong>R001 SAP Risk Response:</strong> KPMG SAP Architect and Infosys SAP Lead to present joint SAP migration strategy and risk mitigation plan to the August 2026 Steering Committee.</li>
      <li><strong>CAPEX Business Case:</strong> Infosys to submit Merger Zone infrastructure CAPEX model to Bosch Finance for QG1 approval (target: 1 October 2026).</li>
      <li><strong>GDPR:</strong> Complete legal basis mapping for all 48 jurisdictions. Resolve CAC approval process for China data transfers. Target: Aug 2026.</li>
      <li><strong>TUPE — France:</strong> File EWC (European Works Council) notification for changes affecting France sites. Target: August 2026.</li>
      <li><strong>Wave Plan:</strong> KPMG and Infosys to finalise migration wave plan (app categorisation, site sequencing, SAP track) for QG1 sign-off.</li>
    </ul>
  </div>
</div>

<!-- FOOTER -->
<div class="footer">
  Trinity-CAM Monthly Status Report &nbsp;|&nbsp; {TODAY.strftime('%B %Y')} &nbsp;|&nbsp;
  Sources: Trinity-CAM_Project_Schedule.xlsx, Trinity-CAM_Risk_Register.xlsx, Trinity-CAM_Cost_Plan.xlsx &nbsp;|&nbsp;
  CONFIDENTIAL &mdash; KPMG / Bosch Internal &nbsp;|&nbsp; PMO: KPMG
</div>

</div><!-- /page -->
</body>
</html>"""

OUTPUT.write_text(HTML, encoding="utf-8")
print(f"[Trinity-CAM] Monthly Status Report: {OUTPUT}")
