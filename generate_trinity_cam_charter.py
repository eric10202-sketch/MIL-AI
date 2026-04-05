#!/usr/bin/env python3
"""
Generate Trinity-CAM Project Charter (HTML).

Sources:
  Schedule:  active-projects/Trinity-CAM/Trinity-CAM_Project_Schedule.xlsx
  Risks:     active-projects/Trinity-CAM/Trinity-CAM_Risk_Register.xlsx
  Costs:     active-projects/Trinity-CAM/Trinity-CAM_Cost_Plan.xlsx
"""

import base64
from pathlib import Path

HERE   = Path(__file__).parent
LOGO   = HERE / "Bosch.png"
OUTPUT = HERE / "active-projects" / "Trinity-CAM" / "Trinity-CAM_Project_Charter.html"
OUTPUT.parent.mkdir(parents=True, exist_ok=True)

logo_b64 = base64.b64encode(LOGO.read_bytes()).decode() if LOGO.exists() else ""
logo_tag  = f'<img src="data:image/png;base64,{logo_b64}" style="height:36px;" alt="Bosch">' if logo_b64 else ""

HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Trinity-CAM Project Charter</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; color: #1a1a1a; background: #f5f7fa; }}
  .header {{ background: #003B6E; color: #fff; padding: 18px 32px; display: flex; align-items: center; gap: 24px; }}
  .bosch-logo {{ display: flex; align-items: center; }}
  .header-text h1 {{ font-size: 20px; font-weight: 700; letter-spacing: 0.5px; }}
  .header-text h2 {{ font-size: 13px; font-weight: 400; opacity: 0.85; margin-top: 3px; }}
  .page {{ max-width: 1100px; margin: 24px auto; padding: 0 24px 40px; }}
  .section {{ background: #fff; border-radius: 6px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); margin-bottom: 20px; overflow: hidden; }}
  .section-title {{ background: #0066CC; color: #fff; font-size: 12px; font-weight: 700; padding: 9px 18px; letter-spacing: 0.4px; text-transform: uppercase; }}
  .section-body {{ padding: 16px 18px; }}
  .kv-grid {{ display: grid; grid-template-columns: 220px 1fr; row-gap: 7px; }}
  .kv-label {{ font-weight: 600; color: #555; font-size: 12px; }}
  .kv-val {{ color: #1a1a1a; font-size: 12px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
  th {{ background: #003B6E; color: #fff; padding: 7px 10px; text-align: left; font-size: 11px; font-weight: 600; }}
  td {{ padding: 6px 10px; border-bottom: 1px solid #e8ecf2; vertical-align: top; }}
  tr:nth-child(even) {{ background: #EFF4FB; }}
  .badge {{ display: inline-block; padding: 2px 8px; border-radius: 3px; font-size: 10px; font-weight: 700; }}
  .badge-red  {{ background: #e74c3c; color: #fff; }}
  .badge-amber {{ background: #f39c12; color: #fff; }}
  .badge-green {{ background: #27ae60; color: #fff; }}
  .badge-blue  {{ background: #0066CC; color: #fff; }}
  .milestone-bar {{ display: flex; gap: 0; margin: 12px 0; flex-wrap: wrap; }}
  .ms-item {{ flex: 1; min-width: 120px; border-right: 3px solid #fff; background: #0066CC;
              color: #fff; padding: 8px 10px; font-size: 11px; }}
  .ms-item:last-child {{ border-right: none; }}
  .ms-item .ms-date {{ font-size: 10px; opacity: 0.85; margin-top: 2px; }}
  .ms-item.golive {{ background: #003B6E; }}
  .ms-item.qg5 {{ background: #357ab7; }}
  .risk-score {{ display: inline-block; width: 24px; height: 24px; border-radius: 50%;
                 text-align: center; line-height: 24px; font-weight: 700; font-size: 11px; color: #fff; }}
  .rs-high {{ background: #e74c3c; }}
  .rs-med  {{ background: #f39c12; }}
  .rs-low  {{ background: #27ae60; }}
  .cost-total {{ background: #003B6E; color: #fff; font-weight: 700; font-size: 13px; padding: 10px 18px; border-radius: 0 0 5px 5px; }}
  .sig-grid {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-top: 8px; }}
  .sig-box {{ border: 1px solid #dde3ee; border-radius: 4px; padding: 10px; text-align: center; }}
  .sig-role {{ font-size: 11px; color: #666; margin-bottom: 4px; }}
  .sig-name {{ font-weight: 600; font-size: 12px; }}
  .sig-line {{ border-top: 1px solid #ccc; margin: 20px 4px 4px; }}
  .sig-date {{ font-size: 10px; color: #999; }}
  .info-callout {{ background: #EFF4FB; border-left: 4px solid #0066CC; padding: 10px 14px; border-radius: 0 4px 4px 0; font-size: 12px; margin: 8px 0; }}
  p {{ margin: 6px 0; line-height: 1.5; }}
  ul {{ margin: 6px 0 6px 18px; }}
  li {{ margin: 3px 0; font-size: 12px; line-height: 1.5; }}
</style>
</head>
<body>

<!-- HEADER -->
<div class="header">
  <div class="bosch-logo">{logo_tag}</div>
  <div class="header-text">
    <h1>Trinity-CAM — IT Carve-out Project Charter</h1>
    <h2>Johnson Controls International (JCI) &rarr; Robert Bosch GmbH &nbsp;|&nbsp; Air Conditioning Business &nbsp;|&nbsp; Version 1.0 &nbsp;|&nbsp; April 2026</h2>
  </div>
</div>

<div class="page">

<!-- === 1. PROJECT OVERVIEW === -->
<div class="section">
  <div class="section-title">1. Project Overview</div>
  <div class="section-body">
    <div class="kv-grid">
      <span class="kv-label">Project Name</span>      <span class="kv-val">Trinity-CAM</span>
      <span class="kv-label">Seller</span>            <span class="kv-val">Johnson Controls International (JCI)</span>
      <span class="kv-label">Buyer</span>             <span class="kv-val">Robert Bosch GmbH</span>
      <span class="kv-label">Business</span>          <span class="kv-val">JCI Aircondition Business (global HVAC product division)</span>
      <span class="kv-label">Carve-out Model</span>   <span class="kv-val"><strong>Integration</strong> — JCI IT &rarr; Merger Zone (Infosys) &rarr; Bosch IT</span>
      <span class="kv-label">PMO Lead</span>          <span class="kv-val">KPMG</span>
      <span class="kv-label">IT Delivery Partner</span><span class="kv-val">Infosys (Merger Zone build, operation, and migration delivery)</span>
      <span class="kv-label">Worldwide Sites</span>   <span class="kv-val">48 sites across EMEA, APAC, and Americas</span>
      <span class="kv-label">IT Users</span>          <span class="kv-val">12,000 (all currently on JCI infrastructure)</span>
      <span class="kv-label">Applications</span>      <span class="kv-val">1,800+ including an extensive SAP landscape</span>
      <span class="kv-label">Project Start</span>     <span class="kv-val">1 July 2026 (post-TSA expiry)</span>
      <span class="kv-label">GoLive</span>            <span class="kv-val">1 January 2028</span>
      <span class="kv-label">Project Completion</span><span class="kv-val">1 April 2028 (QG5 — 90-day hypercare complete)</span>
      <span class="kv-label">TSA Status</span>        <span class="kv-val">JCI TSA active until 30 June 2026; carve-out commences 1 July 2026</span>
    </div>
  </div>
</div>

<!-- === 2. CONTEXT & BACKGROUND === -->
<div class="section">
  <div class="section-title">2. Context &amp; Background</div>
  <div class="section-body">
    <p>Johnson Controls International (JCI) is divesting its global Aircondition (Aircon) business to Robert Bosch GmbH under a Sale and Purchase Agreement (SPA). The Aircon business designs, manufactures, and services HVAC and air conditioning solutions across 48 worldwide sites employing approximately 12,000 people, all of whom currently operate on JCI IT infrastructure.</p>
    <p>JCI has been providing IT services under a Transitional Service Agreement (TSA) through 30 June 2026. With the TSA expiring, Bosch requires a structured IT carve-out under the <strong>Integration model</strong>: JCI IT infrastructure is separated, migrated through an Infosys-operated Merger Zone, and ultimately integrated into the Bosch IT environment.</p>
    <p>The IT estate in scope comprises <strong>1,800+ applications</strong> — including a large and complex SAP landscape — running across JCI data centres and cloud hosted infrastructure. <strong>Infosys</strong> has been appointed as the primary IT delivery partner responsible for Merger Zone setup, operation, and migration execution across all workstreams.</p>
    <div class="info-callout">
      <strong>Integration Model flow:</strong> JCI IT infrastructure &rarr; Infosys-operated Merger Zone (transient environment) &rarr; Bosch IT environment. The Merger Zone is a temporary landing zone owned by Bosch and operated by Infosys during the migration period.
    </div>
  </div>
</div>

<!-- === 3. OBJECTIVES === -->
<div class="section">
  <div class="section-title">3. Project Objectives</div>
  <div class="section-body">
    <ul>
      <li>Separate all 12,000 JCI Aircon IT users from JCI infrastructure and migrate them to the Merger Zone by QG4 (8 December 2027).</li>
      <li>Migrate all 1,800+ applications — including the full SAP landscape — through the Merger Zone into Bosch-ready state by GoLive (1 January 2028).</li>
      <li>Establish the Infosys-operated Merger Zone as a fully functional IT environment by Phase 2 completion (31 July 2027).</li>
      <li>Execute a secure, GDPR-compliant migration of all personal and business data across 48 sites and 48 jurisdictions.</li>
      <li>Ensure business continuity for JCI Aircon customer-facing systems throughout the migration period with zero tolerated downtime windows.</li>
      <li>Conclude all JCI TSA service dependencies before GoLive; formally exit all TSA service lines during the 90-day hypercare period.</li>
      <li>Deliver the completed integrated IT estate to Bosch IT Operations by QG5 (1 April 2028).</li>
    </ul>
  </div>
</div>

<!-- === 4. SCOPE === -->
<div class="section">
  <div class="section-title">4. Scope</div>
  <div class="section-body">
    <table>
      <tr><th style="width:30%">Workstream</th><th>In Scope</th><th style="width:30%">Out of Scope</th></tr>
      <tr>
        <td><strong>Applications</strong></td>
        <td>All 1,800+ JCI Aircon applications including SAP (FI/CO/SD/MM/HR/PP), non-SAP business apps, customer portals, OT/SCADA-integrated systems</td>
        <td>JCI corporate (non-Aircon) applications; Bosch target-state application development</td>
      </tr>
      <tr>
        <td><strong>Infrastructure</strong></td>
        <td>Merger Zone DC/cloud build (Infosys); WAN/SD-WAN connectivity to 48 sites; security architecture; DR environment</td>
        <td>Bosch target-state infrastructure beyond MZ handover; JCI corporate infrastructure</td>
      </tr>
      <tr>
        <td><strong>End-User Devices</strong></td>
        <td>12,000 user devices (migration to MZ domain); M365 tenant migration; VOIP; device management (Intune)</td>
        <td>Hardware refresh (flagged as CAPEX separate from programme)</td>
      </tr>
      <tr>
        <td><strong>SAP</strong></td>
        <td>Full SAP system copy JCI &rarr; MZ; client separation; interface rewiring; security role redesign; data migration; 2 mock cutovers</td>
        <td>SAP target-state Bosch simplification / harmonisation (post-hypercare Bosch initiative)</td>
      </tr>
      <tr>
        <td><strong>Identity &amp; Access</strong></td>
        <td>Merger Zone AD forest; identity federation JCI &harr; MZ; PAM; MFA for all 12,000 users</td>
        <td>Bosch Active Directory restructuring; Bosch IAM platform consolidation</td>
      </tr>
      <tr>
        <td><strong>Data</strong></td>
        <td>Data classification; GDPR compliance; data migration JCI &rarr; MZ &rarr; Bosch; data segregation and validation</td>
        <td>Data archival and long-term retention policies (Bosch standard; post-QG5)</td>
      </tr>
      <tr>
        <td><strong>Hypercare &amp; Handover</strong></td>
        <td>90-day L3 support; hotfix management; knowledge transfer to Bosch IT; TSA exit; Bosch IT handover</td>
        <td>Ongoing Bosch IT Operations BAU support post-QG5</td>
      </tr>
    </table>
  </div>
</div>

<!-- === 5. MILESTONE SCHEDULE === -->
<div class="section">
  <div class="section-title">5. Key Milestones</div>
  <div class="section-body">
    <div class="milestone-bar">
      <div class="ms-item"><strong>QG0</strong><div class="ms-date">1 Jul 2026</div></div>
      <div class="ms-item"><strong>QG1</strong><div class="ms-date">1 Oct 2026</div></div>
      <div class="ms-item"><strong>QG2&amp;3</strong><div class="ms-date">31 Jul 2027</div></div>
      <div class="ms-item"><strong>QG4</strong><div class="ms-date">8 Dec 2027</div></div>
      <div class="ms-item golive"><strong>GoLive</strong><div class="ms-date">1 Jan 2028</div></div>
      <div class="ms-item qg5"><strong>QG5</strong><div class="ms-date">1 Apr 2028</div></div>
    </div>
    <table>
      <tr><th style="width:10%">Gate</th><th style="width:18%">Date</th><th>Entry Criteria / Deliverables</th></tr>
      <tr><td><strong>QG0</strong></td><td>1 Jul 2026</td><td>Programme kickoff; KPMG PMO mobilised; Infosys SOW signed; all workstream leads appointed</td></tr>
      <tr><td><strong>QG1</strong></td><td>1 Oct 2026</td><td>Concept approved; Merger Zone architecture signed off; application inventory complete; TSA catalogue agreed; wave plan baselined</td></tr>
      <tr><td><strong>QG2&amp;3</strong></td><td>31 Jul 2027</td><td>Merger Zone fully built; SAP interfaces rewired; Wave 1 (~400 apps) validated; infrastructure ready for Phase 3 testing</td></tr>
      <tr><td><strong>QG4</strong></td><td>8 Dec 2027</td><td>All 12,000 users migrated; all 1,800+ apps on MZ; SAP Mock 2 complete; UAT signed off; infrastructure and DR validated; zero blocking defects</td></tr>
      <tr><td><strong>GoLive</strong></td><td>1 Jan 2028</td><td>Final readiness confirmed; Steering Committee sign-off; Merger Zone Day 1 cutover; 24/7 hypercare support active</td></tr>
      <tr><td><strong>QG5</strong></td><td>1 Apr 2028</td><td>90-day hypercare complete; JCI TSA fully exited; Bosch IT handover complete; programme formally closed</td></tr>
    </table>
  </div>
</div>

<!-- === 6. GOVERNANCE === -->
<div class="section">
  <div class="section-title">6. Governance &amp; Organisation</div>
  <div class="section-body">
    <table>
      <tr><th style="width:25%">Role</th><th style="width:30%">Organisation</th><th>Responsibilities</th></tr>
      <tr><td><strong>Programme Sponsor (Customer)</strong></td><td>Robert Bosch GmbH</td><td>Strategic direction; budget authority; deal integration steering; final GoLive approval</td></tr>
      <tr><td><strong>Programme Sponsor (Contractor)</strong></td><td>Johnson Controls International</td><td>JCI system access; TSA compliance; IT staff availability; data quality cooperation</td></tr>
      <tr><td><strong>PMO Lead</strong></td><td>KPMG</td><td>Programme management; schedule governance; risk and issue management; workstream delivery oversight</td></tr>
      <tr><td><strong>IT Delivery Partner</strong></td><td>Infosys</td><td>Merger Zone build and operation; SAP system copy; application migration; 12,000-user migration execution; hypercare L3 support</td></tr>
      <tr><td><strong>Steering Committee</strong></td><td>JCI + Bosch + KPMG</td><td>Monthly governance; programme decisions above KPMG PMO authority; gate approvals; budget releases</td></tr>
      <tr><td><strong>Workstream Leads</strong></td><td>KPMG / Infosys</td><td>Infrastructure, SAP, Applications, Identity &amp; Security, Data Migration, End-User, HR &amp; Legal</td></tr>
    </table>
  </div>
</div>

<!-- === 7. TOP RISKS === -->
<div class="section">
  <div class="section-title">7. Top Programme Risks (from Risk Register)</div>
  <div class="section-body">
    <table>
      <tr>
        <th style="width:6%">ID</th>
        <th style="width:22%">Category</th>
        <th>Risk Description</th>
        <th style="width:12%">P&times;I</th>
        <th style="width:15%">Owner</th>
      </tr>
      <tr>
        <td><strong>R001</strong></td><td>Technology, R&amp;D</td>
        <td>SAP landscape complexity (1,800+ apps, custom Z-programs) causes system copy delay; QG2&amp;3 at risk</td>
        <td><span class="risk-score rs-high">20</span> 70% &times; VH</td>
        <td>KPMG SAP Architect</td>
      </tr>
      <tr>
        <td><strong>R003</strong></td><td>Schedule</td>
        <td>SAP Mock Cutover 2 defects not resolvable before QG4 (Dec 8); GoLive pushed past Jan 1 2028</td>
        <td><span class="risk-score rs-high">15</span> 50% &times; VH</td>
        <td>KPMG PMO Lead</td>
      </tr>
      <tr>
        <td><strong>R002</strong></td><td>Technology, R&amp;D</td>
        <td>Infosys Merger Zone delivery delay; MZ DC not ready by Feb 2027; Phase 3 start impacted</td>
        <td><span class="risk-score rs-med">12</span> 50% &times; H</td>
        <td>Infosys Programme Manager</td>
      </tr>
      <tr>
        <td><strong>R007</strong></td><td>Technology, R&amp;D</td>
        <td>1,800+ app compatibility with MZ; Wave 2/3 remediations exceed capacity; GoLive with incomplete estate</td>
        <td><span class="risk-score rs-med">12</span> 50% &times; H</td>
        <td>Infosys Application Lead</td>
      </tr>
      <tr>
        <td><strong>R006</strong></td><td>Security &amp; Data Protection</td>
        <td>GDPR/data breach during 18-month migration; 12,000 users PII across 48 jurisdictions; EUR 20M fine exposure</td>
        <td><span class="risk-score rs-med">10</span> 30% &times; VH</td>
        <td>Infosys Security Lead</td>
      </tr>
      <tr>
        <td><strong>R005</strong></td><td>Legal &amp; Compliance</td>
        <td>TUPE non-compliance in Germany/France blocks Wave 1/2 user migration (Aug-Sep 2027)</td>
        <td><span class="risk-score rs-med">10</span> 30% &times; VH</td>
        <td>JCI Legal Counsel</td>
      </tr>
      <tr>
        <td><strong>R025</strong></td><td>Technology (Opportunity)</td>
        <td><em>Cloud modernisation: MZ architecture enables cloud-native design; EUR 3-5M/yr Bosch OPEX saving</em></td>
        <td><span class="risk-score rs-low">12</span> 50% &times; H</td>
        <td>Infosys Cloud Lead</td>
      </tr>
    </table>
    <p style="margin-top:8px; font-size:11px; color:#666;">Full register: <em>Trinity-CAM_Risk_Register.xlsx</em> &mdash; 25 risks, 24 threats, 1 opportunity. Risk scores: P&times;I on 1&ndash;5 scale.</p>
  </div>
</div>

<!-- === 8. BUDGET SUMMARY === -->
<div class="section">
  <div class="section-title">8. Budget Summary</div>
  <div class="section-body">
    <table>
      <tr><th>Budget Category</th><th style="width:30%">Estimated Cost (EUR)</th><th>Notes</th></tr>
      <tr><td>KPMG — PMO &amp; Programme Management</td><td>1,378,400</td><td>PMO Lead + Project Manager; full programme duration</td></tr>
      <tr><td>KPMG — Architecture &amp; Technical Advisory</td><td>752,000</td><td>Infrastructure + Data architects</td></tr>
      <tr><td>KPMG — SAP Advisory</td><td>432,000</td><td>SAP Architect; Phase 1-3 heaviest</td></tr>
      <tr><td>Infosys — Programme Management</td><td>828,800</td><td>Infosys Programme Manager; full programme</td></tr>
      <tr><td>Infosys — Infrastructure, Network &amp; Security</td><td>1,642,400</td><td>MZ build; 48-site connectivity; security</td></tr>
      <tr><td>Infosys — Identity &amp; Access Management</td><td>313,600</td><td>AD, federation, PAM, MFA</td></tr>
      <tr><td>Infosys — SAP Build &amp; Configuration</td><td>1,137,600</td><td>SAP copy, separation, rewiring</td></tr>
      <tr><td>Infosys — Application Migration</td><td>518,400</td><td>Waves 1-3, 1,800+ apps</td></tr>
      <tr><td>Infosys — Data Migration</td><td>403,200</td><td>ETL, validation, reconciliation</td></tr>
      <tr><td>Infosys — Service Delivery &amp; Hypercare</td><td>468,000</td><td>Phase 3-5; 12,000-user support</td></tr>
      <tr style="background:#C6D4E8; font-weight:700;"><td><strong>Total Labour (KPMG + Infosys)</strong></td><td><strong>7,873,600</strong></td><td>Subject to contract negotiation and QG1 approval</td></tr>
      <tr><td>Risk Contingency Reserve (15%)</td><td>1,181,040</td><td>For R001, R002, R003; to be approved at QG1</td></tr>
      <tr><td>MZ Infrastructure / Cloud Hosting (CAPEX)</td><td>TBC at QG1</td><td>Infosys contract; cloud-vs-hardware decision at QG1</td></tr>
      <tr><td>Software Licences — M365, ITSM, IAM</td><td>TBC at QG1</td><td>Per-user licence cost 12,000 users</td></tr>
      <tr><td>WAN/SD-WAN Connectivity 48 Sites</td><td>TBC at QG1</td><td>Risk R012; APAC/LATAM ISP premium</td></tr>
    </table>
  </div>
  <div class="cost-total">Total Programme Labour Budget (excl. CAPEX):&nbsp;&nbsp; EUR 7,873,600 &nbsp;|&nbsp; Incl. contingency: EUR 9,054,640</div>
</div>

<!-- === 9. ASSUMPTIONS === -->
<div class="section">
  <div class="section-title">9. Assumptions &amp; Constraints</div>
  <div class="section-body">
    <table>
      <tr><th style="width:50%">Assumptions</th><th>Constraints</th></tr>
      <tr>
        <td>
          <ul>
            <li>JCI TSA formally expires 30 June 2026; carve-out programme start date of 1 July 2026 is firm</li>
            <li>SPA is signed and deal financially closed before QG1 (1 Oct 2026)</li>
            <li>JCI IT cooperation obligations are contractually binding through GoLive</li>
            <li>Infosys Statement of Work is signed before programme start</li>
            <li>Bosch IT integration scope for post-GoLive Bosch standard adoption is defined separately</li>
            <li>All 12,000 users consent to data migration per applicable employment law</li>
            <li>KPMG PMO team is mobilised and available from 1 July 2026</li>
          </ul>
        </td>
        <td>
          <ul>
            <li>GoLive date of 1 January 2028 is a hard business commitment; no flex beyond 1 February 2028</li>
            <li>Infosys is confirmed as the sole MZ delivery partner; no re-tendering</li>
            <li>JCI will not extend TSA services beyond 30 June 2026 except under emergency clause</li>
            <li>Budget approval required at QG1 for CAPEX items; no pre-QG1 hardware procurement</li>
            <li>GDPR and data sovereignty must be resolved before any personal data enters the Merger Zone</li>
            <li>No new SAP development or application features during hypercare (stabilisation only)</li>
          </ul>
        </td>
      </tr>
    </table>
  </div>
</div>

<!-- === 10. APPROVAL === -->
<div class="section">
  <div class="section-title">10. Charter Approval</div>
  <div class="section-body">
    <div class="sig-grid">
      <div class="sig-box">
        <div class="sig-role">Sponsor Customer (Buyer)</div>
        <div class="sig-name">Robert Bosch GmbH</div>
        <div class="sig-line"></div>
        <div class="sig-date">Date: ______________</div>
      </div>
      <div class="sig-box">
        <div class="sig-role">Sponsor Contractor (Seller)</div>
        <div class="sig-name">Johnson Controls International</div>
        <div class="sig-line"></div>
        <div class="sig-date">Date: ______________</div>
      </div>
      <div class="sig-box">
        <div class="sig-role">PMO Lead</div>
        <div class="sig-name">KPMG</div>
        <div class="sig-line"></div>
        <div class="sig-date">Date: ______________</div>
      </div>
      <div class="sig-box">
        <div class="sig-role">IT Delivery Partner</div>
        <div class="sig-name">Infosys</div>
        <div class="sig-line"></div>
        <div class="sig-date">Date: ______________</div>
      </div>
      <div class="sig-box">
        <div class="sig-role">Programme Director</div>
        <div class="sig-name">TBC — to be appointed</div>
        <div class="sig-line"></div>
        <div class="sig-date">Date: ______________</div>
      </div>
      <div class="sig-box">
        <div class="sig-role">Document Status</div>
        <div class="sig-name">DRAFT v1.0</div>
        <div style="margin-top:6px; font-size:11px; color:#666;">For QG0 review and approval<br>April 2026</div>
      </div>
    </div>
  </div>
</div>

</div><!-- /page -->
</body>
</html>"""

OUTPUT.write_text(HTML, encoding="utf-8")
print(f"[Trinity-CAM] Project Charter: {OUTPUT}")
