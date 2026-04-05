#!/usr/bin/env python3
"""
Generate Zebra Monthly Status Report (HTML format)
Report for April 2026 (Project Initiation & QG0 Month)
"""

import base64
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).parent
OUT = HERE / "active-projects" / "Zebra" / "Zebra_Monthly_Status_Report_Apr_2026.html"

logo_b64 = base64.b64encode((HERE / "Bosch.png").read_bytes()).decode()

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Project Zebra – April 2026 Monthly Status Report</title>
  <style>
    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    body {{
      font-family: Calibri, Segoe UI, Arial, sans-serif;
      background: #f4f6f9;
      color: #1a1a1a;
      font-size: 12px;
      line-height: 1.7;
    }}
    .page {{
      max-width: 900px;
      margin: 0 auto;
      padding: 20px;
      background: #fff;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
    }}
    
    /* Header */
    .header {{
      border-bottom: 3px solid #003b6e;
      margin-bottom: 20px;
      padding-bottom: 16px;
    }}
    .logo {{ height: 32px; margin-bottom: 8px; }}
    .header h1 {{
      font-size: 24px;
      font-weight: 700;
      color: #003b6e;
      margin-bottom: 4px;
    }}
    .header p {{ font-size: 12px; color: #666; }}
    
    /* Metadata -->
    .metadata {{
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 12px;
      background: #f0f3f8;
      padding: 12px;
      margin-bottom: 20px;
      border-radius: 6px;
      font-size: 11px;
    }}
    .metadata-item {{ }}
    .metadata-label {{ font-weight: 700; color: #003b6e; }}
    
    /* Section -->
    .section {{
      margin-bottom: 20px;
      page-break-inside: avoid;
    }}
    .section-title {{
      background: #003b6e;
      color: #fff;
      padding: 8px 12px;
      font-weight: 700;
      font-size: 12px;
      margin-bottom: 12px;
      border-radius: 4px;
    }}
    
    /* Status tables -->
    table {{
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 12px;
      font-size: 11px;
    }}
    th {{
      background: #e4edf9;
      padding: 6px;
      text-align: left;
      font-weight: 700;
      color: #003b6e;
      border-bottom: 1px solid #ccc;
    }}
    td {{
      padding: 6px;
      border-bottom: 1px solid #eee;
    }}
    
    /* Pills -->
    .pill {{
      display: inline-block;
      padding: 1px 6px;
      border-radius: 10px;
      font-size: 10px;
      font-weight: 700;
      color: #fff;
    }}
    .pill-green {{ background: #007A33; }}
    .pill-amber {{ background: #E8A000; color: #1a1a1a; }}
    .pill-red {{ background: #CC0000; }}
    
    /* List */
    ul {{ margin-left: 16px; margin-bottom: 8px; }}
    li {{ margin-bottom: 4px; }}
    strong {{ color: #003b6e; }}
    
    /* Footer -->
    .footer {{
      border-top: 1px solid #ddd;
      padding-top: 12px;
      margin-top: 20px;
      font-size: 10px;
      color: #666;
      text-align: center;
    }}
  </style>
</head>
<body>
  <div class="page">
    
    <!-- Header -->
    <div class="header">
      <img src="data:image/png;base64,{logo_b64}" class="logo" alt="Bosch" />
      <h1>PROJECT ZEBRA</h1>
      <p>Packaging Business Carve-Out | Monthly Status Report | April 2026</p>
    </div>
    
    <!-- Metadata -->
    <div class="metadata">
      <div class="metadata-item">
        <div class="metadata-label">Report Period:</div>
        <div>1 - 30 April 2026</div>
      </div>
      <div class="metadata-item">
        <div class="metadata-label">Report Date:</div>
        <div>4 April 2026</div>
      </div>
      <div class="metadata-item">
        <div class="metadata-label">Overall Status:</div>
        <div><span class="pill pill-green">On Track</span></div>
      </div>
      <div class="metadata-item">
        <div class="metadata-label">Days to GoLive:</div>
        <div><strong>427 days</strong></div>
      </div>
      <div class="metadata-item">
        <div class="metadata-label">Budget Status:</div>
        <div><span class="pill pill-green">Approved</span></div>
      </div>
      <div class="metadata-item">
        <div class="metadata-label">Risk Exposure:</div>
        <div><span class="pill pill-red">2 Red Risks</span></div>
      </div>
    </div>
    
    <!-- Executive Summary -->
    <div class="section">
      <div class="section-title">EXECUTIVE SUMMARY</div>
      <p>
        <strong>Project Status: GREEN</strong> — Project Zebra officially commenced on 1 April 2026. The first month focused on 
        project initiation, governance setup, stakeholder alignment, and legal/TSA coordination with the Buyer (currently confidential due to legal hold). 
        All Phase 0 (Initialization) tasks are on track for completion by 17 April 2026, enabling QG0 gate approval. The programme leadership (KPMG PMO + Gill Amandeep Singh, PM) 
        is actively engaged. Budget baseline of EUR 2.34M has been approved. Key dependency: Buyer IT engagement must commence by 1 May 2026 to avoid Phase 1 schedule impact.
      </p>
    </div>
    
    <!-- Phase Progress -->
    <div class="section">
      <div class="section-title">PHASE PROGRESS & MILESTONES</div>
      <table>
        <thead>
          <tr>
            <th>Phase / Milestone</th>
            <th>Planned Date</th>
            <th>Status</th>
            <th>Days to Gate</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td><strong>Phase 0: Initialization (In Progress)</strong></td>
            <td>01-17 Apr 2026</td>
            <td><span class="pill pill-green">90% Done</span></td>
            <td>13 days</td>
          </tr>
          <tr>
            <td>QG0 - Initialization Gate</td>
            <td><strong>17 Apr 2026</strong></td>
            <td><span class="pill pill-green">On Track</span></td>
            <td><strong>13 days</strong></td>
          </tr>
          <tr>
            <td>Phase 1: Concept (Planned)</td>
            <td>18 Apr - 18 May 2026</td>
            <td><span class="pill pill-amber">At Risk</span></td>
            <td>30 days</td>
          </tr>
          <tr>
            <td>QG1 - Concept Gate</td>
            <td>18 May 2026</td>
            <td><span class="pill pill-amber">Risk: Buyer Confidentiality</span></td>
            <td>44 days</td>
          </tr>
          <tr>
            <td>Phase 2: Design & Architecture</td>
            <td>19 May - 27 Jul 2026</td>
            <td><span class="pill pill-green">Planned</span></td>
            <td>76 days</td>
          </tr>
          <tr>
            <td>Phase 3: Build & Test</td>
            <td>28 Jul 2026 - 27 May 2027</td>
            <td><span class="pill pill-green">Planned</span></td>
            <td>115 days</td>
          </tr>
          <tr>
            <td><strong>GoLive Day 1</strong></td>
            <td><strong>1 June 2027</strong></td>
            <td><span class="pill pill-red">CRITICAL</span></td>
            <td><strong>427 days</strong></td>
          </tr>
        </tbody>
      </table>
    </div>
    
    <!-- Accomplishments -->
    <div class="section">
      <div class="section-title">KEY ACCOMPLISHMENTS (WEEK 1)</div>
      <ul>
        <li><strong>Project Kickoff (Apr 1):</strong> Full team mobilization; PMO governance structure activated; steering committee formed.</li>
        <li><strong>Charter & Baseline:</strong> Project charter signed; scope locked (37 sites, 3500+ users, 208 apps); baseline schedule and risk registry established.</li>
        <li><strong>Governance:</strong> Weekly steering committee scheduled; escalation paths defined; KPMG PMO providing daily coordination.</li>
        <li><strong>Legal Alignment:</strong> RoboGmbH Legal and Buyer Legal representatives engaged; TSA scoping commenced; confidentiality agreement on Buyer in place.</li>
        <li><strong>Stakeholder Registry:</strong> 24 key stakeholders mapped (seller IT, finance, legal, procurement); communications plan initiated.</li>
        <li><strong>Risk Management:</strong> Risk register established with 15 identified risks; top 3 risks (P×I ≥ 16) assigned mitigation owners.</li>
        <li><strong>Schedule & Cost Plan:</strong> Detailed project schedule (106 tasks across 5 phases) generated; cost plan (€2.34M labour + CAPEX) developed and queued for budget approval.</li>
      </ul>
    </div>
    
    <!-- Budget Status -->
    <div class="section">
      <div class="section-title">BUDGET & COST STATUS</div>
      <table>
        <thead>
          <tr><th>Category</th><th>Approved Budget</th><th>MTD Spend</th><th>% Consumed</th></tr>
        </thead>
        <tbody>
          <tr>
            <td>KPMG PMO & Governance</td>
            <td>€ 285K</td>
            <td>€ 8K</td>
            <td>2.8%</td>
          </tr>
          <tr>
            <td>KPMG SAP & ERP</td>
            <td>€ 630K</td>
            <td>€ 2K</td>
            <td>0.3%</td>
          </tr>
          <tr>
            <td>KPMG Testing & QA</td>
            <td>€ 495K</td>
            <td>€ 1K</td>
            <td>0.2%</td>
          </tr>
          <tr>
            <td>Other Workstreams</td>
            <td>€ 700K</td>
            <td>€ 2K</td>
            <td>0.3%</td>
          </tr>
          <tr>
            <td><strong>Subtotal (Labour)</strong></td>
            <td><strong>€ 2.11M</strong></td>
            <td><strong>€ 13K</strong></td>
            <td><strong>0.6%</strong></td>
          </tr>
          <tr>
            <td>CAPEX (Risk-Driven)</td>
            <td>€ 226K</td>
            <td>€ 0K</td>
            <td>0%</td>
          </tr>
          <tr>
            <td><strong>TOTAL</strong></td>
            <td><strong>€ 2.34M</strong></td>
            <td><strong>€ 13K</strong></td>
            <td><strong>0.6%</strong></td>
          </tr>
        </tbody>
      </table>
      <p><strong>Status:</strong> Baseline approved. Early April spend is on-plan (primarily Steering Committee, PMO, and kickoff activities). Full resource ramp expected May onwards (Phase 1 design activities).</p>
    </div>
    
    <!-- Risk Summary -->
    <div class="section">
      <div class="section-title">KEY RISKS & MITIGATION</div>
      <ul>
        <li><strong>[RED] Multi-Site Cutover Coordination (P5×I5=25):</strong> 37-site parallel rollout is single highest risk. <em>Mitigation:</em> Detailed site readiness plan due week 8; regional coordinators assigned; rollback playbooks per site. <em>Owner:</em> KPMG PM.</li>
        <li><strong>[RED] SAP Data Separation Incomplete (P4×I4=16):</strong> Packaging/other business unit data leakage post-GoLive. <em>Mitigation:</em> 3 dry-run audits; ABAP custom logic; independent pre-cutover audit. <em>Owner:</em> KPMG SAP CoE.</li>
        <li><strong>[RED] Buyer Infrastructure Gap (P4×I4=16):</strong> Buyer IT readiness unknown due to confidentiality lock. <em>Mitigation:</em> Infrastructure audit scheduled week 6 post-disclosure; capacity validation by week 12. <em>Owner:</em> Buyer IT + KPMG Infra.</li>
        <li><strong>[AMBER] 208-App Portfolio Complexity (P4×I4=16):</strong> ISV licensing delays and vendor lock-in expected. <em>Mitigation:</em> App audit by week 5; early vendor engagement; cost escalation plan. <em>Owner:</em> KPMG Procurement.</li>
        <li><strong>[AMBER] Buyer Confidentiality Delays Phase 1 (P3×I4=12):</strong> Limited buyer IT engagement until legal clears disclosure. <em>Mitigation:</em> Confidentiality structure active; buyer point-of-contact named; encrypted comms established. <em>Owner:</em> RoboGmbH Legal.</li>
      </ul>
    </div>
    
    <!-- Open Actions -->
    <div class="section">
      <div class="section-title">CRITICAL OPEN ACTIONS (P0-P2)</div>
      <table>
        <thead>
          <tr><th>Action</th><th>Due Date</th><th>Owner</th><th>Priority</th></tr>
        </thead>
        <tbody>
          <tr>
            <td>QG0 Steering Committee Approval</td>
            <td>17 Apr 2026</td>
            <td>Gill Amandeep Singh</td>
            <td><span class="pill pill-red">P0</span></td>
          </tr>
          <tr>
            <td>Buyer Legal Disclosure (enable IT engagement)</td>
            <td>01 May 2026</td>
            <td>RoboGmbH Legal</td>
            <td><span class="pill pill-red">P0</span></td>
          </tr>
          <tr>
            <td>Buyer IT Point-of-Contact Assignment</td>
            <td>03 May 2026</td>
            <td>Buyer Exec</td>
            <td><span class="pill pill-red">P0</span></td>
          </tr>
          <tr>
            <td>TSA Service Catalogue (Draft)</td>
            <td>18 May 2026</td>
            <td>KPMG + RoboGmbH IT</td>
            <td><span class="pill pill-amber">P1</span></td>
          </tr>
          <tr>
            <td>Budget Approval (Board Sign-Off)</td>
            <td>10 Apr 2026</td>
            <td>Sponsor Executive</td>
            <td><span class="pill pill-red">P0</span></td>
          </tr>
        </tbody>
      </table>
    </div>
    
    <!-- Outlook -->
    <div class="section">
      <div class="section-title">OUTLOOK & NEXT MONTH</div>
      <ul>
        <li><strong>April 17-30:</strong> Complete Phase 0 tasks; conduct QG0 gate review; secure approval for Phase 1.</li>
        <li><strong>May 1 onwards:</strong> Commence Phase 1 (Concept) with full team including buyer IT (pending disclosure). Requirements &amp; AS-IS documentation begin. SAP roadmap decision gate locked by 10 May.</li>
        <li><strong>Risk Watch:</strong> Buyer confidentiality delays and infrastructure uncertainty remain top gating risks. Mitigation tracks are active.</li>
        <li><strong>Schedule Confidence:</strong> Phase 0 on track (92% complete). Phase 1 at risk if buyer disclosure delayed beyond 1 May; schedule recovery may require parallel design activities. Escalation planned if delay exceeds 1 week.</li>
      </ul>
    </div>
    
    <!-- Footer -->
    <div class="footer">
      <strong>Project Zebra | Monthly Status Report | April 2026</strong><br/>
      Report Date: 04 April 2026 | PM: Gill Amandeep Singh (BD/MIL-PSM1) | PMO: KPMG<br/>
      Confidential — Internal Use Only
    </div>
    
  </div>
</body>
</html>
"""

OUT.write_text(html, encoding="utf-8")
print(f"✓ Monthly status report written: {OUT}")
