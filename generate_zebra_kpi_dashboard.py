#!/usr/bin/env python3
"""
Generate Zebra_Management_KPI_Dashboard.html

Management-level KPI dashboard for Project Zebra.
Includes schedule performance, cost performance, risk summary, and action forecast.
"""

import base64
from datetime import datetime, timedelta
from pathlib import Path

HERE = Path(__file__).parent
OUT = HERE / "active-projects" / "Zebra" / "Zebra_Management_KPI_Dashboard.html"

logo_b64 = base64.b64encode((HERE / "Bosch.png").read_bytes()).decode()
REPORT_DATE = datetime(2026, 4, 4)

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Project Zebra – Management KPI Dashboard</title>
  <style>
    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
      background: #f4f6f9;
      color: #1a1a1a;
      font-size: 13px;
    }}
    .dashboard {{
      max-width: 1200px;
      margin: 0 auto;
      padding: 20px;
    }}
    
    /* Header */
    .header {{
      background: linear-gradient(135deg, #001f45 0%, #003b6e 100%);
      color: #fff;
      padding: 20px;
      border-radius: 8px;
      margin-bottom: 20px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }}
    .logo-container {{
      background: #fff;
      padding: 4px 6px;
      border-radius: 4px;
      width: fit-content;
    }}
    .logo-container img {{ height: 28px; display: block; }}
    .header h1 {{ font-size: 28px; font-weight: 700; margin-bottom: 4px; }}
    .header-right {{ text-align: right; font-size: 12px; }}
    
    /* Card grid (12 columns) */
    .card-grid {{
      display: grid;
      grid-template-columns: repeat(12, 1fr);
      gap: 12px;
      margin-bottom: 20px;
    }}
    .card {{
      background: #fff;
      border-radius: 8px;
      padding: 16px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    }}
    .card-full {{ grid-column: span 12; }}
    .card-half {{ grid-column: span 6; }}
    .card-third {{ grid-column: span 4; }}
    .card-quarter {{ grid-column: span 3; }}
    
    /* Card header */
    .card-header {{
      font-weight: 700;
      color: #003b6e;
      padding-bottom: 8px;
      border-bottom: 1px solid #eee;
      margin-bottom: 12px;
      font-size: 13px;
    }}
    
    /* KPI cards */
    .kpi-value {{
      font-size: 28px;
      font-weight: 700;
      color: #003b6e;
      margin: 8px 0;
    }}
    .kpi-label {{
      font-size: 12px;
      color: #666;
      margin-bottom: 4px;
    }}
    .kpi-status {{
      display: inline-block;
      padding: 2px 8px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: 700;
      color: #fff;
    }}
    .status-green {{ background: #007A33; }}
    .status-amber {{ background: #E8A000; color: #1a1a1a; }}
    .status-red {{ background: #CC0000; }}
    
    /* Progress bar */
    .progress {{
      width: 100%;
      height: 20px;
      background: #eee;
      border-radius: 4px;
      overflow: hidden;
      margin: 8px 0;
    }}
    .progress-fill {{
      height: 100%;
      background: linear-gradient(90deg, #0066cc, #003b6e);
      transition: width 0.3s;
    }}
    
    /* Milestone control */
    .milestone-list {{
      list-style: none;
    }}
    .milestone-list li {{
      padding: 8px 0;
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 12px;
      border-bottom: 1px solid #eee;
    }}
    .milestone-list li:last-child {{ border-bottom: none; }}
    
    /* Action items */
    .action-item {{
      background: #f9fafb;
      padding: 10px;
      margin-bottom: 8px;
      border-radius: 4px;
      border-left: 3px solid #0066cc;
      font-size: 12px;
    }}
    .action-priority-high {{ border-left-color: #CC0000; }}
    .action-priority-med {{ border-left-color: #E8A000; }}
    .action-priority-low {{ border-left-color: #007A33; }}
    
    /* Risk cards in grid */
    .risk-item {{
      background: #f9fafb;
      padding: 8px;
      margin-bottom: 6px;
      border-radius: 4px;
      font-size: 12px;
      border-top: 2px solid;
    }}
    .risk-high {{ border-top-color: #CC0000; }}
    .risk-med {{ border-top-color: #E8A000; }}
    
    /* Table */
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
      margin-top: 8px;
    }}
    th, td {{
      padding: 6px;
      text-align: left;
      border-bottom: 1px solid #eee;
    }}
    th {{
      background: #f0f3f8;
      font-weight: 700;
      color: #003b6e;
    }}
    tbody tr:hover {{ background: #fafbfc; }}
    
    @media print {{
      body {{ background: #fff; }}
      .card {{ box-shadow: none; border: 1px solid #ddd; }}
    }}
  </style>
</head>
<body>
  <div class="dashboard">
    <!-- Header -->
    <div class="header">
      <div>
        <div class="logo-container">
          <img src="data:image/png;base64,{logo_b64}" alt="Bosch" />
        </div>
        <h1>Project Zebra</h1>
        <p style="font-size:12px;opacity:0.9;">Management KPI Dashboard</p>
      </div>
      <div class="header-right">
        <strong>Report Date:</strong> 04 April 2026<br/>
        <strong>Status:</strong> Project Initiation Phase<br/>
        <strong>Overall Health:</strong> <span class="kpi-status status-green">On Track</span>
      </div>
    </div>
    
    <!-- 12-Column KPI Cards -->
    <div class="card-grid">
      
      <!-- Schedule Performance -->
      <div class="card card-third">
        <div class="card-header">SCHEDULE PERFORMANCE</div>
        <div class="kpi-label">Schedule Performance Index (SPI)</div>
        <div class="kpi-value">1.00</div>
        <div class="progress">
          <div class="progress-fill" style="width: 100%;"></div>
        </div>
        <div class="kpi-status status-green">On Schedule</div>
      </div>
      
      <!-- Cost Performance -->
      <div class="card card-third">
        <div class="card-header">COST PERFORMANCE</div>
        <div class="kpi-label">Cost Performance Index (CPI)</div>
        <div class="kpi-value">1.00</div>
        <div class="progress">
          <div class="progress-fill" style="width: 100%;"></div>
        </div>
        <div class="kpi-status status-green">On Budget</div>
      </div>
      
      <!-- Resource Readiness -->
      <div class="card card-third">
        <div class="card-header">RESOURCE READINESS</div>
        <div class="kpi-label">Assigned Resources</div>
        <div class="kpi-value">85%</div>
        <div class="progress">
          <div class="progress-fill" style="width: 85%;"></div>
        </div>
        <div class="kpi-status status-amber">Amber</div>
      </div>
      
      <!-- Quality Gate Status -->
      <div class="card card-quarter">
        <div class="card-header">QG0 READINESS</div>
        <div class="kpi-value">90%</div>
        <div class="kpi-label">Due: 17 Apr 2026</div>
        <div class="kpi-status status-amber">On Track</div>
      </div>
      
      <!-- SAP Confidence -->
      <div class="card card-quarter">
        <div class="card-header">SAP CONFIDENCE</div>
        <div class="kpi-value">3/5</div>
        <div class="kpi-label">Separation Risk: High</div>
        <div class="kpi-status status-amber">Amber</div>
      </div>
      
      <!-- Buyer Readiness -->
      <div class="card card-quarter">
        <div class="card-header">BUYER READINESS</div>
        <div class="kpi-value">2/5</div>
        <div class="kpi-label">Limited Visibility</div>
        <div class="kpi-status status-amber">Amber</div>
      </div>
      
      <!-- Risk Exposure -->
      <div class="card card-quarter">
        <div class="card-header">RISK EXPOSURE</div>
        <div class="kpi-value">15/5</div>
        <div class="kpi-label">High: 2 Red Risks</div>
        <div class="kpi-status status-red">Critical</div>
      </div>
      
      <!-- Milestone Controls -->
      <div class="card card-half">
        <div class="card-header">CRITICAL MILESTONES (NEXT 90 DAYS)</div>
        <ul class="milestone-list">
          <li>
            <span>QG0 - Initialization Gate</span>
            <span><strong>17 Apr</strong> <span class="kpi-status status-green">On Track</span></span>
          </li>
          <li>
            <span>Phase 1 Kickoff</span>
            <span><strong>18 Apr</strong> <span class="kpi-status status-green">Planned</span></span>
          </li>
          <li>
            <span>Buyer IT Engagement</span>
            <span><strong>01 May</strong> <span class="kpi-status status-amber">Risk</span></span>
          </li>
          <li>
            <span>SAP Roadmap Locked</span>
            <span><strong>10 May</strong> <span class="kpi-status status-amber">Amber</span></span>
          </li>
          <li>
            <span>QG1 - Concept Gate</span>
            <span><strong>18 May</strong> <span class="kpi-status status-amber">At Risk</span></span>
          </li>
        </ul>
      </div>
      
      <!-- Top Open Actions -->
      <div class="card card-half">
        <div class="card-header">TOP 5 OPEN ACTIONS (P0-P2)</div>
        <div class="action-item action-priority-high">
          <strong>P0: Buyer Legal Disclosure</strong> - Enable buyer IT engagement; due 01-May-2026
        </div>
        <div class="action-item action-priority-high">
          <strong>P0: TSA Scope Definition</strong> - Align seller/buyer on service levels; due 18-May-2026
        </div>
        <div class="action-item action-priority-med">
          <strong>P1: SAP Separation Strategy</strong> - Finalize instance vs. separation approach; due 10-May-2026
        </div>
        <div class="action-item action-priority-med">
          <strong>P1: App Portfolio Triage</strong> - Confirm retain/retire/transition for 208 apps; due 30-Apr-2026
        </div>
        <div class="action-item action-priority-med">
          <strong>P2: Budget Approval</strong> - Secure QG0 board sign-off for €2.34M; due 10-Apr-2026
        </div>
      </div>
      
      <!-- Risk Summary -->
      <div class="card card-half">
        <div class="card-header">TOP RISKS (P×I ≥ 12)</div>
        <div class="risk-item risk-high">
          <strong>37-Site Cutover Coordination Failure</strong><br/>
          P5×I5=25 | Multi-geography parallel rollout | Mitigation: Readiness checklist, QA validation
        </div>
        <div class="risk-item risk-high">
          <strong>SAP Data Separation Incomplete</strong><br/>
          P4×I4=16 | Data leakage post-GoLive | Mitigation: 3 dry-run audits, ABAP logic
        </div>
        <div class="risk-item risk-high">
          <strong>Buyer Infrastructure Gap</strong><br/>
          P4×I4=16 | Buyer IT capacity unknown | Mitigation: Infrastructure audit week 6
        </div>
        <div class="risk-item risk-med">
          <strong>208-App Portfolio Complexity</strong><br/>
          P4×I4=16 | ISV licensing delays | Mitigation: Early vendor engagement
        </div>
      </div>
      
      <!-- 90-Day Action Forecast -->
      <div class="card card-half">
        <div class="card-header">90-DAY ACTION FORECAST</div>
        <table>
          <thead>
            <tr><th>Timeframe</th><th>Key Activities</th><th>Owner</th></tr>
          </thead>
          <tbody>
            <tr>
              <td><strong>Weeks 1-2</strong><br/>(Apr 4-18)</td>
              <td>Kickoff; charter; PMO setup; legal coordination</td>
              <td>KPMG</td>
            </tr>
            <tr>
              <td><strong>Weeks 3-6</strong><br/>(Apr 19-May 16)</td>
              <td>AS-IS docs; requirements; SAP maps; vendor outreach</td>
              <td>KPMG + Ops</td>
            </tr>
            <tr>
              <td><strong>Weeks 7-13</strong><br/>(May 17-Jun 27)</td>
              <td>Design; architecture; TSA finalization; QG2/3 prep</td>
              <td>KPMG + IT</td>
            </tr>
            <tr>
              <td><strong>Critical Gate</strong></td>
              <td><strong>QG1 (18 May)</strong> – reqs approved; <strong>QG2/3 (27 Jul)</strong> – designs signed</td>
              <td>SteerCo</td>
            </tr>
          </tbody>
        </table>
      </div>
      
      <!-- Buyer vs Seller Model Differences -->
      <div class="card card-full">
        <div class="card-header">STAND ALONE MODEL vs TSA: Key Execution Implications</div>
        <table>
          <thead>
            <tr>
              <th>Aspect</th>
              <th>Stand Alone Model (Primary)</th>
              <th>TSA (Post-GoLive, 4-6 Months)</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><strong>Infrastructure</strong></td>
              <td>100% buyer-owned; no residual Seller systems post-GoLive</td>
              <td>Fallback access to legacy systems for knowledge transfer only</td>
            </tr>
            <tr>
              <td><strong>Buyer IT Readiness</strong></td>
              <td>Critical blocker; must be QG4-ready (pre-GoLive)</td>
              <td>TSA covers transition gaps; phased seller wind-down</td>
            </tr>
            <tr>
              <td><strong>Data Residency</strong></td>
              <td>Buyer assumes compliance; residency rules buyer-determined</td>
              <td>Seller responsible for GDPR/residency until TSA exit</td>
            </tr>
            <tr>
              <td><strong>Risk</strong></td>
              <td>Buyer carries infrastructure/integration risk from Day 1</td>
              <td>Seller provides hypercare support; shared risk in transition</td>
            </tr>
            <tr>
              <td><strong>Cost</strong></td>
              <td>Buyer IT infrastructure cost (not in carve-out budget)</td>
              <td>TSA service costs (hourly, estimated €200K over 6 months)</td>
            </tr>
          </tbody>
        </table>
      </div>
      
    </div>
    
    <!-- Footer -->
    <div style="text-align: center; font-size: 11px; color: #666; margin-top: 20px; padding: 12px; background: #f0f3f8; border-radius: 6px;">
      <strong>Project Zebra | Management KPI Dashboard</strong><br/>
      Report Date: 04 April 2026 | Next Review: 18 April 2026 (QG0) | Confidential — Internal Use Only
    </div>
    
  </div>
</body>
</html>
"""

OUT.write_text(html, encoding="utf-8")
print(f"✓ Management KPI dashboard written: {OUT}")
