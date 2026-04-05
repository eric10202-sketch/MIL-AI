#!/usr/bin/env python3
"""
Generate Zebra_Executive_Dashboard.html

Executive-level dashboard for Project Zebra.
Combines schedule, cost, and risk data into a strategic visual summary.
"""

import base64
from datetime import datetime, timedelta
from pathlib import Path

HERE = Path(__file__).parent
OUT = HERE / "active-projects" / "Zebra" / "Zebra_Executive_Dashboard.html"

# Load Bosch logo
logo_b64 = base64.b64encode((HERE / "Bosch.png").read_bytes()).decode()

# Key dates
START_DATE = datetime(2026, 4, 1)
GOLIVE_DATE = datetime(2027, 6, 1)
COMPLETE_DATE = datetime(2027, 10, 31)
REPORT_DATE = datetime(2026, 4, 4)

def days_until(target_date):
    """Calculate days from report date to target."""
    delta = target_date - REPORT_DATE
    return delta.days

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Project Zebra – Executive Dashboard</title>
  <style>
    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      background: #f4f6f9;
      color: #1a1a1a;
      font-size: 13px;
      line-height: 1.6;
    }}
    .page {{
      max-width: 1024px;
      margin: 0 auto;
      padding: 20px;
      background: #fff;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      page-break-after: always;
    }}
    .page:last-child {{ page-break-after: auto; }}
    
    /* Header */
    .header {{
      background: linear-gradient(135deg, #001f45 0%, #003b6e 50%, #0066cc 100%);
      color: #fff;
      padding: 24px;
      border-radius: 8px;
      margin-bottom: 20px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }}
    .header-left {{
      flex: 1;
    }}
    .logo-container {{
      background: #fff;
      padding: 4px 8px;
      border-radius: 4px;
      width: fit-content;
      margin-bottom: 12px;
    }}
    .logo-container img {{
      height: 36px;
      display: block;
    }}
    .header h1 {{
      font-size: 32px;
      font-weight: 700;
      margin-bottom: 4px;
    }}
    .header p {{
      font-size: 14px;
      opacity: 0.9;
    }}
    .header-right {{
      text-align: right;
      font-size: 12px;
    }}
    .countdown {{
      background: #E8A000;
      color: #1a1a1a;
      padding: 4px 12px;
      border-radius: 4px;
      font-weight: 700;
      margin-top: 8px;
      display: inline-block;
    }}
    
    /* Countdown strip */
    .countdown-strip {{
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 12px;
      margin-bottom: 20px;
      background: #f0f3f8;
      padding: 16px;
      border-radius: 8px;
    }}
    .countdown-box {{
      background: #fff;
      border-left: 3px solid #CC0000;
      padding: 12px;
      border-radius: 4px;
      font-size: 12px;
    }}
    .countdown-box .label {{
      font-weight: 700;
      color: #003b6e;
      margin-bottom: 4px;
    }}
    .countdown-box .value {{
      font-size: 20px;
      font-weight: 700;
      color: #CC0000;
    }}
    .countdown-box .unit {{
      font-size: 11px;
      color: #666;
    }}
    
    /* Sections */
    .section {{
      margin-bottom: 20px;
      background: #fff;
      border-radius: 8px;
      overflow: hidden;
    }}
    .section-header {{
      background: #003b6e;
      color: #fff;
      padding: 12px 16px;
      font-weight: 700;
      font-size: 14px;
    }}
    .section-content {{
      padding: 16px;
    }}
    
    /* 2-column layout */
    .two-col {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
      margin-bottom: 16px;
    }}
    
    /* Overview box */
    .overview-text {{
      font-size: 13px;
      line-height: 1.7;
      margin-bottom: 12px;
    }}
    .info-box {{
      background: #e4edf9;
      border-left: 3px solid #003b6e;
      padding: 12px;
      margin-bottom: 12px;
      border-radius: 4px;
      font-size: 12px;
    }}
    .info-box .label {{
      font-weight: 700;
      color: #003b6e;
      margin-bottom: 2px;
    }}
    .info-box .value {{
      font-weight: 700;
      color: #0066cc;
    }}
    
    /* Stats row */
    .stats-row {{
      display: grid;
      grid-template-columns: repeat(6, 1fr);
      gap: 12px;
      margin-bottom: 16px;
    }}
    .stat-card {{
      background: #f0f3f8;
      padding: 12px;
      text-align: center;
      border-radius: 6px;
      border-top: 3px solid #0066cc;
    }}
    .stat-card .icon {{ font-size: 24px; margin-bottom: 4px; }}
    .stat-card .number {{ font-weight: 700; font-size: 16px; color: #003b6e; }}
    .stat-card .label {{ font-size: 11px; color: #666; margin-top: 4px; }}
    
    /* Timeline */
    .timeline {{
      display: flex;
      height: 40px;
      border-radius: 6px;
      overflow: hidden;
      margin-bottom: 16px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }}
    .timeline-phase {{
      display: flex;
      align-items: center;
      justify-content: center;
      color: #fff;
      font-size: 11px;
      font-weight: 700;
    }}
    .phase-0 {{ background: #9B59B6; flex: 2%; }}
    .phase-1 {{ background: #3498DB; flex: 8%; }}
    .phase-2 {{ background: #2ECC71; flex: 15%; }}
    .phase-3 {{ background: #E67E22; flex: 45%; }}
    .phase-4 {{ background: #E74C3C; flex: 30%; }}
    
    /* Milestones table */
    table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
      margin-bottom: 12px;
    }}
    th {{
      background: #f0f3f8;
      padding: 8px;
      text-align: left;
      font-weight: 700;
      color: #003b6e;
      border-bottom: 1px solid #ddd;
    }}
    td {{
      padding: 8px;
      border-bottom: 1px solid #eee;
    }}
    tbody tr:nth-child(odd) {{
      background: #fafbfc;
    }}
    
    /* Status pills */
    .pill {{
      display: inline-block;
      padding: 2px 8px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: 700;
      color: #fff;
    }}
    .pill-green {{ background: #007A33; }}
    .pill-amber {{ background: #E8A000; color: #1a1a1a; }}
    .pill-red {{ background: #CC0000; }}
    
    /* Risk grid */
    .risk-grid {{
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 12px;
      margin-bottom: 12px;
    }}
    .risk-card {{
      background: #f0f3f8;
      padding: 12px;
      border-radius: 6px;
      border-left: 3px solid;
      text-align: center;
    }}
    .risk-card.high {{ border-color: #CC0000; }}
    .risk-card.med {{ border-color: #E8A000; }}
    .risk-card.low {{ border-color: #007A33; }}
    .risk-card .count {{ font-size: 20px; font-weight: 700; margin-bottom: 4px; }}
    .risk-card .label {{ font-size: 12px; color: #666; }}
    
    /* Workstream grid */
    .workstream-grid {{
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 12px;
      margin-bottom: 12px;
    }}
    .workstream {{
      background: #f0f3f8;
      padding: 12px;
      border-radius: 6px;
      border-top: 3px solid #0066cc;
      font-size: 12px;
    }}
    .workstream .title {{ font-weight: 700; color: #003b6e; margin-bottom: 6px; }}
    .workstream ul {{ margin-left: 16px; }}
    .workstream li {{
      margin-bottom: 3px;
      font-size: 11px;
    }}
    
    /* Footer */
    .footer {{
      background: #e4edf9;
      padding: 12px;
      margin-top: 20px;
      border-radius: 6px;
      font-size: 11px;
      color: #666;
      text-align: center;
    }}
    
    @media print {{
      body {{ background: #fff; }}
      .page {{ box-shadow: none; }}
      .page:nth-child(n+2) {{ page-break-before: always; }}
    }}
  </style>
</head>
<body>
  <!-- PAGE 1 -->
  <div class="page">
    <!-- Header -->
    <div class="header">
      <div class="header-left">
        <div class="logo-container">
          <img src="data:image/png;base64,{logo_b64}" alt="Bosch — Invented for Life" />
        </div>
        <h1>Project Zebra</h1>
        <p>Packaging Business Carve-Out | 37 Global Sites | Stand Alone + TSA</p>
      </div>
      <div class="header-right">
        <strong>Executive Dashboard</strong><br/>
        Dashboard Date: 04 April 2026<br/>
        <div class="countdown">{days_until(GOLIVE_DATE)} days to GoLive</div>
      </div>
    </div>
    
    <!-- Countdown Strip -->
    <div class="countdown-strip">
      <div class="countdown-box">
        <div class="label">Project Kickoff</div>
        <div class="value">1</div>
        <div class="unit">days from start (today)</div>
      </div>
      <div class="countdown-box">
        <div class="label">GoLive Day 1</div>
        <div class="value">{days_until(GOLIVE_DATE)}</div>
        <div class="unit">days to go</div>
      </div>
      <div class="countdown-box">
        <div class="label">Project Completion</div>
        <div class="value">{days_until(COMPLETE_DATE)}</div>
        <div class="unit">days to closure</div>
      </div>
    </div>
    
    <!-- Project Overview -->
    <div class="section">
      <div class="section-header">PROJECT OVERVIEW</div>
      <div class="section-content">
        <div class="two-col">
          <div>
            <div class="overview-text">
              Zebra executes the complete separation of the Packaging business from Robert Bosch GmbH's global IT infrastructure. The project spans 37 worldwide sites with 3,500+ IT users, 208 integrated applications (including SAP ERP), and complex multi-geography supply-chain systems. Full separation and operational handover to the Buyer must complete by GoLive Day 1 (1 June 2027), followed by Transition Services Agreement (TSA) support through October 2027.
            </div>
          </div>
          <div>
            <div class="info-box">
              <div class="label">CARVE-OUT MODEL</div>
              <div class="value">Stand Alone + 6-Month TSA</div>
            </div>
            <div class="info-box">
              <div class="label">KEY PARTIES</div>
              <div class="value">Seller: Robert Bosch GmbH<br/>Buyer: [Confidential - Legal Hold]</div>
            </div>
            <div class="info-box">
              <div class="label">BUDGET BASELINE</div>
              <div class="value">€2.34M (Labour + CAPEX)</div>
            </div>
            <div class="info-box">
              <div class="label">PROGRAMME LEAD</div>
              <div class="value">Gill Amandeep Singh (BD/MIL-PSM1)<br/>PMO: KPMG</div>
            </div>
          </div>
        </div>
      </div>
    </div>
    
    <!-- Key Stats -->
    <div class="section">
      <div class="section-header">SCOPE & SCALE</div>
      <div class="section-content">
        <div class="stats-row">
          <div class="stat-card">
            <div class="icon">🌍</div>
            <div class="number">37</div>
            <div class="label">Worldwide Sites</div>
          </div>
          <div class="stat-card">
            <div class="icon">👥</div>
            <div class="number">3.5K+</div>
            <div class="label">IT Users</div>
          </div>
          <div class="stat-card">
            <div class="icon">💻</div>
            <div class="number">208</div>
            <div class="label">Applications</div>
          </div>
          <div class="stat-card">
            <div class="icon">📊</div>
            <div class="number">SAP</div>
            <div class="label">ERP System</div>
          </div>
          <div class="stat-card">
            <div class="icon">📅</div>
            <div class="number">20</div>
            <div class="label">Months Duration</div>
          </div>
          <div class="stat-card">
            <div class="icon">🤝</div>
            <div class="number">6</div>
            <div class="label">TSA Months</div>
          </div>
        </div>
      </div>
    </div>
    
    <!-- Timeline -->
    <div class="section">
      <div class="section-header">PROJECT TIMELINE & PHASES</div>
      <div class="section-content">
        <div class="timeline">
          <div class="timeline-phase phase-0" title="Phase 0: Initialization">P0</div>
          <div class="timeline-phase phase-1" title="Phase 1: Concept">Phase 1</div>
          <div class="timeline-phase phase-2" title="Phase 2: Design">Phase 2</div>
          <div class="timeline-phase phase-3" title="Phase 3: Build & Test">Phase 3</div>
          <div class="timeline-phase phase-4" title="Phase 4: GoLive & Closure">Phase 4</div>
        </div>
        <table>
          <thead>
            <tr>
              <th>milestone</th>
              <th>date</th>
              <th>Days from Today</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>QG0 - Initialization Gate</td>
              <td>17 Apr 2026</td>
              <td>{days_until(datetime(2026, 4, 17))}</td>
              <td><span class="pill pill-green">Upcoming</span></td>
            </tr>
            <tr>
              <td>QG1 - Concept Gate</td>
              <td>18 May 2026</td>
              <td>{days_until(datetime(2026, 5, 18))}</td>
              <td><span class="pill pill-green">Planned</span></td>
            </tr>
            <tr>
              <td>QG2/3 - Design & Build Gate</td>
              <td>27 Jul 2026</td>
              <td>{days_until(datetime(2026, 7, 27))}</td>
              <td><span class="pill pill-green">Planned</span></td>
            </tr>
            <tr>
              <td>QG4 - Pre-GoLive Quality Gate</td>
              <td>29 May 2027</td>
              <td>{days_until(datetime(2027, 5, 29))}</td>
              <td><span class="pill pill-amber">Critical</span></td>
            </tr>
            <tr>
              <td><strong>GoLive Day 1</strong></td>
              <td><strong>1 Jun 2027</strong></td>
              <td><strong>{days_until(GOLIVE_DATE)}</strong></td>
              <td><span class="pill pill-red">Critical Gate</span></td>
            </tr>
            <tr>
              <td>QG5 - Programme Closure</td>
              <td>31 Oct 2027</td>
              <td>{days_until(COMPLETE_DATE)}</td>
              <td><span class="pill pill-green">Final Gate</span></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
    
  </div>
  
  <!-- PAGE 2 -->
  <div class="page">
    
    <!-- Workstream Coverage -->
    <div class="section">
      <div class="section-header">IT WORKSTREAM COVERAGE (9 WORKSTREAMS)</div>
      <div class="section-content">
        <div class="workstream-grid">
          <div class="workstream">
            <div class="title">1. SAP ERP Separation</div>
            <ul>
              <li>System instance copy</li>
              <li>Master data segregation</li>
              <li>Chart of accounts split</li>
              <li>Confidence: <span class="pill pill-amber">Amber</span></li>
            </ul>
          </div>
          <div class="workstream">
            <div class="title">2. Application Portfolio (208 Apps)</div>
            <ul>
              <li>Transition/retain decisions</li>
              <li>ISV licensing negotiation</li>
              <li>Integration testing</li>
              <li>Confidence: <span class="pill pill-amber">Amber</span></li>
            </ul>
          </div>
          <div class="workstream">
            <div class="title">3. Data Migration</div>
            <ul>
              <li>15+ years master data</li>
              <li>Dry-run validation (3x)</li>
              <li>Post-cutover reconciliation</li>
              <li>Confidence: <span class="pill pill-green">Green</span></li>
            </ul>
          </div>
          <div class="workstream">
            <div class="title">4. Infrastructure & Network</div>
            <ul>
              <li>37-site global footprint</li>
              <li>WAN interconnects</li>
              <li>Cloud capacity planning</li>
              <li>Confidence: <span class="pill pill-amber">Amber</span></li>
            </ul>
          </div>
          <div class="workstream">
            <div class="title">5. Security & IAM</div>
            <ul>
              <li>Data residency rules</li>
              <li>GDPR compliance</li>
              <li>Separation controls</li>
              <li>Confidence: <span class="pill pill-amber">Amber</span></li>
            </ul>
          </div>
          <div class="workstream">
            <div class="title">6. Testing & QA</div>
            <ul>
              <li>37-site UAT coordination</li>
              <li>SAP transaction validation</li>
              <li>Dress rehearsal</li>
              <li>Confidence: <span class="pill pill-green">Green</span></li>
            </ul>
          </div>
          <div class="workstream">
            <div class="title">7. Change Management</div>
            <ul>
              <li>3,500+ user training</li>
              <li>Change readiness</li>
              <li>Stakeholder comms</li>
              <li>Confidence: <span class="pill pill-amber">Amber</span></li>
            </ul>
          </div>
          <div class="workstream">
            <div class="title">8. TSA & Handover</div>
            <ul>
              <li>Service level definition</li>
              <li>Exit criteria alignment</li>
              <li>Seller IT rundown</li>
              <li>Confidence: <span class="pill pill-amber">Amber</span></li>
            </ul>
          </div>
          <div class="workstream">
            <div class="title">9. Programme Management</div>
            <ul>
              <li>Schedule tracking</li>
              <li>Risk management</li>
              <li>Governance gates</li>
              <li>Confidence: <span class="pill pill-green">Green</span></li>
            </ul>
          </div>
        </div>
      </div>
    </div>
    
    <!-- Risk Indicators -->
    <div class="section">
      <div class="section-header">TOP RISK INDICATORS</div>
      <div class="section-content">
        <div class="risk-grid">
          <div class="risk-card high">
            <div class="count">3</div>
            <div class="label">High-Risk (P×I ≥ 16)</div>
          </div>
          <div class="risk-card med">
            <div class="count">7</div>
            <div class="label">Medium-Risk (8 ≤ P×I < 16)</div>
          </div>
          <div class="risk-card low">
            <div class="count">5</div>
            <div class="label">Low-Risk (P×I < 8)</div>
          </div>
        </div>
        <table>
          <thead>
            <tr>
              <th>TOP RISKS (P×I ≥ 16)</th>
              <th>Rating</th>
              <th>Mitigation Lead</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Multi-Site Cutover Coordination Failure (37 sites)</td>
              <td><span class="pill pill-red">P5×I5=25</span></td>
              <td>KPMG PM + Ops Teams</td>
            </tr>
            <tr>
              <td>SAP Data Separation Incomplete</td>
              <td><span class="pill pill-red">P4×I4=16</span></td>
              <td>KPMG SAP CoE</td>
            </tr>
            <tr>
              <td>Buyer Infrastructure Readiness Gap</td>
              <td><span class="pill pill-red">P4×I4=16</span></td>
              <td>Buyer IT + KPMG Infra</td>
            </tr>
            <tr>
              <td>208-App Portfolio Transition Delays</td>
              <td><span class="pill pill-amber">P4×I4=16</span></td>
              <td>KPMG IT Architecture</td>
            </tr>
            <tr>
              <td>Data Residency / GDPR Compliance</td>
              <td><span class="pill pill-amber">P3×I5=15</span></td>
              <td>KPMG + InfoSec</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
    
    <!-- Budget Summary -->
    <div class="section">
      <div class="section-header">BUDGET SUMMARY</div>
      <div class="section-content">
        <div class="two-col">
          <div>
            <div class="info-box">
              <div class="label">TOTAL PROJECT BUDGET</div>
              <div class="value">€ 2.34 Million</div>
            </div>
            <table>
              <thead>
                <tr><th>Category</th><th>Amount</th></tr>
              </thead>
              <tbody>
                <tr>
                  <td>KPMG PMO & Governance</td>
                  <td>€ 285K</td>
                </tr>
                <tr>
                  <td>KPMG SAP & ERP Build</td>
                  <td>€ 630K</td>
                </tr>
                <tr>
                  <td>KPMG Testing & QA</td>
                  <td>€ 495K</td>
                </tr>
                <tr>
                  <td>KPMG Change Mgmt & Training</td>
                  <td>€ 280K</td>
                </tr>
                <tr>
                  <td>KPMG Data & Integration</td>
                  <td>€ 330K</td>
                </tr>
                <tr>
                  <td>RoboGmbH IT Support</td>
                  <td>€ 228K</td>
                </tr>
                <tr style="font-weight:700; border-top:2px solid #ccc;">
                  <td>Subtotal (Labor)</td>
                  <td>€ 2.12M</td>
                </tr>
              </tbody>
            </table>
          </div>
          <div>
            <table>
              <thead>
                <tr><th>CAPEX</th><th>Amount</th></tr>
              </thead>
              <tbody>
                <tr>
                  <td>SAP Audit & Validation</td>
                  <td>€ 35K</td>
                </tr>
                <tr>
                  <td>App Tools & Licenses</td>
                  <td>€ 25K</td>
                </tr>
                <tr>
                  <td>Regulatory Compliance Audit</td>
                  <td>€ 40K</td>
                </tr>
                <tr>
                  <td>Parallel Run Environment</td>
                  <td>€ 20K</td>
                </tr>
                <tr>
                  <td>Contingency (5% of labor)</td>
                  <td>€ 106K</td>
                </tr>
                <tr style="font-weight:700; border-top:2px solid #ccc;">
                  <td>CAPEX Subtotal</td>
                  <td>€ 226K</td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
    
  </div>
  
  <!-- Footer -->
  <div style="padding: 8px 20px; background: #f4f6f9; font-size: 11px; color: #666;">
    <strong>Project Zebra | Executive Dashboard</strong> | Report Date: 04 April 2026 | Sources: Project Schedule, Risk Register, Cost Plan | Confidential — Internal Use Only
  </div>
  
</body>
</html>
"""

OUT.write_text(html, encoding="utf-8")
print(f"✓ Executive dashboard written: {OUT}")
