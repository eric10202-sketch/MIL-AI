#!/usr/bin/env python3
"""
Generate Trinity-CAM Management KPI Dashboard (HTML).
Follows management-kpi-dashboard-generation SKILL.md.
All metrics derived from Trinity-CAM schedule/cost/risk data.
"""

import base64
from datetime import date
from pathlib import Path

HERE   = Path(__file__).parent
LOGO   = HERE / "Bosch.png"
OUTPUT = HERE / "active-projects" / "Trinity-CAM" / "Trinity-CAM_Management_KPI_Dashboard.html"
OUTPUT.parent.mkdir(parents=True, exist_ok=True)

logo_b64 = base64.b64encode(LOGO.read_bytes()).decode() if LOGO.exists() else ""
logo_tag  = f'<img src="data:image/png;base64,{logo_b64}" alt="Bosch" style="height:36px;display:block;" />' if logo_b64 else ""

TODAY      = date.today()
KICKOFF    = date(2026, 7, 1)
QG1        = date(2026, 10, 1)
QG23       = date(2027, 7, 31)
QG4        = date(2027, 12, 8)
GOLIVE     = date(2028, 1, 1)
COMPLETION = date(2028, 4, 1)
REPORT     = TODAY.strftime("%d %B %Y")

def ddays(d):
    n = (d - TODAY).days
    if n > 0:   return f"{n} days"
    elif n < 0: return f"{abs(n)}d ago"
    else:       return "TODAY"

HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Trinity-CAM Management KPI Dashboard</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:'Segoe UI',Arial,sans-serif;font-size:12px;color:#1a1a1a;background:#f4f6f9;}}

/* --- HEADER --- */
.hdr{{background:linear-gradient(135deg,#003b6e 0%,#005199 100%);color:#fff;padding:14px 28px;display:flex;align-items:center;gap:20px;}}
.bosch-logo{{display:flex;align-items:center;background:#fff;padding:4px 8px;border-radius:4px;}}
.hdr-center{{flex:1;}}
.hdr-center h1{{font-size:17px;font-weight:700;letter-spacing:0.4px;}}
.hdr-center h2{{font-size:11px;font-weight:400;opacity:.85;margin-top:3px;}}
.hdr-right{{text-align:right;font-size:11px;opacity:.85;}}
.hdr-right strong{{font-size:13px;display:block;}}

/* --- RAG STRIP --- */
.rag-strip{{background:#003b6e;display:flex;gap:0;border-top:2px solid rgba(255,255,255,.1);}}
.rag-box{{flex:1;border-right:1px solid rgba(255,255,255,.12);padding:7px 14px;}}
.rag-box:last-child{{border-right:none;}}
.rag-label{{font-size:9px;color:#fff;opacity:.7;text-transform:uppercase;letter-spacing:.4px;}}
.rag-val{{font-size:13px;font-weight:700;}}
.rag-sub{{font-size:9px;color:#fff;opacity:.6;}}
.rv-g{{color:#2ecc71;}} .rv-a{{color:#f1c40f;}} .rv-r{{color:#e74c3c;}} .rv-b{{color:#5dade2;}}

/* --- PAGE --- */
.page{{max-width:1200px;margin:0 auto;padding:14px 18px 30px;}}

/* --- 12-COL GRID SYSTEM --- */
.grid{{display:grid;grid-template-columns:repeat(12,1fr);gap:12px;margin-bottom:12px;}}
.col-3{{grid-column:span 3;}}
.col-4{{grid-column:span 4;}}
.col-5{{grid-column:span 5;}}
.col-6{{grid-column:span 6;}}
.col-7{{grid-column:span 7;}}
.col-8{{grid-column:span 8;}}
.col-12{{grid-column:span 12;}}

/* --- KPI CARD --- */
.kpi-card{{background:#fff;border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.08);overflow:hidden;}}
.kpi-title{{background:#005199;color:#fff;font-size:10px;font-weight:700;padding:6px 12px;letter-spacing:.4px;text-transform:uppercase;}}
.kpi-body{{padding:10px 12px;}}

/* --- GAUGE RING (CSS) --- */
.gauge-wrap{{text-align:center;padding:4px 0 8px;}}
.gauge-ring{{display:inline-flex;align-items:center;justify-content:center;width:80px;height:80px;border-radius:50%;font-size:18px;font-weight:700;color:#fff;margin:0 auto;}}
.g-green{{background:conic-gradient(#007A33 var(--pct),#e8ecf2 0);color:#007A33;background-clip:border-box;}}
.gauge-inner{{background:#fff;width:60px;height:60px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:700;color:#1a1a1a;}}
.gauge-label{{font-size:10px;color:#666;margin-top:4px;}}

/* --- CONFIDENCE BAR --- */
.conf-bar-wrap{{margin:4px 0;}}
.conf-bar-label{{display:flex;justify-content:space-between;font-size:10px;margin-bottom:2px;}}
.conf-bar-outer{{height:10px;background:#e8ecf2;border-radius:5px;overflow:hidden;}}
.conf-bar-inner{{height:100%;border-radius:5px;}}

/* --- MILESTONE TIMELINE --- */
.ms-timeline{{position:relative;padding:8px 0;}}
.ms-line{{position:absolute;top:22px;left:12px;right:12px;height:3px;background:#dee3ed;border-radius:2px;}}
.ms-nodes{{display:flex;justify-content:space-between;position:relative;padding:0 6px;}}
.ms-node{{text-align:center;width:80px;}}
.ms-dot{{width:20px;height:20px;border-radius:50%;margin:0 auto 4px;border:3px solid #fff;box-shadow:0 0 0 2px #005199;background:#005199;}}
.ms-dot.done{{background:#007A33;box-shadow:0 0 0 2px #007A33;}}
.ms-dot.active{{background:#E8A000;box-shadow:0 0 0 2px #E8A000;}}
.ms-node-name{{font-size:9px;font-weight:700;color:#003b6e;}}
.ms-node-date{{font-size:8px;color:#888;}}
.ms-node-cd{{font-size:8px;color:#0066CC;font-weight:600;}}

/* --- TABLES --- */
table.kpi-tbl{{width:100%;border-collapse:collapse;font-size:10px;}}
.kpi-tbl th{{background:#003b6e;color:#fff;padding:5px 7px;text-align:left;font-size:9px;font-weight:600;}}
.kpi-tbl td{{padding:5px 7px;border-bottom:1px solid #e8ecf2;vertical-align:top;}}
.kpi-tbl tr:nth-child(even){{background:#EFF4FB;}}

/* --- PILL --- */
.pill{{display:inline-block;padding:2px 6px;border-radius:10px;font-size:8px;font-weight:700;}}
.p-g{{background:#e9f7ef;color:#007A33;}} .p-a{{background:#fef9e7;color:#b7770d;}} .p-r{{background:#fde;color:#CC0000;}}
.p-blue{{background:#EFF4FB;color:#005199;}}

/* --- SCORE BADGE --- */
.score-badge{{display:inline-block;width:22px;height:22px;border-radius:50%;font-size:10px;font-weight:700;text-align:center;line-height:22px;color:#fff;}}
.sb-r{{background:#e74c3c;}} .sb-a{{background:#f39c12;}} .sb-g{{background:#27ae60;}}

/* --- 90-DAY ACTIONS --- */
.action-row{{display:flex;gap:8px;padding:5px 0;border-bottom:1px solid #e8ecf2;align-items:flex-start;font-size:10px;}}
.action-row:last-child{{border-bottom:none;}}
.action-cat{{min-width:90px;font-weight:600;color:#005199;}}
.action-text{{flex:1;color:#333;}}
.action-due{{min-width:75px;text-align:right;font-size:9px;color:#888;}}

/* --- MODEL DIFF TABLE --- */
.model-tbl td:first-child{{font-weight:600;color:#003b6e;width:30%;}}

/* --- FOOTER --- */
.footer{{text-align:center;font-size:9px;color:#aaa;padding:8px 0 16px;}}
</style>
</head>
<body>

<!-- ========== HEADER ========== -->
<div class="hdr">
  <div class="bosch-logo">{logo_tag}</div>
  <div class="hdr-center">
    <h1>Trinity-CAM &mdash; Management KPI Dashboard</h1>
    <h2>JCI Aircondition &rarr; Merger Zone (Infosys) &rarr; Robert Bosch GmbH &nbsp;|&nbsp; Integration Model &nbsp;|&nbsp; 48 Sites &nbsp;|&nbsp; 12,000 Users</h2>
  </div>
  <div class="hdr-right">
    <strong>{REPORT}</strong>
    GoLive: 1 Jan 2028 &nbsp;({ddays(GOLIVE)})
  </div>
</div>

<!-- RAG STRIP -->
<div class="rag-strip">
  <div class="rag-box"><div class="rag-label">Schedule (SPI)</div><div class="rag-val rv-a">SPI 1.00</div><div class="rag-sub">On plan — QG0 baseline</div></div>
  <div class="rag-box"><div class="rag-label">Cost (CPI)</div><div class="rag-val rv-a">CPI 1.00</div><div class="rag-sub">On budget — QG0 baseline</div></div>
  <div class="rag-box"><div class="rag-label">Day 1 Readiness</div><div class="rag-val rv-a">12%</div><div class="rag-sub">Initiation phase</div></div>
  <div class="rag-box"><div class="rag-label">TSA Confidence</div><div class="rag-val rv-a">AMBER</div><div class="rag-sub">TSA expired; R014 open</div></div>
  <div class="rag-box"><div class="rag-label">Top Risk</div><div class="rag-val rv-r">Score 20</div><div class="rag-sub">R001 SAP complexity</div></div>
  <div class="rag-box"><div class="rag-label">Overall RAG</div><div class="rag-val rv-a">AMBER</div><div class="rag-sub">Baseline programme</div></div>
</div>

<!-- ========== PAGE BODY ========== -->
<div class="page">

<!-- ROW 1: SPI | CPI | Day-1 Readiness | TSA Confidence -->
<div class="grid">

  <div class="kpi-card col-3">
    <div class="kpi-title">Schedule Performance (SPI)</div>
    <div class="kpi-body">
      <div class="gauge-wrap">
        <div style="position:relative;display:inline-block;width:80px;height:80px;">
          <svg viewBox="0 0 80 80" width="80" height="80">
            <circle cx="40" cy="40" r="34" fill="none" stroke="#e8ecf2" stroke-width="8"/>
            <circle cx="40" cy="40" r="34" fill="none" stroke="#E8A000" stroke-width="8"
              stroke-dasharray="213.6" stroke-dashoffset="0" stroke-linecap="round" transform="rotate(-90 40 40)"/>
          </svg>
          <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);font-size:16px;font-weight:700;color:#1a1a1a;">1.00</div>
        </div>
      </div>
      <div class="gauge-label" style="text-align:center;">SPI = EV / PV</div>
      <div style="margin-top:8px;font-size:10px;">
        <div style="display:flex;justify-content:space-between;"><span>Planned Value (PV)</span><span>EUR 629K</span></div>
        <div style="display:flex;justify-content:space-between;"><span>Earned Value (EV)</span><span>EUR 629K</span></div>
        <div style="display:flex;justify-content:space-between;font-weight:600;margin-top:3px;"><span>SPI Status</span><span class="pill p-a">ON PLAN</span></div>
      </div>
    </div>
  </div>

  <div class="kpi-card col-3">
    <div class="kpi-title">Cost Performance (CPI)</div>
    <div class="kpi-body">
      <div class="gauge-wrap">
        <div style="position:relative;display:inline-block;width:80px;height:80px;">
          <svg viewBox="0 0 80 80" width="80" height="80">
            <circle cx="40" cy="40" r="34" fill="none" stroke="#e8ecf2" stroke-width="8"/>
            <circle cx="40" cy="40" r="34" fill="none" stroke="#E8A000" stroke-width="8"
              stroke-dasharray="213.6" stroke-dashoffset="0" stroke-linecap="round" transform="rotate(-90 40 40)"/>
          </svg>
          <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);font-size:16px;font-weight:700;color:#1a1a1a;">1.00</div>
        </div>
      </div>
      <div class="gauge-label" style="text-align:center;">CPI = EV / AC</div>
      <div style="margin-top:8px;font-size:10px;">
        <div style="display:flex;justify-content:space-between;"><span>Budget at Completion</span><span>EUR 7.87M</span></div>
        <div style="display:flex;justify-content:space-between;"><span>Actual Cost (AC)</span><span>EUR 629K</span></div>
        <div style="display:flex;justify-content:space-between;font-weight:600;margin-top:3px;"><span>CPI Status</span><span class="pill p-a">ON BUDGET</span></div>
      </div>
    </div>
  </div>

  <div class="kpi-card col-3">
    <div class="kpi-title">Day 1 GoLive Readiness</div>
    <div class="kpi-body">
      <div class="gauge-wrap">
        <div style="position:relative;display:inline-block;width:80px;height:80px;">
          <svg viewBox="0 0 80 80" width="80" height="80">
            <circle cx="40" cy="40" r="34" fill="none" stroke="#e8ecf2" stroke-width="8"/>
            <circle cx="40" cy="40" r="34" fill="none" stroke="#0066CC" stroke-width="8"
              stroke-dasharray="213.6" stroke-dashoffset="186" stroke-linecap="round" transform="rotate(-90 40 40)"/>
          </svg>
          <div style="position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);font-size:16px;font-weight:700;color:#1a1a1a;">12%</div>
        </div>
      </div>
      <div class="gauge-label" style="text-align:center;">Based on completed milestones</div>
      <div style="margin-top:8px;font-size:10px;">
        <div style="display:flex;justify-content:space-between;"><span>Gates Passed</span><span>0 / 6</span></div>
        <div style="display:flex;justify-content:space-between;"><span>Phase</span><span>Initiation</span></div>
        <div style="display:flex;justify-content:space-between;font-weight:600;margin-top:3px;"><span>Target GoLive</span><span>01 Jan 2028</span></div>
      </div>
    </div>
  </div>

  <div class="kpi-card col-3">
    <div class="kpi-title">TSA &amp; Integration Confidence</div>
    <div class="kpi-body">
      <div style="text-align:center;padding:6px 0 8px;">
        <div style="font-size:28px;font-weight:700;color:#E8A000;">AMBER</div>
        <div style="font-size:10px;color:#666;margin-top:3px;">TSA expired 30 Jun 2026</div>
      </div>
      <div style="font-size:10px;margin-top:6px;">
        <div style="display:flex;justify-content:space-between;margin:3px 0;"><span>JCI TSA Status</span><span class="pill p-a">EXPIRED</span></div>
        <div style="display:flex;justify-content:space-between;margin:3px 0;"><span>Infosys SOW</span><span class="pill p-g">SIGNED</span></div>
        <div style="display:flex;justify-content:space-between;margin:3px 0;"><span>MZ Architecture</span><span class="pill p-a">PENDING QG1</span></div>
        <div style="display:flex;justify-content:space-between;margin:3px 0;"><span>GDPR Clearance</span><span class="pill p-a">IN PROGRESS</span></div>
        <div style="display:flex;justify-content:space-between;margin:3px 0;"><span>Integration Model</span><span class="pill p-blue">INTEGRATION</span></div>
      </div>
    </div>
  </div>

</div><!-- /ROW 1 -->

<!-- ROW 2: Workstream Confidence | Milestone Controls -->
<div class="grid">

  <div class="kpi-card col-5">
    <div class="kpi-title">Workstream Confidence Scores</div>
    <div class="kpi-body">
      <div class="conf-bar-wrap"><div class="conf-bar-label"><span>WS1 MZ Infrastructure</span><span style="color:#007A33;font-weight:600;">82%</span></div><div class="conf-bar-outer"><div class="conf-bar-inner" style="width:82%;background:#007A33;"></div></div></div>
      <div class="conf-bar-wrap"><div class="conf-bar-label"><span>WS2 SAP Migration</span><span style="color:#E8A000;font-weight:600;">55%</span></div><div class="conf-bar-outer"><div class="conf-bar-inner" style="width:55%;background:#E8A000;"></div></div></div>
      <div class="conf-bar-wrap"><div class="conf-bar-label"><span>WS3 Application Migration (1,800+)</span><span style="color:#E8A000;font-weight:600;">60%</span></div><div class="conf-bar-outer"><div class="conf-bar-inner" style="width:60%;background:#E8A000;"></div></div></div>
      <div class="conf-bar-wrap"><div class="conf-bar-label"><span>WS4 End-User Workplace (12,000)</span><span style="color:#007A33;font-weight:600;">80%</span></div><div class="conf-bar-outer"><div class="conf-bar-inner" style="width:80%;background:#007A33;"></div></div></div>
      <div class="conf-bar-wrap"><div class="conf-bar-label"><span>WS5 Identity &amp; Access Management</span><span style="color:#007A33;font-weight:600;">85%</span></div><div class="conf-bar-outer"><div class="conf-bar-inner" style="width:85%;background:#007A33;"></div></div></div>
      <div class="conf-bar-wrap"><div class="conf-bar-label"><span>WS6 Security &amp; Compliance (GDPR)</span><span style="color:#E8A000;font-weight:600;">62%</span></div><div class="conf-bar-outer"><div class="conf-bar-inner" style="width:62%;background:#E8A000;"></div></div></div>
      <div class="conf-bar-wrap"><div class="conf-bar-label"><span>WS7 Data Migration</span><span style="color:#007A33;font-weight:600;">78%</span></div><div class="conf-bar-outer"><div class="conf-bar-inner" style="width:78%;background:#007A33;"></div></div></div>
      <div class="conf-bar-wrap"><div class="conf-bar-label"><span>WS8 TSA Exit &amp; HR/Legal</span><span style="color:#E8A000;font-weight:600;">58%</span></div><div class="conf-bar-outer"><div class="conf-bar-inner" style="width:58%;background:#E8A000;"></div></div></div>
      <div class="conf-bar-wrap"><div class="conf-bar-label"><span>WS9 Programme Control &amp; Hypercare</span><span style="color:#007A33;font-weight:600;">88%</span></div><div class="conf-bar-outer"><div class="conf-bar-inner" style="width:88%;background:#007A33;"></div></div></div>
    </div>
  </div>

  <div class="kpi-card col-7">
    <div class="kpi-title">Milestone Gate Control Timeline</div>
    <div class="kpi-body">
      <div class="ms-timeline">
        <div class="ms-line"></div>
        <div class="ms-nodes">
          <div class="ms-node">
            <div class="ms-dot active"></div>
            <div class="ms-node-name">QG0</div>
            <div class="ms-node-date">01 Jul 2026</div>
            <div class="ms-node-cd" style="color:#E8A000;">{ddays(KICKOFF)}</div>
          </div>
          <div class="ms-node">
            <div class="ms-dot"></div>
            <div class="ms-node-name">QG1</div>
            <div class="ms-node-date">01 Oct 2026</div>
            <div class="ms-node-cd">{ddays(QG1)}</div>
          </div>
          <div class="ms-node">
            <div class="ms-dot"></div>
            <div class="ms-node-name">QG2&amp;3</div>
            <div class="ms-node-date">31 Jul 2027</div>
            <div class="ms-node-cd">{ddays(QG23)}</div>
          </div>
          <div class="ms-node">
            <div class="ms-dot"></div>
            <div class="ms-node-name">QG4</div>
            <div class="ms-node-date">08 Dec 2027</div>
            <div class="ms-node-cd">{ddays(QG4)}</div>
          </div>
          <div class="ms-node">
            <div class="ms-dot"></div>
            <div class="ms-node-name">GoLive</div>
            <div class="ms-node-date">01 Jan 2028</div>
            <div class="ms-node-cd">{ddays(GOLIVE)}</div>
          </div>
          <div class="ms-node">
            <div class="ms-dot"></div>
            <div class="ms-node-name">QG5</div>
            <div class="ms-node-date">01 Apr 2028</div>
            <div class="ms-node-cd">{ddays(COMPLETION)}</div>
          </div>
        </div>
      </div>
      <table class="kpi-tbl" style="margin-top:10px;">
        <tr><th>Gate</th><th>Date</th><th>Countdown</th><th>Key Criterion</th><th>Status</th></tr>
        <tr><td><b>QG0</b></td><td>01 Jul 2026</td><td>{ddays(KICKOFF)}</td><td>PMO + Infosys mobilised; Steering constituted</td><td><span class="pill p-a">ACTIVE</span></td></tr>
        <tr><td><b>QG1</b></td><td>01 Oct 2026</td><td>{ddays(QG1)}</td><td>App inventory complete; MZ arch approved; CAPEX budget</td><td><span class="pill p-blue">PLANNED</span></td></tr>
        <tr><td><b>QG2&amp;3</b></td><td>31 Jul 2027</td><td>{ddays(QG23)}</td><td>MZ DC live; SAP copy complete; Wave 1 validated</td><td><span class="pill p-blue">PLANNED</span></td></tr>
        <tr><td><b>QG4</b></td><td>08 Dec 2027</td><td>{ddays(QG4)}</td><td>12,000 users + 1,800 apps migrated; Mock 2 passed</td><td><span class="pill p-blue">PLANNED</span></td></tr>
        <tr><td><b>GoLive</b></td><td>01 Jan 2028</td><td>{ddays(GOLIVE)}</td><td>Day 1 cutover; hypercare active; all P1 defects zero</td><td><span class="pill p-blue">PLANNED</span></td></tr>
        <tr><td><b>QG5</b></td><td>01 Apr 2028</td><td>{ddays(COMPLETION)}</td><td>90-day hypercare done; TSA exit; Bosch handover</td><td><span class="pill p-blue">PLANNED</span></td></tr>
      </table>
    </div>
  </div>

</div><!-- /ROW 2 -->

<!-- ROW 3: TOP RISKS + MODEL DIFFERENCES -->
<div class="grid">

  <div class="kpi-card col-7">
    <div class="kpi-title">Top Risk Register (from Trinity-CAM_Risk_Register.xlsx)</div>
    <div class="kpi-body" style="padding:8px;">
      <table class="kpi-tbl">
        <tr><th>ID</th><th>Risk Description</th><th>Category</th><th>P</th><th>I</th><th>Score</th><th>Owner</th><th>Status</th></tr>
        <tr>
          <td><b>R001</b></td>
          <td>SAP landscape complexity delays system copy; QG2&amp;3 at risk</td>
          <td>Technology</td><td>70%</td><td>VH</td>
          <td><span class="score-badge sb-r">20</span></td>
          <td>KPMG SAP Arch</td>
          <td><span class="pill p-r">OPEN</span></td>
        </tr>
        <tr>
          <td><b>R003</b></td>
          <td>SAP Mock Cutover 2 defects prevent QG4; GoLive delayed</td>
          <td>Schedule</td><td>50%</td><td>VH</td>
          <td><span class="score-badge sb-r">15</span></td>
          <td>KPMG PMO Lead</td>
          <td><span class="pill p-r">OPEN</span></td>
        </tr>
        <tr>
          <td><b>R002</b></td>
          <td>Infosys MZ delivery delay; Phase 3 start pushed</td>
          <td>Technology</td><td>50%</td><td>H</td>
          <td><span class="score-badge sb-a">12</span></td>
          <td>Infosys PM</td>
          <td><span class="pill p-a">MONITORED</span></td>
        </tr>
        <tr>
          <td><b>R007</b></td>
          <td>1,800+ app compatibility; wave capacity exceeded</td>
          <td>Technology</td><td>50%</td><td>H</td>
          <td><span class="score-badge sb-a">12</span></td>
          <td>Infosys App Lead</td>
          <td><span class="pill p-a">MONITORED</span></td>
        </tr>
        <tr>
          <td><b>R006</b></td>
          <td>GDPR/PII breach during 18-month migration; EUR 20M exposure</td>
          <td>Security</td><td>30%</td><td>VH</td>
          <td><span class="score-badge sb-a">10</span></td>
          <td>Infosys Sec Lead</td>
          <td><span class="pill p-a">MITIGATING</span></td>
        </tr>
        <tr>
          <td><b>R005</b></td>
          <td>TUPE non-compliance Germany/France blocks Wave 1 user migration</td>
          <td>Legal</td><td>30%</td><td>VH</td>
          <td><span class="score-badge sb-a">10</span></td>
          <td>JCI Legal</td>
          <td><span class="pill p-a">MITIGATING</span></td>
        </tr>
        <tr>
          <td><b>R004</b></td>
          <td>JCI IT cooperation falls short; data quality issues; ETL delays</td>
          <td>Governance</td><td>40%</td><td>H</td>
          <td><span class="score-badge sb-a">8</span></td>
          <td>KPMG PMO Lead</td>
          <td><span class="pill p-a">MONITORED</span></td>
        </tr>
        <tr>
          <td><b>R025</b></td>
          <td><em>OPPORTUNITY: MZ cloud modernisation saves EUR 3-5M/yr Bosch OPEX</em></td>
          <td>Technology</td><td>50%</td><td>H</td>
          <td><span class="score-badge sb-g">12</span></td>
          <td>Infosys Cloud</td>
          <td><span class="pill p-g">PURSUING</span></td>
        </tr>
      </table>
      <div style="font-size:9px;color:#888;margin-top:6px;">Total: 25 risks | 24 threats, 1 opportunity | Full register: Trinity-CAM_Risk_Register.xlsx</div>
    </div>
  </div>

  <div class="kpi-card col-5">
    <div class="kpi-title">Integration Model — Key Characteristics</div>
    <div class="kpi-body" style="padding:8px;">
      <table class="kpi-tbl model-tbl">
        <tr><th>Characteristic</th><th>Stand-Alone</th><th style="background:#005199;">Integration (Trinity-CAM)</th></tr>
        <tr><td>Merger Zone</td><td style="color:#888;">Not required</td><td style="color:#005199;font-weight:600;">Required — Infosys build &amp; operate</td></tr>
        <tr><td>Migration Path</td><td style="color:#888;">JCI direct</td><td style="color:#005199;font-weight:600;">JCI &rarr; MZ &rarr; Bosch IT</td></tr>
        <tr><td>TSA Role</td><td style="color:#888;">Extended TSA</td><td style="color:#005199;font-weight:600;">TSA expired; MZ provides continuity</td></tr>
        <tr><td>SAP Strategy</td><td style="color:#888;">New Bosch instance</td><td style="color:#005199;font-weight:600;">Copied JCI SAP lands on MZ then integrates</td></tr>
        <tr><td>Identity Strategy</td><td style="color:#888;">Separate AD</td><td style="color:#005199;font-weight:600;">MZ AD forest; federation then merge</td></tr>
        <tr><td>Site Connectivity</td><td style="color:#888;">Maintained</td><td style="color:#005199;font-weight:600;">New SD-WAN; 48 sites re-pointed to MZ</td></tr>
        <tr><td>Hypercare</td><td style="color:#888;">30 days</td><td style="color:#005199;font-weight:600;">90 days (Infosys L3); complex SAP estate</td></tr>
        <tr><td>Wave Strategy</td><td style="color:#888;">Direct migration</td><td style="color:#005199;font-weight:600;">3 app waves + parallel SAP track</td></tr>
        <tr><td>DR Architecture</td><td style="color:#888;">JCI DR remains</td><td style="color:#005199;font-weight:600;">MZ DR built by Infosys; validated pre-QG4</td></tr>
        <tr><td>GoLive Risk</td><td style="color:#888;">Lower (simpler)</td><td style="color:#005199;font-weight:600;">Higher — 1,800 apps, 12,000 users, SAP</td></tr>
      </table>
    </div>
  </div>

</div><!-- /ROW 3 -->

<!-- ROW 4: NEXT 90 DAYS + BUDGET CONTROL -->
<div class="grid">

  <div class="kpi-card col-8">
    <div class="kpi-title">Next 90-Day Action Forecast (Programme Control)</div>
    <div class="kpi-body" style="padding:8px 10px;">
      <div class="action-row">
        <div class="action-cat">Programme</div>
        <div class="action-text">KPMG Programme Director and all workstream leads confirmed and mobilised. Initial SteerCo meeting convened. Programme plan v1 baselined and distributed to JCI and Bosch sponsors.</div>
        <div class="action-due">Jul 2026 — QG0</div>
      </div>
      <div class="action-row">
        <div class="action-cat">Applications</div>
        <div class="action-text">Complete application inventory: catalogue all 1,800+ JCI Aircon applications. Assign ownership, category (SAP / core / non-core), and migration wave. Target completion by 15 Aug 2026.</div>
        <div class="action-due">Aug 2026</div>
      </div>
      <div class="action-row">
        <div class="action-cat">MZ Architecture</div>
        <div class="action-text">Infosys to submit Merger Zone architecture proposal (DC vs cloud, SD-WAN topology, security model) for KPMG technical review and Bosch steering approval at QG1.</div>
        <div class="action-due">Sep 2026</div>
      </div>
      <div class="action-row">
        <div class="action-cat">SAP</div>
        <div class="action-text">KPMG SAP Architect to complete SAP landscape discovery: system count, client topology, interface inventory, Z-programs, and data volumes. Deliver SAP migration strategy document for QG1.</div>
        <div class="action-due">Sep 2026</div>
      </div>
      <div class="action-row">
        <div class="action-cat">Legal / GDPR</div>
        <div class="action-text">Engage GDPR counsel in Germany, France, China, Japan, and USA. Establish legal basis for each data transfer. Complete Data Protection Impact Assessment (DPIA) framework for MZ data flow.</div>
        <div class="action-due">Sep 2026</div>
      </div>
      <div class="action-row">
        <div class="action-cat">CAPEX</div>
        <div class="action-text">Infosys to submit Merger Zone infrastructure cost model (CAPEX) to Bosch finance for QG1 approval. Includes DC hardware / cloud, WAN links, and software licences (M365, ITSM).</div>
        <div class="action-due">Sep 2026 — QG1</div>
      </div>
      <div class="action-row">
        <div class="action-cat">Risk</div>
        <div class="action-text">Escalate R001 (SAP complexity) response plan to steering: assign dedicated SAP copy environment, appoint Infosys SAP offshore lead, schedule 2 dry-run weeks before full system copy in Apr 2027.</div>
        <div class="action-due">Aug 2026</div>
      </div>
      <div class="action-row">
        <div class="action-cat">TSA</div>
        <div class="action-text">Formal TSA exit audit: confirm zero residual service dependencies. Any emergency clauses invoked must be registered as scope issues. Produce clean TSA exit report for QG1 sign-off.</div>
        <div class="action-due">Oct 2026 — QG1</div>
      </div>
    </div>
  </div>

  <div class="kpi-card col-4">
    <div class="kpi-title">Budget &amp; Cost Control</div>
    <div class="kpi-body">
      <div style="text-align:center;margin-bottom:10px;">
        <div style="font-size:22px;font-weight:700;color:#003b6e;">EUR 9,054,640</div>
        <div style="font-size:10px;color:#888;">Labour incl. 15% contingency | CAPEX TBC at QG1</div>
      </div>
      <table class="kpi-tbl">
        <tr><th>Phase</th><th>EUR</th><th>%</th></tr>
        <tr><td>Ph1 Initiation</td><td>629,888</td><td>8%</td></tr>
        <tr><td>Ph2 MZ Build</td><td>2,755,760</td><td>35%</td></tr>
        <tr><td>Ph3 Test &amp; Migrate</td><td>2,362,080</td><td>30%</td></tr>
        <tr><td>Ph4 Final Readiness</td><td>708,624</td><td>9%</td></tr>
        <tr><td>GoLive &amp; Cutover</td><td>236,208</td><td>3%</td></tr>
        <tr><td>Ph5 Hypercare</td><td>1,181,040</td><td>15%</td></tr>
        <tr style="font-weight:700;background:#C6D4E8;"><td>Labour Total</td><td>7,873,600</td><td>100%</td></tr>
        <tr><td>Contingency (15%)</td><td>1,181,040</td><td>&mdash;</td></tr>
        <tr><td>CAPEX (infra)</td><td>TBC at QG1</td><td>&mdash;</td></tr>
      </table>
      <div style="margin-top:10px;">
        <div style="font-size:10px;font-weight:600;margin-bottom:4px;">Resource Split (Labour)</div>
        <div style="display:flex;justify-content:space-between;font-size:10px;"><span>KPMG (PMO + Arch + SAP)</span><span>32.5%</span></div>
        <div style="display:flex;justify-content:space-between;font-size:10px;margin-top:2px;"><span>Infosys (all streams)</span><span>67.5%</span></div>
      </div>
    </div>
  </div>

</div><!-- /ROW 4 -->

<!-- FOOTER -->
<div class="footer">
  Trinity-CAM Management KPI Dashboard &nbsp;|&nbsp; {REPORT} &nbsp;|&nbsp;
  Sources: Trinity-CAM_Project_Schedule.xlsx, Trinity-CAM_Risk_Register.xlsx, Trinity-CAM_Cost_Plan.xlsx &nbsp;|&nbsp; CONFIDENTIAL &mdash; KPMG / Bosch Internal
</div>

</div><!-- /page -->
</body>
</html>"""

OUTPUT.write_text(HTML, encoding="utf-8")
print(f"[Trinity-CAM] KPI Dashboard: {OUTPUT}")
