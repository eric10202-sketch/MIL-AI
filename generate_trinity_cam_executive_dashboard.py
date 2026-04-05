#!/usr/bin/env python3
"""
Generate Trinity-CAM Executive Dashboard (HTML).
Follows executive-dashboard-generation SKILL.md — 3-page layout.
All content derived from Trinity-CAM schedule/cost/risk data.
"""

import base64
from datetime import date
from pathlib import Path

HERE   = Path(__file__).parent
LOGO   = HERE / "Bosch.png"
OUTPUT = HERE / "active-projects" / "Trinity-CAM" / "Trinity-CAM_Executive_Dashboard.html"
OUTPUT.parent.mkdir(parents=True, exist_ok=True)

logo_b64 = base64.b64encode(LOGO.read_bytes()).decode() if LOGO.exists() else ""
logo_tag = f'<img src="data:image/png;base64,{logo_b64}" alt="Bosch" style="height:36px;display:block;" />' if logo_b64 else ""

TODAY      = date.today()
KICKOFF    = date(2026, 7, 1)
GOLIVE     = date(2028, 1, 1)
COMPLETION = date(2028, 4, 1)
QG1        = date(2026, 10, 1)
QG23       = date(2027, 7, 31)
QG4        = date(2027, 12, 8)
REPORT     = TODAY.strftime("%d %B %Y")

def ddays(d): return (d - TODAY).days
def fmt_countdown(d):
    n = ddays(d)
    if n > 0:   return f"+{n} days"
    elif n < 0: return f"{abs(n)} days ago"
    else:       return "TODAY"

HTML = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Trinity-CAM Executive Dashboard</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:'Segoe UI',Arial,sans-serif;font-size:12px;color:#1a1a1a;background:#f4f6f9;}}

/* --- HEADER --- */
.hdr{{background:linear-gradient(135deg,#003b6e 0%,#005199 100%);color:#fff;padding:16px 28px;display:flex;align-items:center;gap:20px;}}
.bosch-logo{{display:flex;align-items:center;background:#fff;padding:4px 8px;border-radius:4px;}}
.hdr-center{{flex:1;}}
.hdr-center h1{{font-size:18px;font-weight:700;letter-spacing:0.4px;}}
.hdr-center h2{{font-size:11px;font-weight:400;opacity:.85;margin-top:3px;}}
.hdr-right{{text-align:right;font-size:11px;opacity:.85;white-space:nowrap;}}
.hdr-right strong{{font-size:14px;display:block;}}

/* --- COUNTDOWN STRIP --- */
.countdown-strip{{background:#003b6e;display:flex;gap:0;}}
.cd-box{{flex:1;border-right:1px solid rgba(255,255,255,.15);padding:8px 14px;color:#fff;}}
.cd-box:last-child{{border-right:none;}}
.cd-label{{font-size:9px;text-transform:uppercase;letter-spacing:.5px;opacity:.75;}}
.cd-n{{font-size:22px;font-weight:700;}}
.cd-sub{{font-size:9px;opacity:.7;}}

/* --- PAGE SECTIONS --- */
.page{{max-width:1140px;margin:0 auto;padding:16px 20px 8px;}}
.page-break{{page-break-before:always;border-top:3px solid #003b6e;margin:24px 0 0;}}

/* --- SECTION CARD --- */
.card{{background:#fff;border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:14px;overflow:hidden;}}
.card-title{{background:#005199;color:#fff;font-size:11px;font-weight:700;padding:7px 14px;letter-spacing:.4px;text-transform:uppercase;}}
.card-body{{padding:12px 14px;}}

/* --- OVERVIEW GRID --- */
.overview-grid{{display:grid;grid-template-columns:1fr 340px;gap:12px;}}
.overview-text p{{font-size:12px;line-height:1.55;margin:0 0 8px;}}
.fact-box{{background:#EFF4FB;border-radius:4px;padding:10px;font-size:11px;}}
.fact-box .fact-label{{font-size:9px;text-transform:uppercase;letter-spacing:.4px;color:#555;margin-bottom:2px;}}
.fact-box .fact-val{{font-weight:700;font-size:13px;color:#003b6e;}}
.fact-sep{{border-top:1px solid #ccd8ed;margin:6px 0;}}

/* --- STAT STRIP --- */
.stat-strip{{display:grid;grid-template-columns:repeat(6,1fr);gap:8px;}}
.stat-tile{{background:#fff;border-radius:5px;box-shadow:0 1px 3px rgba(0,0,0,.07);padding:10px 8px;text-align:center;}}
.stat-tile .si{{font-size:22px;}}
.stat-tile .sn{{font-size:16px;font-weight:700;color:#003b6e;}}
.stat-tile .sl{{font-size:9px;color:#666;text-transform:uppercase;letter-spacing:.4px;margin-top:2px;}}

/* --- PHASE TIMELINE --- */
.phase-bar{{display:flex;height:36px;border-radius:4px;overflow:hidden;}}
.pb-seg{{display:flex;align-items:center;justify-content:center;color:#fff;font-size:9px;font-weight:700;text-align:center;white-space:nowrap;overflow:hidden;padding:2px;}}
.phase-labels{{display:flex;font-size:9px;color:#555;margin-top:3px;}}
.phase-labels span{{flex:1;white-space:nowrap;}}

/* --- TWO-COL LAYOUT --- */
.two-col{{display:grid;grid-template-columns:1fr 1fr;gap:12px;}}
.three-col{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;}}
.four-col{{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;}}

/* --- MILESTONES TABLE --- */
.ms-table{{width:100%;border-collapse:collapse;font-size:11px;}}
.ms-table th{{background:#003b6e;color:#fff;padding:6px 8px;text-align:left;font-size:10px;}}
.ms-table td{{padding:6px 8px;border-bottom:1px solid #e8ecf2;vertical-align:middle;}}
.ms-table tr:nth-child(even){{background:#EFF4FB;}}
.pill{{display:inline-block;padding:2px 7px;border-radius:10px;font-size:9px;font-weight:700;}}
.pill-future{{background:#EFF4FB;color:#005199;}}
.pill-active{{background:#FFF2CC;color:#c07000;}}
.pill-done{{background:#d5f5e3;color:#1e8449;}}
.pill-red{{background:#fde;color:#c00;}}

/* --- BUDGET DONUT (CSS only) --- */
.budget-summary{{text-align:center;padding:6px 0;}}
.budget-total{{font-size:22px;font-weight:700;color:#003b6e;}}
.budget-sub{{font-size:10px;color:#666;margin-bottom:8px;}}
.budget-bar-wrap{{margin:6px 0;}}
.budget-bar{{height:16px;border-radius:8px;overflow:hidden;display:flex;}}
.bb-seg{{height:100%;}}
.budget-legend{{display:flex;flex-wrap:wrap;gap:6px;justify-content:center;margin-top:6px;font-size:9px;}}
.bl-dot{{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:3px;}}

/* --- WORKSTREAM GRID --- */
.ws-card{{background:#EFF4FB;border-radius:4px;padding:8px 10px;border-left:4px solid #0066CC;}}
.ws-title{{font-weight:700;font-size:11px;color:#003b6e;margin-bottom:4px;}}
.ws-bullets{{font-size:10px;color:#333;line-height:1.5;}}
.ws-bullets li{{margin:2px 0;}}
.ws-conf{{display:inline-block;margin-top:6px;padding:2px 7px;border-radius:10px;font-size:9px;font-weight:700;}}
.conf-g{{background:#27ae60;color:#fff;}}
.conf-a{{background:#f39c12;color:#fff;}}
.conf-r{{background:#e74c3c;color:#fff;}}

/* --- QG TRACKER --- */
.qg-row{{display:flex;gap:10px;padding:7px 0;border-bottom:1px solid #e8ecf2;align-items:flex-start;}}
.qg-row:last-child{{border-bottom:none;}}
.qg-date{{min-width:90px;font-size:10px;font-weight:600;color:#003b6e;}}
.qg-name{{font-weight:700;font-size:11px;}}
.qg-sub{{font-size:10px;color:#666;margin-top:2px;}}
.qg-cd{{min-width:70px;text-align:right;font-size:10px;font-weight:600;color:#0066CC;}}

/* --- RISK GRID --- */
.risk-card{{border-radius:4px;padding:8px 10px;}}
.risk-high{{background:#fdecea;border-left:4px solid #e74c3c;}}
.risk-med{{background:#fff9ec;border-left:4px solid #f39c12;}}
.risk-low{{background:#eafbf1;border-left:4px solid #27ae60;}}
.risk-id{{font-weight:700;font-size:10px;}}
.risk-title{{font-size:11px;font-weight:600;margin:2px 0;}}
.risk-score{{font-size:9px;color:#666;}}

/* --- WAVE BARS --- */
.wave-row{{display:flex;align-items:center;gap:8px;margin:5px 0;}}
.wave-label{{min-width:110px;font-size:10px;font-weight:600;}}
.wave-bar{{height:16px;border-radius:3px;background:#0066CC;display:flex;align-items:center;padding-left:6px;color:#fff;font-size:9px;font-weight:700;}}
.wave-n{{min-width:60px;font-size:10px;color:#555;}}

/* --- HOTSPOT GRID --- */
.hotspot-card{{background:#EFF4FB;border-radius:4px;padding:8px 10px;}}
.hs-country{{font-weight:700;font-size:12px;margin-bottom:4px;color:#003b6e;}}
.hs-items{{font-size:10px;color:#333;line-height:1.5;}}

/* --- CRITICAL PATH --- */
.cp-cell{{background:#EFF4FB;border-radius:4px;padding:8px 10px;}}
.cp-title{{font-weight:700;font-size:11px;color:#003b6e;margin-bottom:5px;border-bottom:2px solid #0066CC;padding-bottom:3px;}}
.cp-item{{font-size:10px;line-height:1.5;padding:2px 0;border-bottom:1px dotted #c8d4e8;}}
.cp-item:last-child{{border-bottom:none;}}

/* --- REGION BARS --- */
.region-row{{display:flex;align-items:center;gap:8px;margin:4px 0;font-size:11px;}}
.region-name{{min-width:80px;color:#333;}}
.region-bar{{height:14px;background:#0066CC;border-radius:2px;}}
.region-n{{font-size:10px;color:#555;min-width:40px;}}

/* --- FOOTER --- */
.footer{{text-align:center;font-size:9px;color:#aaa;padding:10px 0 20px;}}

/* --- STATS DARK STRIP --- */
.stats-dark{{background:#003b6e;display:grid;grid-template-columns:repeat(4,1fr);gap:0;border-radius:5px;overflow:hidden;margin-bottom:12px;}}
.sd-cell{{padding:10px 12px;border-right:1px solid rgba(255,255,255,.12);color:#fff;text-align:center;}}
.sd-cell:last-child{{border-right:none;}}
.sd-n{{font-size:22px;font-weight:700;}}
.sd-l{{font-size:9px;opacity:.75;text-transform:uppercase;letter-spacing:.4px;margin-top:2px;}}
</style>
</head>
<body>

<!-- ========== HEADER ========== -->
<div class="hdr">
  <div class="bosch-logo">{logo_tag}</div>
  <div class="hdr-center">
    <h1>Trinity-CAM &nbsp;&mdash;&nbsp; IT Carve-out Executive Dashboard</h1>
    <h2>Johnson Controls International (JCI) &rarr; Robert Bosch GmbH &nbsp;|&nbsp; Aircondition Business &nbsp;|&nbsp; Integration Model</h2>
  </div>
  <div class="hdr-right">
    <strong>{REPORT}</strong>
    GoLive: 1 Jan 2028 &nbsp;|&nbsp; {fmt_countdown(GOLIVE)}
  </div>
</div>

<!-- COUNTDOWN STRIP -->
<div class="countdown-strip">
  <div class="cd-box"><div class="cd-label">Days to QG1</div><div class="cd-n">{ddays(QG1)}</div><div class="cd-sub">1 Oct 2026 — Concept</div></div>
  <div class="cd-box"><div class="cd-label">Days to QG2&amp;3</div><div class="cd-n">{ddays(QG23)}</div><div class="cd-sub">31 Jul 2027 — MZ Ready</div></div>
  <div class="cd-box"><div class="cd-label">Days to QG4</div><div class="cd-n">{ddays(QG4)}</div><div class="cd-sub">8 Dec 2027 — Final Ready</div></div>
  <div class="cd-box"><div class="cd-label">Days to GoLive</div><div class="cd-n">{ddays(GOLIVE)}</div><div class="cd-sub">1 Jan 2028 — Day 1</div></div>
  <div class="cd-box"><div class="cd-label">Days to Completion</div><div class="cd-n">{ddays(COMPLETION)}</div><div class="cd-sub">1 Apr 2028 — QG5</div></div>
  <div class="cd-box"><div class="cd-label">Programme Duration</div><div class="cd-n">21</div><div class="cd-sub">months (Jul 2026 – Apr 2028)</div></div>
</div>

<!-- ========== PAGE 1 ========== -->
<div class="page">

<!-- --- PROJECT OVERVIEW --- -->
<div class="card">
  <div class="card-title">Project Overview</div>
  <div class="card-body overview-grid">
    <div class="overview-text">
      <p>Project Trinity-CAM is the strategic IT separation and integration programme supporting Bosch&rsquo;s acquisition of Johnson Controls International&rsquo;s (JCI) global Aircondition division. All 12,000 Aircon IT users, 1,800+ applications and the full SAP landscape must be separated from JCI and migrated through an Infosys-operated Merger Zone into the Bosch IT environment.</p>
      <p>The programme operates under the <strong>Integration model</strong>: a transient Merger Zone (MZ) is built and operated by Infosys as a landing zone for all JCI assets before final Bosch integration. JCI TSA services expired 30 June 2026; carve-out commenced 1 July 2026. Infosys is sole MZ delivery partner covering infrastructure build, SAP migration, application migration, and 24/7 hypercare support.</p>
      <p>The programme spans 21 months across 5 phases, culminating in GoLive 1 January 2028 and formal programme closure at QG5 on 1 April 2028 following 90-day hypercare.</p>
    </div>
    <div>
      <div class="fact-box">
        <div class="fact-label">Carve-out Model</div>
        <div class="fact-val">Integration</div>
        <div style="font-size:10px;color:#444;margin-top:3px;">JCI IT &rarr; Merger Zone (Infosys) &rarr; Bosch IT</div>
        <div class="fact-sep"></div>
        <div class="fact-label">Buyer (Sponsor Customer)</div>
        <div class="fact-val">Robert Bosch GmbH</div>
        <div class="fact-sep"></div>
        <div class="fact-label">Seller (Sponsor Contractor)</div>
        <div class="fact-val">Johnson Controls International</div>
        <div class="fact-sep"></div>
        <div class="fact-label">PMO</div>
        <div class="fact-val">KPMG</div>
        <div class="fact-label" style="margin-top:4px;">IT Delivery Partner</div>
        <div class="fact-val">Infosys</div>
        <div class="fact-sep"></div>
        <div class="fact-label">Programme Budget (Labour)</div>
        <div class="fact-val">EUR 7,873,600</div>
        <div style="font-size:9px;color:#888;">+ EUR 1,181,040 contingency (15%) | CAPEX TBC at QG1</div>
      </div>
    </div>
  </div>
</div>

<!-- --- STATS ROW --- -->
<div class="stat-strip" style="margin-bottom:14px;">
  <div class="stat-tile"><div class="si">&#127758;</div><div class="sn">48</div><div class="sl">Global Sites</div></div>
  <div class="stat-tile"><div class="si">&#128100;</div><div class="sn">12,000</div><div class="sl">IT Users</div></div>
  <div class="stat-tile"><div class="si">&#128187;</div><div class="sn">12,000</div><div class="sl">Client Devices</div></div>
  <div class="stat-tile"><div class="si">&#9881;</div><div class="sn">1,800+</div><div class="sl">Applications</div></div>
  <div class="stat-tile"><div class="si">&#128197;</div><div class="sn">21 mo</div><div class="sl">Programme Duration</div></div>
  <div class="stat-tile"><div class="si">&#127975;</div><div class="sn">Expired</div><div class="sl">JCI TSA (30 Jun 2026)</div></div>
</div>

<!-- --- PHASE TIMELINE --- -->
<div class="card">
  <div class="card-title">Programme Phase Timeline</div>
  <div class="card-body">
    <div class="phase-bar">
      <div class="pb-seg" style="flex:13;background:#357ab7;">Ph1<br>Initiation</div>
      <div class="pb-seg" style="flex:30;background:#0066CC;">Ph2<br>MZ Build</div>
      <div class="pb-seg" style="flex:20;background:#005199;">Ph3<br>Testing &amp; Migration</div>
      <div class="pb-seg" style="flex:7;background:#003b6e;">Ph4<br>Final Readiness</div>
      <div class="pb-seg" style="flex:3;background:#1a1a6e;">GoLive</div>
      <div class="pb-seg" style="flex:12;background:#444;">Ph5 Hypercare</div>
    </div>
    <div class="phase-labels">
      <span>QG0 01 Jul 2026</span>
      <span>QG1 01 Oct 2026</span>
      <span>QG2&amp;3 31 Jul 2027</span>
      <span>QG4 08 Dec 2027</span>
      <span>GoLive 01 Jan 2028</span>
      <span style="text-align:right">QG5 01 Apr 2028</span>
    </div>
  </div>
</div>

<!-- --- TWO-COL: MILESTONES + BUDGET --- -->
<div class="two-col">

  <div class="card">
    <div class="card-title">Key Milestones &amp; Quality Gates</div>
    <div class="card-body" style="padding:8px;">
      <table class="ms-table">
        <tr><th>Gate</th><th>Description</th><th>Date</th><th>Countdown</th><th>Status</th></tr>
        <tr>
          <td><strong>QG0</strong></td>
          <td>Programme kickoff; PMO + Infosys mobilised</td>
          <td>01 Jul 2026</td>
          <td style="font-size:10px;color:#0066CC;">{fmt_countdown(KICKOFF)}</td>
          <td><span class="pill pill-future">PLANNED</span></td>
        </tr>
        <tr>
          <td><strong>QG1</strong></td>
          <td>Concept approved; MZ architecture signed off; app inventory complete</td>
          <td>01 Oct 2026</td>
          <td style="font-size:10px;color:#0066CC;">{fmt_countdown(QG1)}</td>
          <td><span class="pill pill-future">PLANNED</span></td>
        </tr>
        <tr>
          <td><strong>QG2&amp;3</strong></td>
          <td>Merger Zone fully built; SAP interfaces rewired; Wave 1 apps validated</td>
          <td>31 Jul 2027</td>
          <td style="font-size:10px;color:#0066CC;">{fmt_countdown(QG23)}</td>
          <td><span class="pill pill-future">PLANNED</span></td>
        </tr>
        <tr>
          <td><strong>QG4</strong></td>
          <td>All 12,000 users &amp; 1,800+ apps migrated; SAP Mock 2 complete; UAT passed</td>
          <td>08 Dec 2027</td>
          <td style="font-size:10px;color:#0066CC;">{fmt_countdown(QG4)}</td>
          <td><span class="pill pill-future">PLANNED</span></td>
        </tr>
        <tr>
          <td><strong>GoLive</strong></td>
          <td>Day 1 MZ cutover; all systems live; hypercare active</td>
          <td>01 Jan 2028</td>
          <td style="font-size:10px;color:#0066CC;">{fmt_countdown(GOLIVE)}</td>
          <td><span class="pill pill-active">TARGET</span></td>
        </tr>
        <tr>
          <td><strong>QG5</strong></td>
          <td>90-day hypercare complete; TSA exit confirmed; Bosch IT handover</td>
          <td>01 Apr 2028</td>
          <td style="font-size:10px;color:#0066CC;">{fmt_countdown(COMPLETION)}</td>
          <td><span class="pill pill-future">PLANNED</span></td>
        </tr>
      </table>
    </div>
  </div>

  <div class="card">
    <div class="card-title">Budget Distribution</div>
    <div class="card-body">
      <div class="budget-summary">
        <div class="budget-total">EUR 7,873,600</div>
        <div class="budget-sub">Total Programme Labour (KPMG + Infosys) | CAPEX TBC at QG1</div>
      </div>
      <div class="budget-bar-wrap">
        <div class="budget-bar">
          <div class="bb-seg" style="flex:17.5;background:#003b6e;" title="KPMG PMO 17.5%"></div>
          <div class="bb-seg" style="flex:9.5;background:#005199;" title="KPMG Arch 9.5%"></div>
          <div class="bb-seg" style="flex:5.5;background:#0066CC;" title="KPMG SAP 5.5%"></div>
          <div class="bb-seg" style="flex:10.5;background:#357ab7;" title="Infosys PM 10.5%"></div>
          <div class="bb-seg" style="flex:20.9;background:#5b9bd5;" title="Infosys Infra 20.9%"></div>
          <div class="bb-seg" style="flex:4.0;background:#7db4de;" title="Infosys IAM 4.0%"></div>
          <div class="bb-seg" style="flex:14.5;background:#9ccce8;" title="Infosys SAP 14.5%"></div>
          <div class="bb-seg" style="flex:6.6;background:#b8dff2;" title="Infosys Apps 6.6%"></div>
          <div class="bb-seg" style="flex:5.1;background:#d0edf8;" title="Infosys Data 5.1%"></div>
          <div class="bb-seg" style="flex:5.9;background:#e4f5fb;" title="Infosys Service 5.9%"></div>
        </div>
      </div>
      <div class="budget-legend">
        <div><span class="bl-dot" style="background:#003b6e;"></span>KPMG PMO 17.5%</div>
        <div><span class="bl-dot" style="background:#005199;"></span>KPMG Arch 9.5%</div>
        <div><span class="bl-dot" style="background:#0066CC;"></span>KPMG SAP 5.5%</div>
        <div><span class="bl-dot" style="background:#357ab7;"></span>Infosys PM 10.5%</div>
        <div><span class="bl-dot" style="background:#5b9bd5;"></span>Infosys Infra 20.9%</div>
        <div><span class="bl-dot" style="background:#7db4de;"></span>Infosys IAM 4.0%</div>
        <div><span class="bl-dot" style="background:#9ccce8;"></span>Infosys SAP 14.5%</div>
        <div><span class="bl-dot" style="background:#b8dff2;"></span>Infosys Apps 6.6%</div>
        <div><span class="bl-dot" style="background:#d0edf8;"></span>Infosys Data 5.1%</div>
        <div><span class="bl-dot" style="background:#e4f5fb;background:#6aadcc;"></span>Infosys Service 5.9%</div>
      </div>
      <table style="width:100%;border-collapse:collapse;font-size:10px;margin-top:10px;">
        <tr><th style="background:#003b6e;color:#fff;padding:4px 6px;text-align:left;">Phase</th><th style="background:#003b6e;color:#fff;padding:4px 6px;text-align:right;">EUR</th><th style="background:#003b6e;color:#fff;padding:4px 6px;text-align:right;">%</th></tr>
        <tr><td style="padding:3px 6px;border-bottom:1px solid #e8ecf2;">Ph1 Initiation (Jul–Oct 2026)</td><td style="padding:3px 6px;text-align:right;border-bottom:1px solid #e8ecf2;">629,888</td><td style="padding:3px 6px;text-align:right;border-bottom:1px solid #e8ecf2;">8%</td></tr>
        <tr style="background:#EFF4FB;"><td style="padding:3px 6px;border-bottom:1px solid #e8ecf2;">Ph2 MZ Build (Oct 2026–Jul 2027)</td><td style="padding:3px 6px;text-align:right;border-bottom:1px solid #e8ecf2;">2,755,760</td><td style="padding:3px 6px;text-align:right;border-bottom:1px solid #e8ecf2;">35%</td></tr>
        <tr><td style="padding:3px 6px;border-bottom:1px solid #e8ecf2;">Ph3 Testing &amp; Migration (Jul–Dec 2027)</td><td style="padding:3px 6px;text-align:right;border-bottom:1px solid #e8ecf2;">2,362,080</td><td style="padding:3px 6px;text-align:right;border-bottom:1px solid #e8ecf2;">30%</td></tr>
        <tr style="background:#EFF4FB;"><td style="padding:3px 6px;border-bottom:1px solid #e8ecf2;">Ph4 Final Readiness (Dec 2027)</td><td style="padding:3px 6px;text-align:right;border-bottom:1px solid #e8ecf2;">708,624</td><td style="padding:3px 6px;text-align:right;border-bottom:1px solid #e8ecf2;">9%</td></tr>
        <tr><td style="padding:3px 6px;border-bottom:1px solid #e8ecf2;">GoLive &amp; Cutover</td><td style="padding:3px 6px;text-align:right;border-bottom:1px solid #e8ecf2;">236,208</td><td style="padding:3px 6px;text-align:right;border-bottom:1px solid #e8ecf2;">3%</td></tr>
        <tr style="background:#EFF4FB;"><td style="padding:3px 6px;">Ph5 Hypercare (Jan–Apr 2028)</td><td style="padding:3px 6px;text-align:right;">1,181,040</td><td style="padding:3px 6px;text-align:right;">15%</td></tr>
        <tr style="font-weight:700;background:#C6D4E8;"><td style="padding:4px 6px;">Total Labour</td><td style="padding:4px 6px;text-align:right;">7,873,600</td><td style="padding:4px 6px;text-align:right;">100%</td></tr>
      </table>
    </div>
  </div>

</div><!-- /two-col -->

<!-- ========== PAGE 2 ========== -->
<div class="page-break"></div>

<!-- --- WORKSTREAM COVERAGE --- -->
<div class="card">
  <div class="card-title">IT Workstream Coverage — Confidence Overview</div>
  <div class="card-body">
    <div class="three-col">
      <div class="ws-card">
        <div class="ws-title">WS1 — Merger Zone Infrastructure</div>
        <ul class="ws-bullets">
          <li>DC/cloud build (Infosys); 24-site multi-DC</li>
          <li>SD-WAN to 48 sites; MPLS failover</li>
          <li>DR site; backup orchestration</li>
        </ul>
        <span class="ws-conf conf-g">HIGH CONFIDENCE</span>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS2 — SAP Migration</div>
        <ul class="ws-bullets">
          <li>Full system copy JCI &rarr; MZ</li>
          <li>Client separation; interface rewiring</li>
          <li>2 mock cutovers; SAP security redesign</li>
        </ul>
        <span class="ws-conf conf-a">MEDIUM — R001 SAP complexity at-risk</span>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS3 — Application Migration (1,800+)</div>
        <ul class="ws-bullets">
          <li>3 waves; Wave 1 ~400 apps Jul-Sep 2027</li>
          <li>Wave 2 ~800 apps; Wave 3 ~600 apps</li>
          <li>Compatibility testing; remediations</li>
        </ul>
        <span class="ws-conf conf-a">MEDIUM — Wave capacity risk R007</span>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS4 — End-User Workplace</div>
        <ul class="ws-bullets">
          <li>12,000 users; M365 tenant migration</li>
          <li>Intune MDM onboarding; VOIP cut</li>
          <li>Site-by-site wave plan; 6 site clusters</li>
        </ul>
        <span class="ws-conf conf-g">HIGH CONFIDENCE</span>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS5 — Identity &amp; Access Management</div>
        <ul class="ws-bullets">
          <li>MZ Active Directory forest build</li>
          <li>Identity federation JCI &harr; MZ; PAM</li>
          <li>MFA rollout for all 12,000 users</li>
        </ul>
        <span class="ws-conf conf-g">HIGH CONFIDENCE</span>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS6 — Security &amp; Compliance</div>
        <ul class="ws-bullets">
          <li>GDPR data classification 48 jurisdictions</li>
          <li>SOC monitoring; vulnerability management</li>
          <li>Pen test pre-GoLive; ISMS baseline</li>
        </ul>
        <span class="ws-conf conf-a">MEDIUM — R006 GDPR breach risk</span>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS7 — Data Migration</div>
        <ul class="ws-bullets">
          <li>ETL framework; 48-country data mapping</li>
          <li>GDPR legal basis per data category</li>
          <li>Reconciliation reports at each wave</li>
        </ul>
        <span class="ws-conf conf-g">HIGH CONFIDENCE</span>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS8 — TSA Exit &amp; HR/Legal</div>
        <ul class="ws-bullets">
          <li>TSA service dependency catalogue</li>
          <li>TUPE compliance Germany, France, UK</li>
          <li>Formal exit confirmation by QG5</li>
        </ul>
        <span class="ws-conf conf-a">MEDIUM — R005 TUPE risk</span>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS9 — Programme Control &amp; Hypercare</div>
        <ul class="ws-bullets">
          <li>KPMG PMO; weekly steering; risk cadence</li>
          <li>90-day hypercare L3 support (Infosys)</li>
          <li>Knowledge transfer; Bosch IT handover</li>
        </ul>
        <span class="ws-conf conf-g">HIGH CONFIDENCE</span>
      </div>
    </div>
  </div>
</div>

<!-- --- QG TRACKER --- -->
<div class="card">
  <div class="card-title">Quality Gate Tracker</div>
  <div class="card-body">
    <div class="qg-row">
      <div class="qg-date">01 Jul 2026</div>
      <div style="flex:1;"><div class="qg-name">QG0 — Programme Kickoff</div><div class="qg-sub">KPMG PMO mobilised; Infosys SOW signed; all workstream leads appointed; Steering Committee constituted; programme plan baselined</div></div>
      <div class="qg-cd">{fmt_countdown(KICKOFF)}</div>
      <div style="min-width:70px;text-align:right;"><span class="pill pill-future">PLANNED</span></div>
    </div>
    <div class="qg-row">
      <div class="qg-date">01 Oct 2026</div>
      <div style="flex:1;"><div class="qg-name">QG1 — Concept Approved</div><div class="qg-sub">Application inventory complete (1,800+ catalogued); MZ architecture &amp; design approved; TSA catalogue agreed; wave plan baselined; CAPEX budget approved; risk register v1 signed off</div></div>
      <div class="qg-cd">{fmt_countdown(QG1)}</div>
      <div style="min-width:70px;text-align:right;"><span class="pill pill-future">PLANNED</span></div>
    </div>
    <div class="qg-row">
      <div class="qg-date">31 Jul 2027</div>
      <div style="flex:1;"><div class="qg-name">QG2&amp;3 — MZ Ready / SAP Build Complete</div><div class="qg-sub">Merger Zone DC/cloud fully provisioned; all 48-site connectivity live; SAP system copy complete; interfaces rewired; Wave 1 (~400 apps) validated; IAM/AD fully functional</div></div>
      <div class="qg-cd">{fmt_countdown(QG23)}</div>
      <div style="min-width:70px;text-align:right;"><span class="pill pill-future">PLANNED</span></div>
    </div>
    <div class="qg-row">
      <div class="qg-date">08 Dec 2027</div>
      <div style="flex:1;"><div class="qg-name">QG4 — GoLive Readiness Confirmed</div><div class="qg-sub">All 12,000 users migrated to MZ; all 1,800+ applications certified; SAP Mock Cutover 2 passed with zero P1/P2 defects; DR tested; performance validated; Steering sign-off granted</div></div>
      <div class="qg-cd">{fmt_countdown(QG4)}</div>
      <div style="min-width:70px;text-align:right;"><span class="pill pill-future">PLANNED</span></div>
    </div>
    <div class="qg-row">
      <div class="qg-date">01 Jan 2028</div>
      <div style="flex:1;"><div class="qg-name">GoLive — Merger Zone Day 1</div><div class="qg-sub">Final readiness confirmed; executive cut decision approved; Merger Zone live as operational IT environment; 24/7 hypercare active; JCI access decommissioned</div></div>
      <div class="qg-cd">{fmt_countdown(GOLIVE)}</div>
      <div style="min-width:70px;text-align:right;"><span class="pill pill-active">TARGET</span></div>
    </div>
    <div class="qg-row">
      <div class="qg-date">01 Apr 2028</div>
      <div style="flex:1;"><div class="qg-name">QG5 — Programme Closure</div><div class="qg-sub">90-day hypercare complete; JCI TSA formally exited; Bosch IT operations fully handed over; lessons learned documented; programme team stood down</div></div>
      <div class="qg-cd">{fmt_countdown(COMPLETION)}</div>
      <div style="min-width:70px;text-align:right;"><span class="pill pill-future">PLANNED</span></div>
    </div>
  </div>
</div>

<!-- --- TWO COL: REGIONAL SCOPE + RISK INDICATORS --- -->
<div class="two-col">

  <div class="card">
    <div class="card-title">Regional Site Distribution</div>
    <div class="card-body">
      <div class="region-row"><span class="region-name">EMEA</span><div class="region-bar" style="width:220px;"></div><span class="region-n">22 sites</span></div>
      <div class="region-row"><span class="region-name">APAC</span><div class="region-bar" style="width:140px;"></div><span class="region-n">14 sites</span></div>
      <div class="region-row"><span class="region-name">Americas</span><div class="region-bar" style="width:120px;"></div><span class="region-n">12 sites</span></div>
      <div style="margin-top:12px;font-size:10px;color:#555;">
        <strong>EMEA Hotspots:</strong> Germany (6 sites, TUPE), France (3 sites), Netherlands (2), UK (2), Poland, Hungary, Czechia, Turkey, Italy<br>
        <strong>APAC Hotspots:</strong> China (5), Japan (3), India (3), South Korea (2), Singapore<br>
        <strong>Americas:</strong> USA (8), Mexico (2), Brazil (2)
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-title">Key Risk Indicators (from Risk Register)</div>
    <div class="card-body" style="display:flex;flex-direction:column;gap:6px;">
      <div class="risk-card risk-high">
        <div class="risk-id">R001 — HIGH</div>
        <div class="risk-title">SAP Landscape Complexity</div>
        <div class="risk-score">P 70% &times; I Very High &nbsp;|&nbsp; Score: <strong>20</strong> &nbsp;|&nbsp; QG2&amp;3 at risk</div>
      </div>
      <div class="risk-card risk-high">
        <div class="risk-id">R003 — HIGH</div>
        <div class="risk-title">SAP Mock Cutover 2 Timing</div>
        <div class="risk-score">P 50% &times; I Very High &nbsp;|&nbsp; Score: <strong>15</strong> &nbsp;|&nbsp; GoLive at risk</div>
      </div>
      <div class="risk-card risk-med">
        <div class="risk-id">R002 — MEDIUM</div>
        <div class="risk-title">Infosys MZ Delivery Delay</div>
        <div class="risk-score">P 50% &times; I High &nbsp;|&nbsp; Score: <strong>12</strong> &nbsp;|&nbsp; Phase 3 start risk</div>
      </div>
      <div class="risk-card risk-med">
        <div class="risk-id">R007 — MEDIUM</div>
        <div class="risk-title">App Compatibility (1,800+)</div>
        <div class="risk-score">P 50% &times; I High &nbsp;|&nbsp; Score: <strong>12</strong> &nbsp;|&nbsp; Wave capacity risk</div>
      </div>
      <div class="risk-card risk-med">
        <div class="risk-id">R006 — MEDIUM</div>
        <div class="risk-title">GDPR / Data Breach Exposure</div>
        <div class="risk-score">P 30% &times; I Very High &nbsp;|&nbsp; Score: <strong>10</strong> &nbsp;|&nbsp; EUR 20M fine exposure</div>
      </div>
      <div style="font-size:10px;color:#666;margin-top:4px;">Full risk register: <em>Trinity-CAM_Risk_Register.xlsx</em> — 25 risks (24 threats, 1 opportunity)</div>
    </div>
  </div>

</div>

<!-- ========== PAGE 3 ========== -->
<div class="page-break"></div>

<!-- --- APPLICATION MIGRATION WAVES --- -->
<div class="card">
  <div class="card-title">Application Migration Waves (1,800+ Applications)</div>
  <div class="card-body">
    <div style="margin-bottom:8px;font-size:11px;color:#555;">Infosys delivers 3 migration waves. SAP migrates via dedicated parallel track (system copy + interface rewiring).</div>
    <div class="wave-row">
      <div class="wave-label">SAP Track (all systems)</div>
      <div class="wave-bar" style="flex:0 0 380px;background:#003b6e;">Apr 2027 &rarr; Nov 2027</div>
      <div class="wave-n">~50 SAP systems</div>
    </div>
    <div class="wave-row">
      <div class="wave-label">Wave 1 — Pilot Apps</div>
      <div class="wave-bar" style="flex:0 0 180px;background:#0066CC;">Jul &rarr; Sep 2027</div>
      <div class="wave-n">~400 apps (22%)</div>
    </div>
    <div class="wave-row">
      <div class="wave-label">Wave 2 — Core Business Apps</div>
      <div class="wave-bar" style="flex:0 0 260px;background:#357ab7;">Sep &rarr; Nov 2027</div>
      <div class="wave-n">~800 apps (44%)</div>
    </div>
    <div class="wave-row">
      <div class="wave-label">Wave 3 — Remaining + Long Tail</div>
      <div class="wave-bar" style="flex:0 0 200px;background:#5b9bd5;">Oct &rarr; Nov 2027</div>
      <div class="wave-n">~550 apps (31%)</div>
    </div>
    <div style="margin-top:10px;font-size:10px;color:#555;">Apps decommissioned at source: post-GoLive JCI access revoked. Legacy apps not viable for MZ: managed via TSA extension clause (risk R011).</div>
  </div>
</div>

<!-- --- COUNTRY COMPLEXITY HOTSPOTS --- -->
<div class="card">
  <div class="card-title">Country-Specific Complexity Hotspots</div>
  <div class="card-body">
    <div class="three-col">
      <div class="hotspot-card">
        <div class="hs-country">&#127465;&#127466; Germany (6 sites)</div>
        <div class="hs-items">
          <div>TUPE / works council approvals required before migration (risk R005)</div>
          <div>Betriebsverfassungsgesetz co-determination obligations</div>
          <div>IT works agreements for monitoring all migrated systems</div>
          <div>Frankfurt hub: SAP primary instance location</div>
        </div>
      </div>
      <div class="hotspot-card">
        <div class="hs-country">&#127464;&#127475; China (5 sites)</div>
        <div class="hs-items">
          <div>PIPL/DSL data residency: personal data must stay onshore (risk R013)</div>
          <div>MLPS 2.0 compliance required for MZ instances</div>
          <div>Great Firewall: VPN and connectivity constraints</div>
          <div>Data export approval via CAC mandatory before migration</div>
        </div>
      </div>
      <div class="hotspot-card">
        <div class="hs-country">&#127467;&#127479; France (3 sites)</div>
        <div class="hs-items">
          <div>RGPD + CNIL notifications required for all data migrations</div>
          <div>Social consultation process for IT service changes</div>
          <div>Paris site: secondary SAP instance; EWC notification needed</div>
        </div>
      </div>
      <div class="hotspot-card">
        <div class="hs-country">&#127482;&#127480; USA (8 sites)</div>
        <div class="hs-items">
          <div>Largest site cluster; Wave 2 primary migration target</div>
          <div>CCPA/state privacy laws (CA, VA, CO) data handling requirements</div>
          <div>Infosys near-shore team for Wave 2 execution support</div>
        </div>
      </div>
      <div class="hotspot-card">
        <div class="hs-country">&#127471;&#127477; Japan (3 sites)</div>
        <div class="hs-items">
          <div>APPI data privacy; cross-border transfer restrictions to MZ</div>
          <div>Osaka DC: local MZ node required; Infosys local ISP dependency</div>
          <div>On-site migration window restricted to Golden Week avoidance</div>
        </div>
      </div>
      <div class="hotspot-card">
        <div class="hs-country">&#127470;&#127475; India (3 sites)</div>
        <div class="hs-items">
          <div>DPDPA (2023) compliance; Infosys domestic team co-location advantage</div>
          <div>Mumbai, Bangalore sites: Wave 1 pilot candidates</div>
          <div>Infosys offshore capacity: SAP QA / test resources for all waves</div>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- --- DARK STATS STRIP --- -->
<div class="stats-dark">
  <div class="sd-cell"><div class="sd-n">118</div><div class="sd-l">Schedule Tasks</div></div>
  <div class="sd-cell"><div class="sd-n">10</div><div class="sd-l">Resource Groups</div></div>
  <div class="sd-cell"><div class="sd-n">EUR 7.9M</div><div class="sd-l">Labour Budget</div></div>
  <div class="sd-cell"><div class="sd-n">3</div><div class="sd-l">Migration Waves</div></div>
</div>

<!-- --- CRITICAL PATH & PRINCIPLES --- -->
<div class="card">
  <div class="card-title">Critical Path &amp; Guiding Principles</div>
  <div class="card-body">
    <div class="four-col">
      <div class="cp-cell">
        <div class="cp-title">Infrastructure Critical Path</div>
        <div class="cp-item">MZ DC procurement complete by Dec 2026</div>
        <div class="cp-item">MZ DC build complete by Apr 2027</div>
        <div class="cp-item">48-site SD-WAN live by Jun 2027</div>
        <div class="cp-item">DR site validated by Oct 2027</div>
        <div class="cp-item">Production DR test passed before QG4</div>
      </div>
      <div class="cp-cell">
        <div class="cp-title">SAP Critical Path</div>
        <div class="cp-item">SAP system copy start Apr 2027</div>
        <div class="cp-item">SAP system copy complete Jun 2027</div>
        <div class="cp-item">SAP Mock Cutover 1: Jul 2027</div>
        <div class="cp-item">Interface rewiring complete Sep 2027</div>
        <div class="cp-item">SAP Mock Cutover 2: Nov 2027</div>
        <div class="cp-item">SAP UAT passed by Dec 2027</div>
      </div>
      <div class="cp-cell">
        <div class="cp-title">End-User Workplace</div>
        <div class="cp-item">M365 MZ tenant provisioned Dec 2026</div>
        <div class="cp-item">AD forest ready; federation live Mar 2027</div>
        <div class="cp-item">Wave 1 user migration pilot: Jul 2027</div>
        <div class="cp-item">Mass migration (10,000 users): Aug-Nov 2027</div>
        <div class="cp-item">All 12,000 users migrated by QG4</div>
      </div>
      <div class="cp-cell">
        <div class="cp-title">Programme Principles</div>
        <div class="cp-item">GoLive 1 Jan 2028 is hard business deadline</div>
        <div class="cp-item">No SAP feature development in hypercare</div>
        <div class="cp-item">Infosys sole MZ delivery partner; no re-tendering</div>
        <div class="cp-item">GDPR resolved before any PII enters MZ</div>
        <div class="cp-item">Zero TSA extension after QG5</div>
        <div class="cp-item">KPMG fortnightly programme health reviews</div>
      </div>
    </div>
  </div>
</div>

<!-- FOOTER -->
<div class="footer">
  Trinity-CAM Executive Dashboard &nbsp;|&nbsp; {REPORT} &nbsp;|&nbsp; Data sources: Trinity-CAM_Project_Schedule.xlsx, Trinity-CAM_Risk_Register.xlsx, Trinity-CAM_Cost_Plan.xlsx &nbsp;|&nbsp; CONFIDENTIAL
</div>

</div><!-- /page -->
</body>
</html>"""

OUTPUT.write_text(HTML, encoding="utf-8")
print(f"[Trinity-CAM] Executive Dashboard: {OUTPUT}")
