"""
generate_bravo_executive_dashboard.py
Generates Bravo_Executive_Dashboard.html

Project Bravo: BGSW AI Business Carve-Out into 50/50 JV with Tata
"""

import base64
from pathlib import Path
from datetime import date

HERE     = Path(__file__).parent
OUT      = HERE / "Bravo" / "Bravo_Executive_Dashboard.html"
logo_b64 = base64.b64encode((HERE / "Bosch.png").read_bytes()).decode()
today    = date.today()

def days_to(target: date) -> int:
    return (target - today).days

golive     = date(2026, 7, 1)
qg0        = date(2026, 4, 30)
qg123      = date(2026, 6, 26)
qg4        = date(2026, 7, 2)
qg5        = date(2026, 10, 30)
signing    = date(2026, 4, 7)
hypercare  = date(2026, 9, 25)
closure    = date(2026, 10, 30)

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>Project Bravo – Executive Dashboard</title>
  <style>
    :root {{
      --navy:#003b6e; --mid:#005199; --acc:#0066CC; --lt:#e4edf9;
      --bg:#f4f6f9;   --card:#fff;    --ink:#1a1a1a; --muted:#5a6478;
      --line:#d8dde8; --good:#007A33; --warn:#E8A000; --bad:#CC0000;
    }}
    *{{box-sizing:border-box;margin:0;padding:0}}
    body{{font-family:"Segoe UI",system-ui,Arial,sans-serif;background:var(--bg);color:var(--ink);font-size:13px;line-height:1.6}}
    .page{{max-width:1100px;margin:0 auto;padding:24px 20px 56px}}

    /* ── HEADER ── */
    .header{{
      background:linear-gradient(135deg,#001f45 0%,var(--navy) 55%,var(--acc) 100%);
      border-radius:16px;color:#fff;padding:28px 32px 24px;margin-bottom:16px;
      display:flex;justify-content:space-between;align-items:flex-start;
    }}
    .header-left h1{{font-size:26px;font-weight:700;margin-bottom:4px}}
    .header-left .sub{{font-size:13px;opacity:.85}}
    .header-right{{text-align:right;font-size:12px;opacity:.85}}
    .header-right .days{{font-size:22px;font-weight:700;color:#ffd700}}
    .bosch-logo{{display:flex;align-items:center;background:#fff;padding:4px 8px;border-radius:4px;width:fit-content;margin-bottom:12px}}

    /* ── COUNTDOWN STRIP ── */
    .strip{{
      display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:16px;
    }}
    .strip-card{{
      background:var(--navy);border-radius:10px;color:#fff;
      padding:14px 16px;text-align:center;
    }}
    .strip-card .n{{font-size:28px;font-weight:800;color:#ffd700}}
    .strip-card .lbl{{font-size:11px;opacity:.8;margin-top:2px}}
    .strip-card .dt {{font-size:11px;opacity:.65;margin-top:2px}}

    /* ── SECTIONS ── */
    .section{{background:var(--card);border-radius:12px;padding:20px 22px;margin-bottom:14px}}
    .section-title{{
      font-size:13px;font-weight:700;text-transform:uppercase;
      letter-spacing:.8px;color:#fff;background:var(--mid);
      padding:6px 14px;border-radius:6px;margin-bottom:14px;display:inline-block;
    }}
    .grid-2{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}
    .grid-3{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}}
    .grid-6{{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:14px}}

    /* ── STAT BOXES ── */
    .stat{{background:var(--lt);border-radius:8px;padding:14px;text-align:center}}
    .stat .n{{font-size:26px;font-weight:800;color:var(--navy)}}
    .stat .lbl{{font-size:11px;color:var(--muted);margin-top:2px}}

    /* ── PHASE TIMELINE ── */
    .timeline{{display:flex;gap:4px;margin-top:10px;border-radius:8px;overflow:hidden}}
    .tphase{{flex:1;padding:10px 8px;text-align:center;font-size:11px;font-weight:700;color:#fff}}
    .p1{{background:#005199}}
    .p2{{background:#0066CC}}
    .p3{{background:#0088E0}}
    .p4{{background:#007A33}}
    .p5{{background:#005E27}}

    /* ── MILESTONES TABLE ── */
    table{{width:100%;border-collapse:collapse}}
    th,td{{text-align:left;padding:9px 10px;border-bottom:1px solid var(--line);font-size:12px}}
    th{{background:var(--lt);font-weight:700;color:var(--navy)}}
    .pill{{display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:700;color:#fff}}
    .g{{background:var(--good)}} .a{{background:var(--warn);color:#1a1a1a}} .r{{background:var(--bad)}}

    /* ── WORKSTREAM GRID ── */
    .ws-card{{background:var(--lt);border-radius:8px;padding:12px 14px}}
    .ws-title{{font-weight:700;color:var(--navy);font-size:12px;margin-bottom:6px}}
    .ws-body{{font-size:11px;color:var(--muted);line-height:1.5}}
    .ws-conf{{font-size:11px;font-weight:700;margin-top:6px;padding:2px 8px;border-radius:10px;display:inline-block}}
    .c-high{{background:#cce8d5;color:#007A33}}
    .c-med {{background:#fff3cc;color:#b06000}}
    .c-low {{background:#fdd;color:#aa0000}}

    /* ── RISK CARDS ── */
    .risk-card{{border-left:4px solid var(--bad);padding:10px 14px;background:#fff8f8;border-radius:0 6px 6px 0;margin-bottom:8px;font-size:12px}}
    .risk-card.amber{{border-color:var(--warn);background:#fff8e1}}
    .risk-card .rt{{font-weight:700;color:var(--ink);margin-bottom:2px}}
    .risk-card .rm{{color:var(--muted)}}

    /* ── BUDGET ── */
    .bcat{{display:flex;align-items:center;margin-bottom:8px;font-size:12px}}
    .bcat .label{{width:220px;color:var(--ink)}}
    .bcat .bar-wrap{{flex:1;background:#eee;border-radius:4px;height:14px;overflow:hidden;margin:0 10px}}
    .bcat .bar{{height:100%;background:var(--acc);border-radius:4px}}
    .bcat .val{{width:80px;text-align:right;font-weight:700;color:var(--navy)}}

    /* ── FOOTER ── */
    .footer{{background:var(--lt);border-radius:10px;padding:12px;margin-top:18px;font-size:11px;color:var(--muted);text-align:center}}

    /* ── PRINT ── */
    @media print {{
      .page-break{{page-break-before:always}}
    }}
  </style>
</head>
<body>
<div class="page">

  <!-- HEADER -->
  <div class="header">
    <div class="header-left">
      <div class="bosch-logo">
        <img src="data:image/png;base64,{logo_b64}" alt="Bosch" style="height:36px;display:block;"/>
      </div>
      <h1>Project Bravo — Executive Dashboard</h1>
      <div class="sub">BGSW AI Business Carve-Out → 50/50 JV with Tata &nbsp;|&nbsp; Bosch Leadership Control &nbsp;|&nbsp; Combination Model</div>
      <div class="sub" style="margin-top:6px">Seller: Bosch BGSW &nbsp;|&nbsp; Buyer: Tata &nbsp;|&nbsp; PM: Riyaz Ahmed Syed Ahmed (BD/MIL-PSM4)</div>
    </div>
    <div class="header-right">
      <div style="font-size:12px;opacity:.8">Dashboard date: {today.strftime('%d %b %Y')}</div>
      <div class="days">{days_to(golive)}</div>
      <div style="font-size:12px">days to GoLive (01 Jul 2026)</div>
      <div style="margin-top:8px;font-size:11px;opacity:.7">Total Duration: 7 months</div>
    </div>
  </div>

  <!-- COUNTDOWN STRIP -->
  <div class="strip">
    <div class="strip-card">
      <div class="n">{days_to(qg0)}</div>
      <div class="lbl">Days to QG0</div>
      <div class="dt">30 Apr 2026</div>
    </div>
    <div class="strip-card">
      <div class="n">{days_to(qg123)}</div>
      <div class="lbl">Days to QG1/2/3</div>
      <div class="dt">26 Jun 2026</div>
    </div>
    <div class="strip-card">
      <div class="n">{days_to(golive)}</div>
      <div class="lbl">Days to GoLive</div>
      <div class="dt">01 Jul 2026</div>
    </div>
    <div class="strip-card">
      <div class="n">{days_to(qg5)}</div>
      <div class="lbl">Days to Programme Close</div>
      <div class="dt">30 Oct 2026</div>
    </div>
  </div>

  <!-- OVERVIEW + STATS -->
  <div class="section">
    <span class="section-title">Programme Overview</span>
    <div class="grid-2">
      <div>
        <p><strong>Objective:</strong> Carve out Bosch BGSW's AI Business into a newly formed 50/50 Tata–Bosch Joint Venture with Bosch retaining leadership control. The JV operates effectively as a Bosch-governed entity — eliminating the need for a TSA, Merger Zone, or antitrust filings.</p>
        <p style="margin-top:10px"><strong>Scope:</strong> 17 AI applications, 70 users, 2 India sites. No ERP or SAP in scope. Single-wave application migration. No TSA post-GoLive.</p>
        <p style="margin-top:10px"><strong>Key Simplification:</strong> Bosch's operational leadership of the JV dramatically reduces separation complexity versus a full Stand Alone model. BGSW governance and escalation paths remain intact post-GoLive.</p>
      </div>
      <div>
        <div class="grid-2" style="gap:10px">
          <div class="stat"><div class="n">2</div><div class="lbl">India Sites</div></div>
          <div class="stat"><div class="n">70</div><div class="lbl">IT Users</div></div>
          <div class="stat"><div class="n">17</div><div class="lbl">AI Applications</div></div>
          <div class="stat"><div class="n">0</div><div class="lbl">ERP Systems</div></div>
          <div class="stat"><div class="n">7</div><div class="lbl">Months Duration</div></div>
          <div class="stat"><div class="n">0</div><div class="lbl">TSA Services</div></div>
        </div>
      </div>
    </div>
    <div style="margin-top:14px">
      <div style="font-size:12px;font-weight:700;margin-bottom:6px;color:var(--navy)">Phase Timeline</div>
      <div class="timeline">
        <div class="tphase p1" style="flex:22">Phase 1: Initialization<br/><span style="font-weight:400;font-size:10px">01 Apr – 30 Apr 2026</span></div>
        <div class="tphase p2" style="flex:21">Phase 2: Concept &amp; Architecture<br/><span style="font-weight:400;font-size:10px">01 May – 29 May 2026</span></div>
        <div class="tphase p3" style="flex:20">Phase 3: Build &amp; Test<br/><span style="font-weight:400;font-size:10px">01 Jun – 26 Jun 2026</span></div>
        <div class="tphase p4" style="flex:4">P4<br/><span style="font-weight:400;font-size:10px">29 Jun–2 Jul</span></div>
        <div class="tphase p5" style="flex:85">Phase 5: Hypercare &amp; Closure<br/><span style="font-weight:400;font-size:10px">06 Jul – 30 Oct 2026</span></div>
      </div>
    </div>
  </div>

  <!-- MILESTONES + BUDGET -->
  <div class="grid-2" style="margin-bottom:14px">
    <!-- Milestones -->
    <div class="section">
      <span class="section-title">Key Milestones &amp; Quality Gates</span>
      <table>
        <tr><th>Gate</th><th>Date</th><th>Days</th><th>Status</th></tr>
        <tr><td>Signing – Frozen Zone</td><td>07 Apr 2026</td><td>{days_to(signing):+d}</td><td><span class="pill g">DONE</span></td></tr>
        <tr><td>QG0 – Initialization</td><td>30 Apr 2026</td><td>{days_to(qg0):+d}</td><td><span class="pill a">UPCOMING</span></td></tr>
        <tr><td>GoLive (Day 1) – 01 Jul 2026</td><td>01 Jul 2026</td><td>{days_to(golive):+d}</td><td><span class="pill a">PLANNED</span></td></tr>
        <tr><td>QG1/2/3 – Combined Gate</td><td>26 Jun 2026</td><td>{days_to(qg123):+d}</td><td><span class="pill a">PLANNED</span></td></tr>
        <tr><td>QG4 – GoLive Gate</td><td>02 Jul 2026</td><td>{days_to(qg4):+d}</td><td><span class="pill a">PLANNED</span></td></tr>
        <tr><td>Hypercare Close</td><td>25 Sep 2026</td><td>{days_to(hypercare):+d}</td><td><span class="pill a">PLANNED</span></td></tr>
        <tr><td>QG5 – Programme Closure</td><td>30 Oct 2026</td><td>{days_to(qg5):+d}</td><td><span class="pill a">PLANNED</span></td></tr>
      </table>
    </div>
    <!-- Budget -->
    <div class="section">
      <span class="section-title">Budget Distribution</span>
      <p style="font-size:28px;font-weight:800;color:var(--navy)">EUR 554K</p>
      <p style="font-size:12px;color:var(--muted);margin-bottom:14px">Labour only &nbsp;|&nbsp; CAPEX TBC – to be approved at QG0</p>
      <div class="bcat"><div class="label">Programme Management</div><div class="bar-wrap"><div class="bar" style="width:30%"></div></div><div class="val">EUR 164K</div></div>
      <div class="bcat"><div class="label">IT Project Management</div><div class="bar-wrap"><div class="bar" style="width:21%"></div></div><div class="val">EUR 114K</div></div>
      <div class="bcat"><div class="label">Hypercare &amp; Closure</div><div class="bar-wrap"><div class="bar" style="width:15%"></div></div><div class="val">EUR 86K</div></div>
      <div class="bcat"><div class="label">Infrastructure &amp; Cloud</div><div class="bar-wrap"><div class="bar" style="width:11%"></div></div><div class="val">EUR 59K</div></div>
      <div class="bcat"><div class="label">Application Migration</div><div class="bar-wrap"><div class="bar" style="width:11%"></div></div><div class="val">EUR 58K</div></div>
      <div class="bcat"><div class="label">Architecture &amp; Design</div><div class="bar-wrap"><div class="bar" style="width:6%"></div></div><div class="val">EUR 32K</div></div>
      <div class="bcat"><div class="label">Legal &amp; Compliance</div><div class="bar-wrap"><div class="bar" style="width:4%"></div></div><div class="val">EUR 24K</div></div>
      <div class="bcat"><div class="label">Client Workplace</div><div class="bar-wrap"><div class="bar" style="width:3%"></div></div><div class="val">EUR 14K</div></div>
      <div class="bcat"><div class="label">HR IT</div><div class="bar-wrap"><div class="bar" style="width:1%"></div></div><div class="val">EUR 4K</div></div>
    </div>
  </div>

  <!-- PAGE BREAK -->
  <div class="page-break"></div>

  <!-- WORKSTREAM COVERAGE -->
  <div class="section">
    <span class="section-title">IT Workstream Coverage</span>
    <div class="grid-3">
      <div class="ws-card">
        <div class="ws-title">WS1 – PMO &amp; Governance</div>
        <div class="ws-body">Programme setup, RACI, SteerCo facilitation, QG management, change control</div>
        <div class="ws-conf c-high">HIGH Confidence</div>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS2 – JV Legal &amp; Entity Setup</div>
        <div class="ws-body">India MCA registration (50/50 Tata-Bosch; Bosch leadership); no antitrust filing</div>
        <div class="ws-conf c-med">MEDIUM Confidence</div>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS3 – Infrastructure &amp; Cloud</div>
        <div class="ws-body">JV AD forest, M365 &lt;70 mailboxes, Azure tenant, network on 2 India sites</div>
        <div class="ws-conf c-high">HIGH Confidence</div>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS4 – Application Migration</div>
        <div class="ws-body">17 AI apps in single wave; re-point to JV domain/AD/M365; no ERP dependency</div>
        <div class="ws-conf c-high">HIGH Confidence</div>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS5 – Data Separation</div>
        <div class="ws-body">AI business data separated from BGSW; model artefacts, training data ownership mapped</div>
        <div class="ws-conf c-med">MEDIUM Confidence</div>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS6 – Client Workplace</div>
        <div class="ws-body">70 devices; JV standard image; reimaging across 2 India sites; no travel required</div>
        <div class="ws-conf c-high">HIGH Confidence</div>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS7 – Security &amp; IAM</div>
        <div class="ws-body">JV-independent security platform; ISO 27001 baseline; BGSW CISO standards retained</div>
        <div class="ws-conf c-high">HIGH Confidence</div>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS8 – HR IT</div>
        <div class="ws-body">India payroll &amp; HR mapping for 70 AI business users; 2 India site coverage</div>
        <div class="ws-conf c-high">HIGH Confidence</div>
      </div>
      <div class="ws-card">
        <div class="ws-title">WS9 – Licence &amp; Procurement</div>
        <div class="ws-body">Change-of-control review on 17 AI app licences; M365/Azure JV procurement</div>
        <div class="ws-conf c-med">MEDIUM Confidence</div>
      </div>
    </div>
  </div>

  <!-- QG TRACKER + RISKS -->
  <div class="grid-2" style="margin-bottom:14px">
    <!-- QG Tracker -->
    <div class="section">
      <span class="section-title">Quality Gate Tracker</span>
      <table>
        <tr><th>Gate</th><th>Date</th><th>Criteria</th></tr>
        <tr>
          <td><strong>QG0</strong></td><td>30 Apr 2026</td>
          <td>Charter approved; 17-app inventory confirmed; JV model &amp; governance agreed; proceed to Concept</td>
        </tr>
        <tr>
          <td><strong>QG1/2/3</strong></td><td>26 Jun 2026</td>
          <td>Architecture approved; all 17 apps tested; 70 devices configured; JV infra ready; UAT passed; proceed to Cutover</td>
        </tr>
        <tr>
          <td><strong>QG4</strong></td><td>02 Jul 2026</td>
          <td>Day 1 GoLive confirmed; all 17 apps live in JV; 70 users on JV domain; Bosch-led JV operational; no TSA; proceed to Hypercare</td>
        </tr>
        <tr>
          <td><strong>QG5</strong></td><td>30 Oct 2026</td>
          <td>Hypercare stable; JV IT fully operational; all 17 apps stable; Project Bravo formally closed</td>
        </tr>
      </table>
    </div>
    <!-- Risks -->
    <div class="section">
      <span class="section-title">Key Risk Indicators</span>
      <div class="risk-card amber">
        <div class="rt">BGSW/JV Bandwidth — P4×I4 = 16 ● AMBER</div>
        <div class="rm">3-month sprint Apr–Jun; split focus. Dedicated lead + SteerCo cadence in place.</div>
      </div>
      <div class="risk-card amber">
        <div class="rt">JV Legal Registration (India MCA) — P3×I5 = 15 ● AMBER</div>
        <div class="rm">4–8 week MCA processing time. External India counsel engaged immediately.</div>
      </div>
      <div class="risk-card amber">
        <div class="rt">Jun Build — Key BGSW Staff Unavailable — P3×I4 = 12 ● AMBER</div>
        <div class="rm">June leave freeze enforced; backups identified by 29 May.</div>
      </div>
      <div class="risk-card" style="border-color:var(--good);background:#f0fff5">
        <div class="rt" style="color:var(--good)">AD JV Forest Separation — P2×I5 = 10 ● GREEN</div>
        <div class="rm">AD build starts Jun 1; dress rehearsal Jun 24–26; no backout plan needed.</div>
      </div>
      <div class="risk-card" style="border-color:var(--good);background:#f0fff5">
        <div class="rt" style="color:var(--good)">UAT Window Compressed — P3×I4 = 12 ● GREEN</div>
        <div class="rm">UAT leads engaged May 1; test environment ready Jun 1; smoke tests daily from Jun 15.</div>
      </div>
    </div>
  </div>

  <!-- PAGE BREAK -->
  <div class="page-break"></div>

  <!-- APPLICATION MIGRATION -->
  <div class="section">
    <span class="section-title">Application Migration — 17 AI Apps (Single Wave)</span>
    <div class="grid-2">
      <div>
        <p><strong>Strategy:</strong> All 17 AI applications migrate in a single wave during Phase 3 (Jun 2026). The small app count and the Bosch-led JV model (same AD/security standards) makes a multi-wave approach unnecessary.</p>
        <p style="margin-top:10px"><strong>Approach:</strong> Re-point each app from BGSW domain / M365 / Azure to JV domain / M365 / Azure. No ERP dependencies. No SAP. Licence change-of-control reviewed by 14 May 2026.</p>
        <p style="margin-top:10px"><strong>Day 1 Activation:</strong> All 17 apps go live simultaneously on 1 Jul 2026 following UAT sign-off on 24 Jun 2026.</p>
      </div>
      <div>
        <table>
          <tr><th>Wave</th><th>Apps</th><th>Window</th><th>Status</th></tr>
          <tr>
            <td><strong>Wave 1 (only wave)</strong></td>
            <td>All 17 AI apps</td>
            <td>01 Jun – 26 Jun 2026</td>
            <td><span class="pill a">PLANNED</span></td>
          </tr>
          <tr>
            <td>App Reconfig</td>
            <td>17 apps</td>
            <td>01–16 Jun 2026</td>
            <td><span class="pill a">PLANNED</span></td>
          </tr>
          <tr>
            <td>Integration Testing</td>
            <td>17 apps</td>
            <td>15–23 Jun 2026</td>
            <td><span class="pill a">PLANNED</span></td>
          </tr>
          <tr>
            <td>UAT (70 users)</td>
            <td>17 apps</td>
            <td>18–24 Jun 2026</td>
            <td><span class="pill a">PLANNED</span></td>
          </tr>
          <tr>
            <td>Day 1 Activation</td>
            <td>17 apps</td>
            <td>01 Jul 2026</td>
            <td><span class="pill a">PLANNED</span></td>
          </tr>
        </table>
      </div>
    </div>
  </div>

  <!-- REGIONAL SCOPE + CRITICAL PATH -->
  <div class="grid-2" style="margin-bottom:14px">
    <!-- Regional -->
    <div class="section">
      <span class="section-title">Regional Scope (India Only)</span>
      <table>
        <tr><th>Site</th><th>Role</th><th>Users</th><th>Devices</th></tr>
        <tr><td><strong>India Site 1</strong></td><td>Primary JV Hub</td><td>~45</td><td>~45</td></tr>
        <tr><td><strong>India Site 2</strong></td><td>Secondary JV Site</td><td>~25</td><td>~25</td></tr>
        <tr><td><strong>Total</strong></td><td>India (domestic only)</td><td><strong>70</strong></td><td><strong>~70</strong></td></tr>
      </table>
      <div style="margin-top:12px;font-size:12px;color:var(--muted)">
        Domestic India connectivity only. No international WAN circuits required. Cloud-first approach (Azure India) minimises on-premise footprint.
      </div>
    </div>
    <!-- Critical Path -->
    <div class="section">
      <span class="section-title">Critical Path</span>
      <ol style="font-size:12px;margin-left:18px">
        <li style="margin-bottom:8px"><strong>JV Legal Entity Registration (India)</strong> — Must be filed in week 1 of Apr; completed by 30 Apr (QG0). <span class="pill a" style="font-size:10px">CRITICAL</span></li>
        <li style="margin-bottom:8px"><strong>BGSW Resource Commitment (Jun)</strong> — BGSW infra, AD, and dev teams reserved for all of Jun. <span class="pill a" style="font-size:10px">CRITICAL</span></li>
        <li style="margin-bottom:8px"><strong>AD JV Forest Build (Jun 1–12)</strong> — Foundation for all 70 users and 17 apps. No parallel path. <span class="pill a" style="font-size:10px">CRITICAL</span></li>
        <li style="margin-bottom:8px"><strong>UAT Sign-Off (24 Jun)</strong> — Enables QG1/2/3 on 26 Jun; enables GoLive on 1 Jul. <span class="pill a" style="font-size:10px">CRITICAL</span></li>
        <li style="margin-bottom:8px"><strong>Day 1 GoLive (1 Jul 2026)</strong> — Legal JV effective date; tied to JV shareholder agreement. <span class="pill r" style="font-size:10px">HARD DATE</span></li>
      </ol>
    </div>
  </div>

  <div class="footer">
    <strong>Project Bravo | Executive Dashboard</strong> &nbsp;|&nbsp;
    Seller: Bosch BGSW &nbsp;|&nbsp; Buyer: Tata &nbsp;|&nbsp;
    Date: {today.strftime('%d %b %Y')} &nbsp;|&nbsp;
    Data: Bravo_Project_Schedule.csv / Bravo_Risk_Register.xlsx / Bravo_Cost_Plan.csv &nbsp;|&nbsp;
    Confidential — Internal Use Only
  </div>

</div>
</body>
</html>"""

OUT.write_text(html, encoding="utf-8")
print(f"Executive Dashboard written: {OUT}  ({len(html):,} chars)")
