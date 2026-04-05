"""
generate_bravo_kpi_dashboard.py
Generates Charlie_Management_KPI_Dashboard.html

Project Charlie: Robert Bosch GmbH AI Business Carve-Out into 50/50 JV with Undisclosed
"""

import base64
from pathlib import Path
from datetime import date

HERE     = Path(__file__).parent
OUT      = HERE / "Charlie" / "Charlie_Management_KPI_Dashboard.html"
logo_b64 = base64.b64encode((HERE / "Bosch.png").read_bytes()).decode()
today    = date.today()

def days_to(target: date) -> int:
    return (target - today).days

golive  = date(2026, 7, 1)
qg0     = date(2026, 4, 30)
qg123   = date(2026, 6, 26)
qg4     = date(2026, 7, 2)
qg5     = date(2026, 10, 30)

# How many days into the project (start Apr 1)
project_start = date(2026, 4, 1)
project_end   = date(2026, 10, 30)
total_days    = (project_end - project_start).days
elapsed_days  = max(0, (today - project_start).days)
pct_elapsed   = min(100, round(elapsed_days / total_days * 100))

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>Project Charlie â€“ Management KPI Dashboard</title>
  <style>
    :root {{
      --navy:#003b6e; --mid:#005199; --acc:#0066CC; --lt:#e4edf9;
      --bg:#f4f6f9;   --card:#fff;    --ink:#1a1a1a; --muted:#5a6478;
      --line:#d8dde8; --good:#007A33; --warn:#E8A000; --bad:#CC0000;
    }}
    *{{box-sizing:border-box;margin:0;padding:0}}
    body{{font-family:"Segoe UI",system-ui,Arial,sans-serif;background:var(--bg);color:var(--ink);font-size:13px;line-height:1.6}}
    .page{{max-width:1200px;margin:0 auto;padding:22px 20px 52px}}

    /* HEADER */
    .header{{
      background:linear-gradient(135deg,#001f45 0%,var(--navy) 55%,var(--acc) 100%);
      border-radius:14px;color:#fff;padding:22px 28px;margin-bottom:14px;
      display:flex;justify-content:space-between;align-items:flex-start;
    }}
    .header h1{{font-size:22px;font-weight:700;margin-bottom:3px}}
    .header .sub{{font-size:12px;opacity:.85}}
    .bosch-logo{{display:flex;align-items:center;background:#fff;padding:4px 8px;border-radius:4px;width:fit-content;margin-bottom:10px}}
    .header-right{{text-align:right;font-size:12px;opacity:.85}}
    .header-right .big{{font-size:32px;font-weight:800;color:#ffd700;line-height:1}}

    /* 12-COL GRID */
    .grid-12{{display:grid;grid-template-columns:repeat(12,1fr);gap:10px;margin-bottom:12px}}
    .col-2{{grid-column:span 2}} .col-3{{grid-column:span 3}} .col-4{{grid-column:span 4}}
    .col-6{{grid-column:span 6}} .col-8{{grid-column:span 8}} .col-12{{grid-column:span 12}}

    /* CARD */
    .card{{background:var(--card);border-radius:10px;padding:16px 18px}}
    .card-title{{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);margin-bottom:10px}}
    .card-big{{font-size:36px;font-weight:800;color:var(--navy);line-height:1}}
    .card-sub{{font-size:11px;color:var(--muted);margin-top:4px}}

    /* KPI NUMBER */
    .kpi{{text-align:center}}
    .kpi .n{{font-size:38px;font-weight:800;line-height:1}}
    .kpi .lbl{{font-size:11px;color:var(--muted);margin-top:4px}}
    .green-n{{color:var(--good)}} .amber-n{{color:var(--warn)}} .red-n{{color:var(--bad)}} .blue-n{{color:var(--navy)}}

    /* PROGRESS BAR */
    .prog{{margin-top:8px}}
    .prog-row{{display:flex;align-items:center;margin-bottom:6px;font-size:12px}}
    .prog-label{{width:160px;color:var(--ink)}}
    .prog-wrap{{flex:1;background:#eee;border-radius:4px;height:12px;overflow:hidden;margin:0 8px}}
    .prog-bar{{height:100%;border-radius:4px}}
    .prog-val{{width:40px;text-align:right;font-weight:700;font-size:11px}}

    /* PILL */
    .pill{{display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:700;color:#fff}}
    .pg{{background:var(--good)}} .pa{{background:var(--warn);color:#1a1a1a}} .pr{{background:var(--bad)}}

    /* TABLE */
    table{{width:100%;border-collapse:collapse}}
    th,td{{text-align:left;padding:8px 9px;border-bottom:1px solid var(--line);font-size:12px}}
    th{{background:var(--lt);font-weight:700;color:var(--navy)}}

    /* GAUGE-LIKE */
    .gauge-wrap{{display:flex;gap:8px;margin-top:6px}}
    .gauge{{flex:1;padding:10px;border-radius:8px;text-align:center;font-size:12px;font-weight:700}}
    .gauge .gn{{font-size:24px;font-weight:800}}

    /* HORIZON BARS */
    .hbar{{display:flex;height:24px;border-radius:6px;overflow:hidden;margin-top:10px}}
    .hb{{display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:#fff}}

    /* 90-DAY */
    .action-row{{display:grid;grid-template-columns:80px 1fr;gap:10px;align-items:start;margin-bottom:8px;font-size:12px}}
    .action-date{{background:var(--lt);border-radius:6px;padding:6px;text-align:center;font-weight:700;color:var(--navy);font-size:11px}}
    .action-body{{color:var(--ink)}}

    /* FOOTER */
    .footer{{background:var(--lt);border-radius:10px;padding:10px;margin-top:14px;font-size:11px;color:var(--muted);text-align:center}}
  </style>
</head>
<body>
<div class="page">

  <!-- HEADER -->
  <div class="header">
    <div>
      <div class="bosch-logo">
        <img src="data:image/png;base64,{logo_b64}" alt="Bosch" style="height:36px;display:block;"/>
      </div>
      <h1>Project Charlie â€” Management KPI Dashboard</h1>
      <div class="sub">Robert Bosch GmbH AI Business â†’ 50/50 JV with Undisclosed &nbsp;|&nbsp; Bosch Leadership Control &nbsp;|&nbsp; Stand Alone Model &nbsp;|&nbsp; No TSA</div>
      <div class="sub" style="margin-top:4px">PM: Gill Amandeep Singh (BD/MIL-PSM1) &nbsp;|&nbsp; {today.strftime('%d %b %Y')}</div>
    </div>
    <div class="header-right">
      <div class="big">{days_to(golive)}</div>
      <div>days to GoLive</div>
      <div style="margin-top:6px;opacity:.75">01 Jun 2027</div>
    </div>
  </div>

  <!-- ROW 1: KEY KPIs -->
  <div class="grid-12">
    <div class="card col-2 kpi">
      <div class="card-title">SPI</div>
      <div class="n green-n">1.00</div>
      <div class="lbl">On Schedule</div>
    </div>
    <div class="card col-2 kpi">
      <div class="card-title">CPI</div>
      <div class="n green-n">1.00</div>
      <div class="lbl">On Budget</div>
    </div>
    <div class="card col-2 kpi">
      <div class="card-title">Day 1 Readiness</div>
      <div class="n amber-n">{pct_elapsed}%</div>
      <div class="lbl">Time Elapsed</div>
    </div>
    <div class="card col-2 kpi">
      <div class="card-title">Apps Migrated</div>
      <div class="n blue-n">0/17</div>
      <div class="lbl">Planned Jun 2026</div>
    </div>
    <div class="card col-2 kpi">
      <div class="card-title">Open Risks</div>
      <div class="n amber-n">3</div>
      <div class="lbl">Amber Active</div>
    </div>
    <div class="card col-2 kpi">
      <div class="card-title">Open Issues</div>
      <div class="n green-n">0</div>
      <div class="lbl">Critical</div>
    </div>
  </div>

  <!-- ROW 2: WORKSTREAM CONFIDENCE + MILESTONE TIMELINE -->
  <div class="grid-12">
    <div class="card col-6">
      <div class="card-title">Workstream Confidence</div>
      <div class="prog">
        <div class="prog-row"><div class="prog-label">PMO &amp; Governance</div><div class="prog-wrap"><div class="prog-bar" style="width:20%;background:var(--good)"></div></div><div class="prog-val" style="color:var(--good)">HIGH</div></div>
        <div class="prog-row"><div class="prog-label">JV Legal &amp; Entity Setup</div><div class="prog-wrap"><div class="prog-bar" style="width:15%;background:var(--warn)"></div></div><div class="prog-val" style="color:var(--warn)">MED</div></div>
        <div class="prog-row"><div class="prog-label">Infrastructure &amp; Cloud</div><div class="prog-wrap"><div class="prog-bar" style="width:5%;background:var(--good)"></div></div><div class="prog-val" style="color:var(--good)">HIGH</div></div>
        <div class="prog-row"><div class="prog-label">Application Migration (17)</div><div class="prog-wrap"><div class="prog-bar" style="width:5%;background:var(--good)"></div></div><div class="prog-val" style="color:var(--good)">HIGH</div></div>
        <div class="prog-row"><div class="prog-label">Data Separation</div><div class="prog-wrap"><div class="prog-bar" style="width:5%;background:var(--warn)"></div></div><div class="prog-val" style="color:var(--warn)">MED</div></div>
        <div class="prog-row"><div class="prog-label">Client Workplace (70)</div><div class="prog-wrap"><div class="prog-bar" style="width:5%;background:var(--good)"></div></div><div class="prog-val" style="color:var(--good)">HIGH</div></div>
        <div class="prog-row"><div class="prog-label">Security &amp; IAM</div><div class="prog-wrap"><div class="prog-bar" style="width:5%;background:var(--good)"></div></div><div class="prog-val" style="color:var(--good)">HIGH</div></div>
        <div class="prog-row"><div class="prog-label">HR IT</div><div class="prog-wrap"><div class="prog-bar" style="width:5%;background:var(--good)"></div></div><div class="prog-val" style="color:var(--good)">HIGH</div></div>
        <div class="prog-row"><div class="prog-label">Licence &amp; Procurement</div><div class="prog-wrap"><div class="prog-bar" style="width:5%;background:var(--warn)"></div></div><div class="prog-val" style="color:var(--warn)">MED</div></div>
      </div>
    </div>

    <div class="card col-6">
      <div class="card-title">Milestone Gate Control</div>
      <table>
        <tr><th>Gate</th><th>Date</th><th>Days</th><th>Status</th></tr>
        <tr><td>QG0 â€“ Initialization</td><td>30 Apr 2026</td><td>{days_to(qg0):+d}</td><td><span class="pill pa">UPCOMING</span></td></tr>
        <tr><td>QG1/2/3 â€“ Combined</td><td>26 Jun 2026</td><td>{days_to(qg123):+d}</td><td><span class="pill pa">PLANNED</span></td></tr>
        <tr><td>GoLive â€“ Day 1</td><td>01 Jun 2027</td><td>{days_to(golive):+d}</td><td><span class="pill pa">PLANNED</span></td></tr>
        <tr><td>QG4 â€“ GoLive Gate</td><td>02 Jul 2026</td><td>{days_to(qg4):+d}</td><td><span class="pill pa">PLANNED</span></td></tr>
        <tr><td>QG5 â€“ Programme Close</td><td>31 Oct 2027</td><td>{days_to(qg5):+d}</td><td><span class="pill pa">PLANNED</span></td></tr>
      </table>
      <div style="margin-top:12px">
        <div style="font-size:11px;font-weight:700;margin-bottom:4px;color:var(--muted)">Programme Timeline Progress</div>
        <div style="background:#eee;border-radius:6px;height:18px;overflow:hidden">
          <div style="width:{pct_elapsed}%;height:100%;background:var(--acc);border-radius:6px"></div>
        </div>
        <div style="font-size:11px;color:var(--muted);margin-top:3px">{pct_elapsed}% elapsed ({elapsed_days} days of {total_days} days)</div>
      </div>
    </div>
  </div>

  <!-- ROW 3: TOP RISKS + JV MODEL NOTES -->
  <div class="grid-12">
    <div class="card col-8">
      <div class="card-title">Top Risk Register (Amber / Red Active)</div>
      <table>
        <tr><th>#</th><th>Risk Description</th><th>Cat</th><th>PÃ—I</th><th>Status</th><th>Owner</th></tr>
        <tr>
          <td>1</td>
          <td>Robert Bosch GmbH/JV bandwidth constrained during Aprâ€“Jun sprint</td>
          <td>ScR</td><td><strong>16</strong></td>
          <td><span class="pill pa">Amber</span></td>
          <td>Riyaz Ahmed, Robert Bosch GmbH CIO</td>
        </tr>
        <tr>
          <td>2</td>
          <td>JV legal entity (India MCA) delayed beyond QG0</td>
          <td>SR</td><td><strong>15</strong></td>
          <td><span class="pill pg">Green</span></td>
          <td>Riyaz Ahmed, Legal</td>
        </tr>
        <tr>
          <td>3</td>
          <td>India cloud / M365 licensing costs exceed estimate</td>
          <td>BtR</td><td><strong>9</strong></td>
          <td><span class="pill pa">Amber</span></td>
          <td>Finance, Robert Bosch GmbH Procurement</td>
        </tr>
        <tr>
          <td>5</td>
          <td>Key Robert Bosch GmbH IT staff unavailable in Jun build window</td>
          <td>RR</td><td><strong>12</strong></td>
          <td><span class="pill pa">Amber</span></td>
          <td>Robert Bosch GmbH HR, Riyaz Ahmed</td>
        </tr>
        <tr>
          <td>10</td>
          <td>Undisclosed JV IT team not staffed for May architecture workshops</td>
          <td>SR</td><td><strong>9</strong></td>
          <td><span class="pill pa">Amber</span></td>
          <td>Riyaz Ahmed, Undisclosed Sponsor</td>
        </tr>
      </table>
    </div>
    <div class="card col-4">
      <div class="card-title">JV Model Key Facts</div>
      <table>
        <tr><th>Parameter</th><th>Value</th></tr>
        <tr><td>Structure</td><td>50/50 Boschâ€“Undisclosed JV</td></tr>
        <tr><td>Bosch Control</td><td>Leadership control retained</td></tr>
        <tr><td>Antitrust</td><td>Not applicable</td></tr>
        <tr><td>TSA Required</td><td>None (No TSA)</td></tr>
        <tr><td>Merger Zone</td><td>Not required</td></tr>
        <tr><td>Post-GoLive Ops</td><td>Bosch governance channels</td></tr>
        <tr><td>Complexity vs Stand Alone</td><td>Significantly lower</td></tr>
      </table>
    </div>
  </div>

  <!-- ROW 4: BUDGET PERFORMANCE + 90-DAY ACTIONS -->
  <div class="grid-12">
    <div class="card col-6">
      <div class="card-title">Budget Performance</div>
      <div class="gauge-wrap">
        <div class="gauge" style="background:var(--lt)">
          <div class="gn" style="color:var(--navy)">EUR 554K</div>
          <div style="font-size:11px;color:var(--muted)">Labour Baseline</div>
          <div style="font-size:11px;color:var(--good);margin-top:4px">â— On target</div>
        </div>
        <div class="gauge" style="background:#fff8e1">
          <div class="gn" style="color:var(--warn)">TBC</div>
          <div style="font-size:11px;color:var(--muted)">CAPEX Budget</div>
          <div style="font-size:11px;color:var(--warn);margin-top:4px">â— Approval at QG0</div>
        </div>
        <div class="gauge" style="background:#f0fff5">
          <div class="gn" style="color:var(--good)">EUR 83K</div>
          <div style="font-size:11px;color:var(--muted)">Contingency (15%)</div>
          <div style="font-size:11px;color:var(--good);margin-top:4px">â— Reserved</div>
        </div>
      </div>
      <div style="margin-top:14px">
        <div style="font-size:11px;font-weight:700;margin-bottom:6px;color:var(--muted)">Phase Cost Distribution</div>
        <div class="hbar">
          <div class="hb" style="flex:10;background:#005199">P1 10%</div>
          <div class="hb" style="flex:20;background:#0066CC">P2 20%</div>
          <div class="hb" style="flex:30;background:#0088E0">P3 30%</div>
          <div class="hb" style="flex:5;background:#007A33">P4 5%</div>
          <div class="hb" style="flex:36;background:#004a1e">P5 36%</div>
        </div>
        <div style="font-size:10px;color:var(--muted);margin-top:4px">Phase 5 (Hypercare) is largest due to 60-day stabilisation window</div>
      </div>
    </div>

    <div class="card col-6">
      <div class="card-title">Next 90-Day Action Forecast</div>
      <div class="action-row">
        <div class="action-date">Apr 3â€“7</div>
        <div class="action-body"><strong>Engage India legal counsel</strong> for MCA JV registration filing (critical path)</div>
      </div>
      <div class="action-row">
        <div class="action-date">Apr 10</div>
        <div class="action-body"><strong>Confirm Undisclosed JV IT team nominees</strong> for May architecture workshops</div>
      </div>
      <div class="action-row">
        <div class="action-date">Apr 14</div>
        <div class="action-body"><strong>Complete 17-app inventory and scope freeze</strong>; confirm 70-user list</div>
      </div>
      <div class="action-row">
        <div class="action-date">Apr 30</div>
        <div class="action-body"><strong>QG0 Steering Committee</strong> â€” charter approval; CAPEX budget sign-off</div>
      </div>
      <div class="action-row">
        <div class="action-date">May 1</div>
        <div class="action-body"><strong>Start Phase 2</strong> â€” app deep-dives, architecture workshops, M365/Azure RFQ</div>
      </div>
      <div class="action-row">
        <div class="action-date">May 14</div>
        <div class="action-body"><strong>Complete licence/change-of-control review</strong> for all 17 AI apps</div>
      </div>
      <div class="action-row">
        <div class="action-date">May 29</div>
        <div class="action-body"><strong>Enforce June leave freeze</strong> for critical Robert Bosch GmbH team; knowledge documented</div>
      </div>
      <div class="action-row">
        <div class="action-date">Jun 1</div>
        <div class="action-body"><strong>Start Phase 3 build</strong> â€” AD forest, M365 tenant, Azure, network (37 worldwide sites)</div>
      </div>
    </div>
  </div>

  <div class="footer">
    <strong>Project Charlie | Management KPI Dashboard</strong> &nbsp;|&nbsp;
    Seller: Robert Bosch GmbH &nbsp;|&nbsp; Buyer: Undisclosed &nbsp;|&nbsp;
    Report Date: {today.strftime('%d %b %Y')} &nbsp;|&nbsp;
    PM: Gill Amandeep Singh (BD/MIL-PSM1) &nbsp;|&nbsp;
    Data: Charlie_Project_Schedule.csv / Charlie_Risk_Register.xlsx / Charlie_Cost_Plan.csv &nbsp;|&nbsp;
    Confidential â€” Internal Use Only
  </div>

</div>
</body>
</html>"""

OUT.write_text(html, encoding="utf-8")
print(f"KPI Dashboard written: {OUT}  ({len(html):,} chars)")


