"""
generate_bravo_charter.py
Generates Charlie_Project_Charter.html

Project Charlie: Robert Bosch GmbH AI Business Carve-Out into 50/50 JV with Undisclosed
stand-alone separation | 37 worldwide sites | 3500+ users | TBD apps | No ERP | No TSA
"""

import base64
from datetime import datetime as _dt
from pathlib import Path

_t0 = _dt.now()
print(f"Started : {_t0.strftime('%Y-%m-%d %H:%M:%S')}")

HERE = Path(__file__).parent
OUT  = HERE / "Charlie" / "Charlie_Project_Charter.html"

logo_b64 = base64.b64encode((HERE / "Bosch.png").read_bytes()).decode()

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Project Charlie â€“ Project Charter</title>
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
      background:#fff; padding:4px 8px; border-radius:4px;
      width:fit-content; margin-bottom:14px;
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
        <img src="data:image/png;base64,{logo_b64}" alt="Bosch â€” Invented for Life" style="height:36px;display:block;" />
      </div>
      <div style="font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;margin-bottom:12px;opacity:0.9;">Project Charter</div>
      <h1>Project Charlie</h1>
      <p>Robert Bosch GmbH AI Business Carve-Out into 50/50 JV with Undisclosed | Bosch Leadership Control | Stand Alone Model</p>
      <div class="cover-meta">
        <div class="meta-item">
          <span class="meta-label">Project Duration</span>
          01 Apr 2026 â€“ 31 Oct 2027 (7 months)
        </div>
        <div class="meta-item">
          <span class="meta-label">Scope</span>
          2 India Sites Â· 3500+ Users Â· TBD Apps Â· No ERP Â· No TSA
        </div>
        <div class="meta-item">
          <span class="meta-label">Carve-Out Model</span>
          Stand Alone (Bosch-led JV with Undisclosed)
        </div>
        <div class="meta-item">
          <span class="meta-label">Budget Baseline</span>
          EUR 554K (Labour) + CAPEX TBC â€“ to be approved at QG0
        </div>
      </div>
    </div>

    <!-- PROJECT OVERVIEW -->
    <div class="section">
      <h2>Project Overview</h2>
      <p><strong>Objective:</strong> Execute the carve-out of the Robert Bosch GmbH AI Business into a newly formed 50/50 Joint Venture (JV) with Undisclosed. Bosch retains leadership control of the JV, making this a non-antitrust-relevant transaction. The project separates AI-specific IT infrastructure, 17 applications, and 3500+ users across 37 worldwide sites from Robert Bosch GmbH's core systems by 1 July 2026 (Day 1 GoLive).</p>
      <p><strong>Strategic Context:</strong> Because Bosch maintains leadership control of the JV, the post-GoLive entity operates effectively as a Bosch-governed company (50% Undisclosed shareholding). This simplifies the separation: no Merger Zone is required, no TSA is needed post-GoLive, and IT continuity risks are significantly reduced. Escalation paths remain via Robert Bosch GmbH governance channels.</p>
      <p><strong>Sponsor Customer (Buyer):</strong> Undisclosed (50% shareholder; Bosch holds leadership control).</p>
      <p><strong>Sponsor Contractor (Seller):</strong> Robert Bosch India (Robert Bosch GmbH) as the carver entity.</p>
      <p><strong>Programme Manager:</strong> Gill Amandeep Singh (BD/MIL-PSM1)</p>
    </div>

    <!-- SCOPE & MODEL -->
    <div class="section">
      <h2>Scope &amp; Model</h2>
      <div class="grid-2">
        <div>
          <h3>In Scope</h3>
          <ul>
            <li>17 AI applications (all in single migration wave)</li>
            <li>70 IT users across 37 worldwide sites</li>
            <li>~70 client devices (reimaging to JV standard image)</li>
            <li>M365 and Azure JV tenant (&lt;70 mailboxes)</li>
            <li>JV Active Directory forest (new, separate from Robert Bosch GmbH)</li>
            <li>Network setup across 2 India JV sites</li>
            <li>Security and IAM platform for JV entity</li>
            <li>JV legal entity setup (India; 50/50 Bosch-Undisclosed)</li>
            <li>HR IT mapping for 70 AI business users</li>
          </ul>
        </div>
        <div>
          <h3>Out of Scope</h3>
          <ul>
            <li>ERP systems (no SAP in scope)</li>
            <li>Robert Bosch GmbH core applications (non-AI business)</li>
            <li>TSA services post-GoLive (not required)</li>
            <li>International geographies (India only)</li>
            <li>Manufacturing / OT systems</li>
            <li>Antitrust or regulatory filings (not required)</li>
            <li>Bosch Group shared services (retained by Bosch)</li>
          </ul>
        </div>
      </div>
      <div class="box" style="margin-top:16px;">
        <div class="box-title">Carve-Out Model: Stand Alone (Bosch-led JV)</div>
        <strong>Structure:</strong> 50/50 Boschâ€“Undisclosed JV with Bosch holding leadership/operational control. Post-GoLive, the JV operates effectively as a Bosch-governed entity. <strong>No TSA Required:</strong> Because Bosch leads the JV, Robert Bosch GmbH IT services are accessible via normal governance channels â€” no formal TSA contract is needed. <strong>Antitrust:</strong> Not applicable; stand-alone separation means no merger control filing is required.
      </div>
      <div class="highlight">
        <strong>Key simplification vs. typical carve-outs:</strong> Bosch's JV leadership control eliminates the need for a Merger Zone, a TSA service catalogue, and antitrust regulatory filings. The separation is significantly simpler than a full Stand Alone model.
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
          <td>1 â€“ Initialization</td>
          <td>01 Apr â€“ 30 Apr 2026</td>
          <td>Charter approved; 17-app inventory confirmed; JV model agreed</td>
          <td><span class="pill green">QG0</span></td>
          <td>30 Apr 2026</td>
        </tr>
        <tr>
          <td>2 â€“ Concept &amp; Architecture</td>
          <td>01 May â€“ 29 May 2026</td>
          <td>JV IT architecture designed; app migration plan locked</td>
          <td><em>flows into QG1/2/3</em></td>
          <td>29 May 2026</td>
        </tr>
        <tr>
          <td>3 â€“ Build &amp; Test</td>
          <td>01 Jun â€“ 26 Jun 2026</td>
          <td>JV infra built; TBD apps tested; 70 devices configured; UAT signed</td>
          <td><span class="pill green">QG1/2/3</span></td>
          <td>26 Jun 2026</td>
        </tr>
        <tr>
          <td>4 â€“ Cutover &amp; GoLive</td>
          <td>29 Jun â€“ 02 Jul 2026</td>
          <td>Day 1 GoLive (1 Jul); all TBD apps live; 3500+ users on JV domain</td>
          <td><span class="pill green">QG4</span></td>
          <td>02 Jul 2026</td>
        </tr>
        <tr>
          <td>5 â€“ Hypercare &amp; Closure</td>
          <td>06 Jul â€“ 31 Oct 2027</td>
          <td>60-day hypercare; programme closed; QG5 approval</td>
          <td><span class="pill green">QG5</span></td>
          <td>31 Oct 2027</td>
        </tr>
      </table>
    </div>

    <!-- KEY SUCCESS CRITERIA -->
    <div class="section">
      <h2>Key Success Criteria</h2>
      <ul>
        <li><strong>Day 1 GoLive (1 Jul 2026):</strong> All 70 IT users and all 17 AI applications fully operational in the independent JV environment.</li>
        <li><strong>JV IT Independence:</strong> JV operates under Bosch leadership with its own AD forest, M365 tenant, and network; zero dependency on Robert Bosch GmbH production systems.</li>
        <li><strong>Data Separation:</strong> All AI business data migrated to JV environment without data loss or compliance breach.</li>
        <li><strong>User Acceptance:</strong> UAT sign-off by 24 Jun 2026; hypercare closes 25 Sep 2026 without critical open issues.</li>
        <li><strong>No TSA Needed:</strong> JV self-sufficient from Day 1; Robert Bosch GmbH accessible only via standard Bosch governance channels (not contractual TSA).</li>
        <li><strong>Budget Compliance:</strong> Project delivers within EUR 554K labour budget + approved CAPEX; no uncontrolled overrun.</li>
        <li><strong>Programme Closed by 31 Oct 2027</strong> with QG5 passed and lessons learned documented.</li>
      </ul>
    </div>

    <!-- RISKS & CONSTRAINTS -->
    <div class="section">
      <h2>Key Risks &amp; Constraints</h2>
      <h3>Top Risks</h3>
      <table>
        <tr><th>Risk</th><th>Rating</th><th>Status</th><th>Mitigation</th></tr>
        <tr>
          <td>Robert Bosch GmbH/JV management bandwidth during 3-month sprint (Aprâ€“Jun)</td>
          <td>P4Ã—I4 = 16</td>
          <td><span class="pill amber">Amber</span></td>
          <td>Dedicated carve-out lead; explicit time protection; weekly SteerCo</td>
        </tr>
        <tr>
          <td>JV legal entity (India MCA) registration delayed beyond QG0</td>
          <td>P3Ã—I5 = 15</td>
          <td><span class="pill green">Green</span></td>
          <td>External India counsel engaged by Apr 3; MCA filing in week 1</td>
        </tr>
        <tr>
          <td>Jun 2026 build window: key Robert Bosch GmbH IT staff unavailable</td>
          <td>P3Ã—I4 = 12</td>
          <td><span class="pill amber">Amber</span></td>
          <td>June leave freeze for critical roles; knowledge documented by May 29</td>
        </tr>
        <tr>
          <td>Active Directory JV forest separation incomplete by GoLive</td>
          <td>P2Ã—I5 = 10</td>
          <td><span class="pill green">Green</span></td>
          <td>AD build Jun 1; dress rehearsal Jun 24â€“26; no backout option</td>
        </tr>
        <tr>
          <td>Undisclosed JV IT team not staffed for May architecture workshops</td>
          <td>P3Ã—I3 = 9</td>
          <td><span class="pill amber">Amber</span></td>
          <td>Confirm Undisclosed nominees at SteerCo kickoff Apr 1; Bosch can proceed independently</td>
        </tr>
      </table>
      <h3>Constraints</h3>
      <ul>
        <li><strong>Hard GoLive Date (1 Jul 2026):</strong> Tied to legal JV effective date; no postponement possible.</li>
        <li><strong>Compressed Timeline:</strong> Only 3 months from start to GoLive (Aprâ€“Jun build); resource availability is critical.</li>
        <li><strong>India MCA Regulatory Timeline:</strong> JV entity registration takes 4â€“8 weeks; must be initiated in first week of project.</li>
        <li><strong>Budget TBC:</strong> CAPEX (cloud, networking, licensing) requires QG0 board approval before commitment.</li>
      </ul>
    </div>

    <!-- ORGANISATION -->
    <div class="section">
      <h2>Organisation &amp; Governance</h2>
      <table>
        <tr><th>Role</th><th>Name / Entity</th><th>Responsibility</th></tr>
        <tr>
          <td>Executive Sponsor</td>
          <td>Robert Bosch GmbH Board + Undisclosed</td>
          <td>Overall JV strategy; budget approval; escalation resolution</td>
        </tr>
        <tr>
          <td>Programme Manager</td>
          <td>Gill Amandeep Singh (BD/MIL-PSM1)</td>
          <td>End-to-end programme delivery; PMO governance; SteerCo chair</td>
        </tr>
        <tr>
          <td>Carver Lead</td>
          <td>Robert Bosch GmbH CIO</td>
          <td>Robert Bosch GmbH IT execution; scope boundary ownership; resource allocation</td>
        </tr>
        <tr>
          <td>JV IT Lead</td>
          <td>Undisclosed (to be nominated by Apr 10)</td>
          <td>JV IT design input; architecture workshops; post-GoLive operations</td>
        </tr>
        <tr>
          <td>Steering Committee</td>
          <td>PM + Robert Bosch GmbH CIO + Finance + Legal + Undisclosed Rep</td>
          <td>Weekly reviews; QG decision authority; risk/change control</td>
        </tr>
        <tr>
          <td>Workstream Leads</td>
          <td>Robert Bosch GmbH Infra, Apps, AD, Azure, CISO, CWP, HR IT</td>
          <td>Domain execution; daily stand-ups; risk reporting</td>
        </tr>
      </table>
    </div>

    <!-- ASSUMPTIONS -->
    <div class="section">
      <h2>Assumptions</h2>
      <ul>
        <li>Bosch holds JV leadership/operational control; Undisclosed holds 50% equity only â€” no antitrust filing is required.</li>
        <li>Undisclosed JV IT team representatives will be nominated and engaged by 10 April 2026.</li>
        <li>Robert Bosch GmbH management commits dedicated carve-out lead for the full 7-month duration.</li>
        <li>Budget baseline (EUR 554K labour + TBC CAPEX) is approved at QG0 Steering Committee (30 Apr 2026).</li>
        <li>No competing major IT initiatives in Robert Bosch GmbH during Aprâ€“Jul 2026 that would draw resources from carve-out.</li>
        <li>Application and user scope is stable at TBD apps and 3500+ users; additions post-QG0 require change control.</li>
        <li>Domestic India connectivity and cloud provisioning lead times do not exceed 4 weeks.</li>
        <li>No TSA contract required post-GoLive; JV operates under Bosch governance as a near-Bosch entity.</li>
      </ul>
    </div>

    <div class="footer">
      <strong>Project Charlie | Project Charter</strong><br/>
      Issued: 03 April 2026 &nbsp;|&nbsp; Seller: Robert Bosch GmbH &nbsp;|&nbsp; Buyer: Undisclosed &nbsp;|&nbsp;
      PM: Gill Amandeep Singh (BD/MIL-PSM1) &nbsp;|&nbsp;
      Next Review: QG0 (30 April 2026) &nbsp;|&nbsp; Confidential â€” Internal Use Only
    </div>
  </div>
</body>
</html>"""

OUT.write_text(html, encoding="utf-8")
print(f"Charter written: {OUT}")
_t1 = _dt.now()
print(f"Finished: {_t1.strftime('%Y-%m-%d %H:%M:%S')}  ({(_t1-_t0).total_seconds():.1f}s elapsed)")


