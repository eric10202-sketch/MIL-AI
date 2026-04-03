"""
Build Astra SWS Management Dashboard
Source: Astra Shoppping List with SWS.xlsx
Output: Astra_SWS_Management_Dashboard.html
"""
import openpyxl, os, collections, base64

BASE = r'c:\Users\kho1sgp\OneDrive - Bosch Group\My Work Documents\AI topics\BDMIL use case\Software_license_project'

def clean(v):
    s = str(v or '').strip()
    return s if s not in ('None', '') else ''

# ── Load Bosch logo ──────────────────────────────────────────────────────────
with open(os.path.join(BASE, 'Bosch_png_b64.txt')) as f:
    LOGO_B64 = f.read().strip()
with open(os.path.join(BASE, 'Bosch_color_theme_png_b64.txt')) as f:
    THEME_B64 = f.read().strip()

LOGO_SRC  = f"data:image/png;base64,{LOGO_B64}"
THEME_SRC = f"data:image/png;base64,{THEME_B64}"

# ── Load data ────────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(os.path.join(BASE, 'Astra Shoppping List with SWS.xlsx'))
ws = wb['Shoppinglist']
rows = list(ws.iter_rows(values_only=True))

records = []
for r in rows[1:]:
    vendor  = clean(r[1])
    product = clean(r[2])
    if not vendor and not product:
        continue
    records.append({
        'vendor':       vendor,
        'product':      product,
        'qty':          clean(r[3]),
        'metric':       clean(r[4]),
        'pay_type':     clean(r[5]),
        'sws':          clean(r[6]),
        'source':       clean(r[7]),
        'bdbu':         clean(r[8]),
        'status':       clean(r[9]),
        'req_date':     r[10],
        'purch_status': clean(r[11]),
        'order_by':     clean(r[12]),
        'purch_date':   r[13],
        'notes':        clean(r[14]),
        'sw_contact':   clean(r[15]),
    })

total = len(records)

# ── KPI Computations ─────────────────────────────────────────────────────────
status_cnt    = collections.Counter(r['status'].lower().strip() for r in records if r['status'])
done          = status_cnt.get('done', 0)
out_scope     = status_cnt.get('out of scope', 0)
out_dup       = status_cnt.get('out-duplicate', 0)
replaced      = status_cnt.get('replaced', 0)
in_clarif     = status_cnt.get('in clarification', 0)

active        = done + in_clarif  # actionable items

purch_cnt     = collections.Counter(r['purch_status'].strip() for r in records if r['purch_status'])
ready_proc    = purch_cnt.get('ready for procurement', 0)
freeware      = purch_cnt.get('freeware', 0)
oss           = purch_cnt.get('Open Source Software (OSS)', 0)
no_proc       = purch_cnt.get('No procurment needed', 0)
bu_proc       = purch_cnt.get('BU procurement', 0)
sws_proc      = purch_cnt.get('SWS procurement', 0)
tsa           = purch_cnt.get('TSA', 0)

sws_cnt       = collections.Counter(r['sws'] for r in records if r['sws'])

bdbu_raw      = collections.Counter(r['bdbu'].strip().upper() for r in records if r['bdbu'])
bd_cnt        = bdbu_raw.get('BD', 0)
bu_cnt        = bdbu_raw.get('BU', 0)

pt_cnt        = collections.Counter((r['pay_type'] or 'blank').lower() for r in records)
sub_cnt       = sum(v for k, v in pt_cnt.items() if 'sub' in k)
perp_cnt      = sum(v for k, v in pt_cnt.items() if 'perp' in k)
free_cnt      = sum(v for k, v in pt_cnt.items() if 'free' in k or 'oss' in k)
tbd_cnt       = sum(v for k, v in pt_cnt.items() if 'tbd' in k or 'blank' in k or 'unk' in k)

unique_vendors = len(set(r['vendor'].lower() for r in records if r['vendor']))

with_req_date  = sum(1 for r in records if r['req_date'])
with_contact   = sum(1 for r in records if r['sw_contact'])
with_notes     = sum(1 for r in records if r['notes'])

vendor_cnt     = collections.Counter(r['vendor'] for r in records if r['vendor'])
top10          = vendor_cnt.most_common(10)

# Procurement completion rate (done / active items)
proc_pct       = round(100 * done / active) if active else 0

# Items needing action
needs_action   = in_clarif

# ── SVG helpers ──────────────────────────────────────────────────────────────
def donut_svg(parts, size=120, stroke=22):
    total_v = sum(p[0] for p in parts)
    if not total_v:
        return ''
    cx = cy = size / 2
    r = (size - stroke) / 2
    circ = 2 * 3.14159 * r
    angle = 0
    slices = []
    for val, color in parts:
        pct = val / total_v
        dash = pct * circ
        slices.append(
            f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="{color}" '
            f'stroke-width="{stroke}" stroke-dasharray="{dash:.1f} {circ:.1f}" '
            f'stroke-dashoffset="-{angle * circ / 360:.1f}" transform="rotate(-90 {cx} {cy})"/>'
        )
        angle += pct * 360
    return f'<svg width="{size}" height="{size}" viewBox="0 0 {size} {size}">{"".join(slices)}</svg>'

# ── Chart data ────────────────────────────────────────────────────────────────
status_donut = donut_svg([
    (done,      '#00783c'),
    (in_clarif, '#e65100'),
    (out_scope, '#9e9e9e'),
    (out_dup,   '#bdbdbd'),
    (replaced,  '#e0e0e0'),
])

purch_donut = donut_svg([
    (ready_proc, '#003c64'),
    (freeware,   '#00783c'),
    (oss,        '#6c757d'),
    (no_proc,    '#9e9e9e'),
    (bu_proc,    '#e65100'),
    (sws_proc,   '#dc0000'),
    (tsa,        '#795548'),
])

paytype_donut = donut_svg([
    (sub_cnt,  '#003c64'),
    (perp_cnt, '#00783c'),
    (free_cnt, '#6c757d'),
    (tbd_cnt,  '#dc0000'),
])

sws_colors = ['#003c64', '#00783c', '#e65100', '#dc0000', '#6c757d', '#795548']
sws_donut_parts = [(v, sws_colors[i % len(sws_colors)]) for i, (k, v) in enumerate(sws_cnt.most_common())]
sws_donut = donut_svg(sws_donut_parts)

# ── Top vendor bars ──────────────────────────────────────────────────────────
max_v = top10[0][1] if top10 else 1
top10_bars = ''
for vendor, cnt in top10:
    w = round(100 * cnt / max_v)
    top10_bars += f'''
    <div class="hbar-row">
      <span class="hbar-label">{vendor[:24]}</span>
      <div class="hbar-track"><div class="hbar-fill" style="width:{w}%"></div></div>
      <span class="hbar-val">{cnt}</span>
    </div>'''

# ── SWS legend rows ───────────────────────────────────────────────────────────
sws_legend = ''
for i, (sws_key, sws_val) in enumerate(sws_cnt.most_common()):
    color = sws_colors[i % len(sws_colors)]
    pct_v = round(100 * sws_val / total)
    sws_legend += f'''
    <div class="legend-item">
      <span class="leg-dot" style="background:{color}"></span>
      {sws_key}<span class="leg-pct">{sws_val} ({pct_v}%)</span>
    </div>'''

# ── Procurement completeness progress ────────────────────────────────────────
def prog_row(label, val, total_v, color='#003c64', warn=False):
    pct = round(100 * val / total_v) if total_v else 0
    style = f'color:{color}' if warn else ''
    return f'''
    <div class="prog-row">
      <span class="prog-label" style="{style}">{label}</span>
      <div class="prog-track"><div class="prog-fill" style="width:{pct}%;background:{color}"></div></div>
      <span class="prog-val" style="{style}">{pct}%</span>
    </div>'''

# ── HTML ─────────────────────────────────────────────────────────────────────
html = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Astra Software License – Management Dashboard</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:"Segoe UI",Arial,sans-serif;background:#f4f5f7;color:#1a1a1a}}

  /* HEADER */
  .header{{background:#1a1a1a;display:flex;flex-direction:column}}
  .header-top{{display:flex;align-items:center;justify-content:space-between;padding:14px 36px;background:#1a1a1a}}
  .header-top img.logo{{height:36px}}
  .header-top .title-block{{text-align:right}}
  .header-top .title-block h1{{color:#fff;font-size:1.35rem;font-weight:600;letter-spacing:.5px}}
  .header-top .title-block p{{color:#bbb;font-size:.78rem;margin-top:3px}}
  .theme-banner{{width:100%;height:60px;overflow:hidden;background:#1a1a1a}}
  .theme-banner img{{width:100%;height:60px;object-fit:cover;object-position:center;opacity:.5}}
  .theme-strip{{width:100%;height:10px;background:linear-gradient(to right,#dc0000 0%,#dc0000 33%,#003c64 33%,#003c64 66%,#00783c 66%,#00783c 100%)}}

  /* LAYOUT */
  .page{{max-width:1300px;margin:0 auto;padding:28px 28px 60px}}

  /* SECTION TITLE */
  .section-title{{display:flex;align-items:center;gap:10px;margin:34px 0 14px}}
  .sticon{{width:28px;height:28px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:13px;flex-shrink:0}}
  .section-title h2{{font-size:1rem;font-weight:700;color:#1a1a1a;text-transform:uppercase;letter-spacing:.6px}}

  /* KPI CARDS */
  .kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:14px;margin-bottom:8px}}
  .kpi{{background:#fff;border-radius:8px;padding:18px 16px;border-top:4px solid #dc0000;box-shadow:0 1px 4px rgba(0,0,0,.08);text-align:center}}
  .kpi.blue{{border-top-color:#003c64}}
  .kpi.green{{border-top-color:#00783c}}
  .kpi.gray{{border-top-color:#6c757d}}
  .kpi.orange{{border-top-color:#e65100}}
  .kpi-val{{font-size:2.2rem;font-weight:700;color:#1a1a1a;line-height:1}}
  .kpi-lbl{{font-size:.72rem;color:#555;margin-top:5px;line-height:1.4}}
  .kpi-sub{{font-size:.67rem;color:#999;margin-top:3px}}

  /* GRID */
  .two-col{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px}}
  .three-col{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;margin-bottom:16px}}
  @media(max-width:900px){{.two-col,.three-col{{grid-template-columns:1fr}}}}

  /* CARD */
  .card{{background:#fff;border-radius:8px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,.07)}}
  .card h3{{font-size:.84rem;font-weight:700;color:#1a1a1a;margin-bottom:14px;text-transform:uppercase;letter-spacing:.4px;border-bottom:2px solid #f0f0f0;padding-bottom:8px}}

  /* DONUT */
  .donut-wrap{{display:flex;align-items:center;gap:20px}}
  .donut-wrap svg{{flex-shrink:0}}
  .legend{{flex:1}}
  .legend-item{{display:flex;align-items:center;gap:8px;margin:5px 0;font-size:.79rem}}
  .leg-dot{{width:11px;height:11px;border-radius:50%;flex-shrink:0}}
  .leg-pct{{margin-left:auto;font-weight:600;color:#333}}

  /* HORIZONTAL BARS */
  .hbar-row{{display:flex;align-items:center;gap:8px;margin:5px 0;font-size:.8rem}}
  .hbar-label{{width:140px;flex-shrink:0;color:#333;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}}
  .hbar-track{{flex:1;background:#f0f0f0;border-radius:4px;height:14px;overflow:hidden}}
  .hbar-fill{{height:14px;background:#003c64;border-radius:4px}}
  .hbar-val{{width:32px;text-align:right;color:#555;font-weight:600}}

  /* PROGRESS BARS */
  .prog-row{{display:flex;align-items:center;gap:10px;margin:7px 0;font-size:.81rem}}
  .prog-label{{width:200px;flex-shrink:0;color:#444}}
  .prog-track{{flex:1;background:#f0f0f0;border-radius:6px;height:16px;overflow:hidden}}
  .prog-fill{{height:16px;border-radius:6px}}
  .prog-val{{width:46px;text-align:right;font-weight:700;color:#333}}

  /* TABLE */
  table{{border-collapse:collapse;width:100%;font-size:.81rem}}
  th{{background:#1a1a1a;color:#fff;padding:8px 10px;text-align:left;font-weight:600;font-size:.77rem;text-transform:uppercase;letter-spacing:.4px}}
  td{{padding:7px 10px;border-bottom:1px solid #eee;vertical-align:top}}
  tr:last-child td{{border-bottom:none}}
  tr:hover td{{background:#f9f9f9}}

  /* CHIPS */
  .chip{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:.7rem;font-weight:600}}
  .chip-green{{background:#e6f4ea;color:#1e6e2f}}
  .chip-red{{background:#fff0f0;color:#c62828}}
  .chip-orange{{background:#fff3e0;color:#7a4100}}
  .chip-blue{{background:#e3f0fb;color:#0d3c6e}}
  .chip-gray{{background:#f5f5f5;color:#555}}

  /* BADGE */
  .badge{{display:inline-block;padding:2px 9px;border-radius:10px;font-size:.7rem;font-weight:700}}
  .badge-sws1{{background:#003c64;color:#fff}}
  .badge-sws4{{background:#e65100;color:#fff}}
  .badge-sws3{{background:#00783c;color:#fff}}
  .badge-sws2{{background:#dc0000;color:#fff}}

  /* SUMMARY BOX */
  .summary-box{{background:#fff;border-left:5px solid #003c64;border-radius:0 8px 8px 0;padding:14px 18px;margin:8px 0;font-size:.83rem;line-height:1.7;color:#333}}
  .summary-box strong{{color:#1a1a1a}}

  /* ACTION ITEMS */
  .action-card{{background:#fff;border-radius:8px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,.07);border-top:4px solid}}
  .action-card ul{{padding-left:16px;line-height:1.9;font-size:.82rem;color:#333}}
  .action-card h3{{font-size:.85rem;font-weight:700;margin-bottom:10px;text-transform:uppercase;letter-spacing:.4px}}

  /* FOOTER */
  .footer{{text-align:center;font-size:.71rem;color:#999;margin-top:44px;padding-top:16px;border-top:1px solid #ddd}}
</style>
</head>
<body>

<!-- HEADER -->
<div class="header">
  <div class="header-top">
    <img class="logo" src="{LOGO_SRC}" alt="Bosch">
    <div class="title-block">
      <h1>Astra Software License – Management Dashboard</h1>
      <p>BDMIL Use Case &nbsp;·&nbsp; Shopping List with SWS &nbsp;·&nbsp; April 2, 2026</p>
    </div>
  </div>
  <div class="theme-banner"><img src="{THEME_SRC}" alt="Bosch theme"></div>
  <div class="theme-strip"></div>
</div>

<div class="page">

<!-- SECTION 1 – EXECUTIVE KPIs -->
<div class="section-title">
  <div class="sticon" style="background:#dc0000;color:#fff">📊</div>
  <h2>Executive Overview</h2>
</div>

<div class="kpi-grid">
  <div class="kpi blue">
    <div class="kpi-val">{total}</div>
    <div class="kpi-lbl">Total Software<br>Products in Scope</div>
    <div class="kpi-sub">Astra Shopping List</div>
  </div>
  <div class="kpi green">
    <div class="kpi-val" style="color:#00783c">{done}</div>
    <div class="kpi-lbl">Procurement<br>Completed (Done)</div>
    <div class="kpi-sub">{round(100*done/total)}% of total</div>
  </div>
  <div class="kpi orange">
    <div class="kpi-val" style="color:#e65100">{in_clarif}</div>
    <div class="kpi-lbl">Items in<br>Clarification</div>
    <div class="kpi-sub">Require immediate attention</div>
  </div>
  <div class="kpi">
    <div class="kpi-val" style="color:#dc0000">{out_scope + out_dup + replaced}</div>
    <div class="kpi-lbl">Out of Scope /<br>Duplicate / Replaced</div>
    <div class="kpi-sub">{out_scope} OOS · {out_dup} dup · {replaced} replaced</div>
  </div>
  <div class="kpi blue">
    <div class="kpi-val">{unique_vendors}</div>
    <div class="kpi-lbl">Unique Vendors<br>Identified</div>
    <div class="kpi-sub">Across all SWS</div>
  </div>
  <div class="kpi green">
    <div class="kpi-val" style="color:#00783c">{ready_proc}</div>
    <div class="kpi-lbl">Ready for<br>Procurement</div>
    <div class="kpi-sub">{round(100*ready_proc/total)}% of total</div>
  </div>
  <div class="kpi gray">
    <div class="kpi-val">{freeware + oss}</div>
    <div class="kpi-lbl">Freeware / OSS<br>Products Tracked</div>
    <div class="kpi-sub">{freeware} freeware · {oss} OSS</div>
  </div>
  <div class="kpi">
    <div class="kpi-val" style="color:#003c64">{with_req_date}</div>
    <div class="kpi-lbl">License Required<br>Date Defined</div>
    <div class="kpi-sub">{round(100*with_req_date/total)}% of total</div>
  </div>
</div>

<!-- Summary narrative -->
<div class="summary-box" style="margin:14px 0 0">
  <strong>Summary:</strong> The Astra software license portfolio covers <strong>{total} products</strong> from 
  <strong>{unique_vendors} vendors</strong> distributed across 6 SWS systems (SWS1–SWS8). 
  <strong>{done} items ({round(100*done/total)}%)</strong> have completed procurement.
  <strong>{ready_proc} items</strong> are ready for procurement and pending action.
  <strong>{in_clarif} items</strong> remain in clarification and require resolution.
  <strong>{freeware + oss} products</strong> are classified as Freeware or Open Source Software.
</div>

<!-- SECTION 2 – STATUS & PROCUREMENT -->
<div class="section-title">
  <div class="sticon" style="background:#003c64;color:#fff">📋</div>
  <h2>Status &amp; Procurement Overview</h2>
</div>

<div class="three-col">

  <!-- Status donut -->
  <div class="card">
    <h3>Item Status Distribution</h3>
    <div class="donut-wrap">
      {status_donut}
      <div class="legend">
        <div class="legend-item"><span class="leg-dot" style="background:#00783c"></span>Done<span class="leg-pct">{done} ({round(100*done/total)}%)</span></div>
        <div class="legend-item"><span class="leg-dot" style="background:#e65100"></span>In Clarification<span class="leg-pct">{in_clarif} ({round(100*in_clarif/total)}%)</span></div>
        <div class="legend-item"><span class="leg-dot" style="background:#9e9e9e"></span>Out of Scope<span class="leg-pct">{out_scope} ({round(100*out_scope/total)}%)</span></div>
        <div class="legend-item"><span class="leg-dot" style="background:#bdbdbd"></span>Out-Duplicate<span class="leg-pct">{out_dup} ({round(100*out_dup/total)}%)</span></div>
        <div class="legend-item"><span class="leg-dot" style="background:#e0e0e0"></span>Replaced<span class="leg-pct">{replaced} ({round(100*replaced/total)}%)</span></div>
      </div>
    </div>
    <div style="margin-top:14px">
      {prog_row('Completed (Done)', done, total, '#00783c')}
      {prog_row('Actionable (Done + Clarif)', active, total, '#003c64')}
      {prog_row('⚠ In Clarification', in_clarif, total, '#e65100', warn=True)}
      {prog_row('Excluded from scope', out_scope + out_dup + replaced, total, '#9e9e9e')}
    </div>
  </div>

  <!-- Purchasing status donut -->
  <div class="card">
    <h3>Purchasing Status Breakdown</h3>
    <div class="donut-wrap">
      {purch_donut}
      <div class="legend">
        <div class="legend-item"><span class="leg-dot" style="background:#003c64"></span>Ready for Proc.<span class="leg-pct">{ready_proc}</span></div>
        <div class="legend-item"><span class="leg-dot" style="background:#00783c"></span>Freeware<span class="leg-pct">{freeware}</span></div>
        <div class="legend-item"><span class="leg-dot" style="background:#6c757d"></span>OSS<span class="leg-pct">{oss}</span></div>
        <div class="legend-item"><span class="leg-dot" style="background:#9e9e9e"></span>No Proc. Needed<span class="leg-pct">{no_proc}</span></div>
        <div class="legend-item"><span class="leg-dot" style="background:#e65100"></span>BU Procurement<span class="leg-pct">{bu_proc}</span></div>
        <div class="legend-item"><span class="leg-dot" style="background:#dc0000"></span>SWS Procurement<span class="leg-pct">{sws_proc}</span></div>
        <div class="legend-item"><span class="leg-dot" style="background:#795548"></span>TSA<span class="leg-pct">{tsa}</span></div>
      </div>
    </div>
  </div>

  <!-- Payment type donut -->
  <div class="card">
    <h3>Payment / License Model</h3>
    <div class="donut-wrap">
      {paytype_donut}
      <div class="legend">
        <div class="legend-item"><span class="leg-dot" style="background:#003c64"></span>Subscription<span class="leg-pct">{sub_cnt}</span></div>
        <div class="legend-item"><span class="leg-dot" style="background:#00783c"></span>Perpetual<span class="leg-pct">{perp_cnt}</span></div>
        <div class="legend-item"><span class="leg-dot" style="background:#6c757d"></span>Free / OSS<span class="leg-pct">{free_cnt}</span></div>
        <div class="legend-item"><span class="leg-dot" style="background:#dc0000"></span>TBD / Not Set<span class="leg-pct">{tbd_cnt}</span></div>
      </div>
    </div>
    <div style="margin-top:14px">
      {prog_row('Subscription', sub_cnt, total, '#003c64')}
      {prog_row('Perpetual', perp_cnt, total, '#00783c')}
      {prog_row('Free / OSS', free_cnt, total, '#6c757d')}
      {prog_row('⚠ TBD / Not defined', tbd_cnt, total, '#dc0000', warn=True)}
    </div>
  </div>

</div>

<!-- SECTION 3 – SWS DISTRIBUTION -->
<div class="section-title">
  <div class="sticon" style="background:#00783c;color:#fff">🖥</div>
  <h2>SWS Distribution</h2>
</div>

<div class="two-col">

  <div class="card">
    <h3>Products per SWS System</h3>
    <div class="donut-wrap">
      {sws_donut}
      <div class="legend">
        {sws_legend}
      </div>
    </div>
  </div>

  <div class="card">
    <h3>SWS Procurement Progress</h3>
    <div style="font-size:.79rem;color:#888;margin-bottom:12px">Items with status = Done, per SWS</div>
    {''.join(
        prog_row(
            sws_key,
            sum(1 for r in records if r['sws'] == sws_key and r['status'].lower().strip() == 'done'),
            sws_val,
            sws_colors[i % len(sws_colors)]
        )
        for i, (sws_key, sws_val) in enumerate(sws_cnt.most_common())
    )}
    <div style="margin-top:14px;font-size:.77rem;color:#888;border-top:1px solid #eee;padding-top:10px">
      <strong>BD vs BU Split:</strong>
      &nbsp;<span class="chip chip-blue">BD: {bd_cnt} products ({round(100*bd_cnt/total)}%)</span>
      &nbsp;<span class="chip chip-orange">BU: {bu_cnt} products ({round(100*bu_cnt/total)}%)</span>
    </div>
  </div>

</div>

<!-- SWS detail table -->
<div class="card" style="margin-bottom:16px">
  <h3>SWS Summary Table</h3>
  <table>
    <thead>
      <tr>
        <th>SWS</th>
        <th>Total Items</th>
        <th>Done</th>
        <th>In Clarification</th>
        <th>Out of Scope / Dup</th>
        <th>Ready for Proc.</th>
        <th>Completion %</th>
      </tr>
    </thead>
    <tbody>
      {''.join(f"""
      <tr>
        <td><strong>{sws_key}</strong></td>
        <td>{sws_val}</td>
        <td><span class="chip chip-green">{sum(1 for r in records if r['sws']==sws_key and r['status'].lower().strip()=='done')}</span></td>
        <td>{"<span class='chip chip-orange'>"+str(sum(1 for r in records if r['sws']==sws_key and 'clarif' in r['status'].lower()))+"</span>" if sum(1 for r in records if r['sws']==sws_key and 'clarif' in r['status'].lower()) else "<span class='chip chip-gray'>0</span>"}</td>
        <td>{sum(1 for r in records if r['sws']==sws_key and ('scope' in r['status'].lower() or 'duplic' in r['status'].lower() or 'replac' in r['status'].lower()))}</td>
        <td>{sum(1 for r in records if r['sws']==sws_key and r['purch_status']=='ready for procurement')}</td>
        <td>{"<span class='chip chip-green'>"+str(round(100*sum(1 for r in records if r['sws']==sws_key and r['status'].lower().strip()=='done')/sws_val))+"%</span>" if sws_val else "–"}</td>
      </tr>""" for sws_key, sws_val in sws_cnt.most_common())}
      <tr style="background:#f8f8f8;font-weight:700">
        <td>TOTAL</td>
        <td>{total}</td>
        <td><span class="chip chip-green">{done}</span></td>
        <td><span class="chip chip-orange">{in_clarif}</span></td>
        <td>{out_scope + out_dup + replaced}</td>
        <td>{ready_proc}</td>
        <td><span class="chip chip-green">{round(100*done/total)}%</span></td>
      </tr>
    </tbody>
  </table>
</div>

<!-- SECTION 4 – VENDOR LANDSCAPE -->
<div class="section-title">
  <div class="sticon" style="background:#e65100;color:#fff">🏢</div>
  <h2>Vendor Landscape</h2>
</div>

<div class="two-col">

  <div class="card">
    <h3>Top 10 Vendors by Product Count</h3>
    {top10_bars}
  </div>

  <div class="card">
    <h3>Data Completeness</h3>
    <div style="font-size:.79rem;color:#888;margin-bottom:12px">How well key fields are populated across all {total} items</div>
    {prog_row('Status defined', sum(1 for r in records if r['status']), total, '#003c64')}
    {prog_row('Purchasing status set', sum(1 for r in records if r['purch_status']), total, '#003c64')}
    {prog_row('License date defined', with_req_date, total, '#00783c')}
    {prog_row('SW Key Contact assigned', with_contact, total, '#00783c')}
    {prog_row('Notes / comments filled', with_notes, total, '#6c757d')}
    {prog_row('Payment type set', sum(1 for r in records if r['pay_type']), total, '#6c757d')}
    {prog_row('⚠ TBD / blank payment type', tbd_cnt, total, '#dc0000', warn=True)}
  </div>

</div>

<!-- SECTION 5 – NEXT STEPS / RECOMMENDATIONS -->
<div class="section-title">
  <div class="sticon" style="background:#003c64;color:#fff">✅</div>
  <h2>Recommended Next Steps</h2>
</div>

<div class="three-col">
  <div class="action-card" style="border-top-color:#dc0000">
    <h3 style="color:#dc0000">🔴 Immediate Actions</h3>
    <ul>
      <li>Resolve <strong>{in_clarif} items in clarification</strong> — assign owners &amp; target dates</li>
      <li>Classify <strong>{tbd_cnt} items with blank/TBD payment type</strong> before procurement cutover</li>
      <li>Progress <strong>{ready_proc} items</strong> marked "ready for procurement" to active ordering</li>
    </ul>
  </div>
  <div class="action-card" style="border-top-color:#e65100">
    <h3 style="color:#e65100">🟠 Short-term Actions</h3>
    <ul>
      <li>Populate <strong>SW Key Contact</strong> for {total - with_contact} items currently unassigned ({round(100*(total-with_contact)/total)}%)</li>
      <li>Set <strong>Required License Date</strong> for {total - with_req_date} items without a target date</li>
      <li>Review <strong>{bu_proc} BU procurement items</strong> to confirm ownership &amp; timelines</li>
    </ul>
  </div>
  <div class="action-card" style="border-top-color:#00783c">
    <h3 style="color:#00783c">🟡 Ongoing / Governance</h3>
    <ul>
      <li>Establish quarterly review for <strong>{out_scope} out-of-scope</strong> and <strong>{replaced} replaced</strong> items</li>
      <li>Audit <strong>{freeware + oss} Freeware / OSS products</strong> for compliance obligations (GPL, MIT etc.)</li>
      <li>Consolidate <strong>{out_dup} duplicate entries</strong> and align master list across SWS teams</li>
    </ul>
  </div>
</div>

<!-- FOOTER -->
<div class="footer">
  Astra Software License Management Dashboard &nbsp;·&nbsp; BDMIL AI Use Case &nbsp;·&nbsp; April 2, 2026
  <br>Generated by GitHub Copilot &nbsp;|&nbsp; Data source: Astra Shoppping List with SWS.xlsx
</div>

</div><!-- end .page -->
</body>
</html>'''

out = os.path.join(BASE, 'Astra_SWS_Management_Dashboard.html')
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)
print(f"Written: {out}  ({len(html)//1024} KB)")
