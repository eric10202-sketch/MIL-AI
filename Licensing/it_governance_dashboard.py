"""
IT Governance Dashboard Generator — JV dBI
==========================================
Generates a self-contained HTML management dashboard with KPIs and charts.
Designed to be scheduled weekly via Power Automate (Run Script / Run Python).

REQUIREMENTS (install once):
    pip install pyodbc pandas plotly

CONFIGURATION:
    Edit the CONFIG block below before first run.
"""

import base64
import os
import sys
import traceback
from datetime import datetime
from pathlib import Path

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION — edit these values to match your environment
# ─────────────────────────────────────────────────────────────────────────────
CONFIG = {
    # SQL Server connection string (pyodbc format)
    "sql_connection_string": (
        "DRIVER={SQL Server};"
        "SERVER=rb-cimi-mr-sql-q.de.bosch.com;"
        "DATABASE=DB_MIC_MIGRATIONS_Q_SQL;"
        "UID=MIC_MIGRATIONS_Q_R_PBI_SUAAS;"
        "PWD=iRZrsEJCvEirqzLX0GSi$;"
    ),

    # Output folder — dashboard HTML will be saved here
    "output_folder": str(Path(__file__).parent),

    # Bosch logo path (relative to this script, or absolute)
    "logo_path": str(Path(__file__).parent / "Bosch.png"),

    # Company/scope label shown in the header
    "report_scope": "JV dBI",
}

# SQL view names — schema.view format
VIEWS = {
    "users":       "[Reporting].[V_BaseData_User_AD_JV]",
    "roles":       "[Reporting].[V_Basedata_User_IdM_Authorizations_RolesRelations_JV]",
    "computers":   "[Reporting].[V_CompMig_Computer_Assets_JV]",
}

# ─────────────────────────────────────────────────────────────────────────────
# DATA LAYER
# ─────────────────────────────────────────────────────────────────────────────

def get_connection():
    try:
        import pyodbc
        conn = pyodbc.connect(CONFIG["sql_connection_string"], timeout=30)
        return conn
    except Exception as e:
        print(f"[WARN] Database connection failed: {e}")
        return None


def query(conn, sql):
    """Run a SQL query and return a pandas DataFrame, or None on failure."""
    try:
        import pandas as pd
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            return pd.read_sql(sql, conn)
    except Exception as e:
        print(f"[WARN] Query failed: {e}")
        return None


def fetch_all_data(conn):
    """Fetch all metrics and chart data. Returns a dict of DataFrames."""
    v = VIEWS
    data = {}

    # ── KPI scalars ──────────────────────────────────────────────────────────
    data["kpi_jv_accounts"] = query(
        conn, f"SELECT COUNT(*) AS val FROM {v['users']}"
    )
    data["kpi_mfa"] = query(
        conn,
        f"SELECT COUNT(*) AS val FROM {v['users']} WHERE MFA = 1"
    )
    data["kpi_idm_roles"] = query(
        conn,
        f"SELECT COUNT(DISTINCT [Role name]) AS val FROM {v['roles']}"
    )
    # Roles with at least one risk flag set
    data["kpi_violated_roles"] = query(
        conn,
        f"""SELECT COUNT(DISTINCT [Role name]) AS val FROM {v['roles']}
            WHERE [Role Risks ADA]=1 OR [Role Risks CAA]=1 OR [Role Risks CCA]=1
                  OR [Role Risks SEA]=1 OR [Role Risks CUA]=1"""
    )
    # Applications that have at least one risky role assigned
    data["kpi_violated_apps"] = query(
        conn,
        f"""SELECT COUNT(DISTINCT Application) AS val FROM {v['roles']}
            WHERE [Role Risks ADA]=1 OR [Role Risks CAA]=1 OR [Role Risks CCA]=1
                  OR [Role Risks SEA]=1 OR [Role Risks CUA]=1"""
    )
    data["kpi_vms"] = query(
        conn,
        f"SELECT COUNT(*) AS val FROM {v['computers']}"
    )

    # ── Chart data ───────────────────────────────────────────────────────────
    data["by_company"] = query(
        conn,
        f"SELECT company AS label, COUNT(*) AS value FROM {v['users']} GROUP BY company ORDER BY value DESC"
    )
    data["by_user_type"] = query(
        conn,
        f"SELECT EmployeeType AS label, COUNT(*) AS value FROM {v['users']} GROUP BY EmployeeType ORDER BY value DESC"
    )
    data["by_department"] = query(
        conn,
        f"SELECT TOP 15 department AS label, COUNT(*) AS value FROM {v['users']} GROUP BY department ORDER BY value DESC"
    )
    data["by_device_type"] = query(
        conn,
        f"SELECT Tier_3 AS label, COUNT(*) AS value FROM {v['computers']} GROUP BY Tier_3 ORDER BY value DESC"
    )
    # Risk assignments by company (users with at least one risky role)
    data["violation_by_company"] = query(
        conn,
        f"""SELECT u.company AS label,
               CASE
                 WHEN r.[Role Risks ADA]=1 THEN 'ADA'
                 WHEN r.[Role Risks CAA]=1 THEN 'CAA'
                 WHEN r.[Role Risks CCA]=1 THEN 'CCA'
                 WHEN r.[Role Risks SEA]=1 THEN 'SEA'
                 WHEN r.[Role Risks CUA]=1 THEN 'CUA'
               END AS vtype,
               COUNT(*) AS value
            FROM {v['roles']} r
            LEFT JOIN {v['users']} u ON r.UserID = u.sAMAccountName
            WHERE r.[Role Risks ADA]=1 OR r.[Role Risks CAA]=1 OR r.[Role Risks CCA]=1
                  OR r.[Role Risks SEA]=1 OR r.[Role Risks CUA]=1
            GROUP BY u.company,
               CASE
                 WHEN r.[Role Risks ADA]=1 THEN 'ADA'
                 WHEN r.[Role Risks CAA]=1 THEN 'CAA'
                 WHEN r.[Role Risks CCA]=1 THEN 'CCA'
                 WHEN r.[Role Risks SEA]=1 THEN 'SEA'
                 WHEN r.[Role Risks CUA]=1 THEN 'CUA'
               END"""
    )
    data["idm_by_risk"] = query(
        conn,
        f"""SELECT
              SUM(CAST([Role Risks ADA] AS INT)) AS ADA,
              SUM(CAST([Role Risks CAA] AS INT)) AS CAA,
              SUM(CAST([Role Risks CCA] AS INT)) AS CCA,
              SUM(CAST([Role Risks SEA] AS INT)) AS SEA,
              SUM(CAST([Role Risks CUA] AS INT)) AS CUA
           FROM {v['roles']}"""
    )
    data["top10_idm_roles"] = query(
        conn,
        f"""SELECT TOP 10 [Role name] AS label, COUNT(*) AS value
            FROM {v['roles']}
            GROUP BY [Role name]
            ORDER BY value DESC"""
    )
    data["mfa_rate_company"] = query(
        conn,
        f"""SELECT company AS label,
              COUNT(*) AS total,
              SUM(CASE WHEN MFA = 1 THEN 1 ELSE 0 END) AS mfa_assigned
           FROM {v['users']}
           GROUP BY company"""
    )

    return data


def safe_kpi(df, col="val"):
    """Extract a single integer KPI value from a 1-row DataFrame."""
    try:
        if df is not None and not df.empty:
            return int(df.iloc[0][col])
    except Exception:
        pass
    return None


# ─────────────────────────────────────────────────────────────────────────────
# CHART BUILDERS  (return plotly div HTML strings)
# ─────────────────────────────────────────────────────────────────────────────

BOSCH_COLORS = [
    "#007BC0",  # Bosch Blue
    "#00A0E3",
    "#0055A0",
    "#E20015",  # Bosch Red
    "#FF6B35",
    "#47B8AF",
    "#8DC63F",
    "#F2A900",
    "#6D2077",
    "#003087",
]

PLOTLY_LAYOUT = dict(
    font=dict(family="Bosch Sans, Segoe UI, Arial", size=12),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    margin=dict(l=10, r=10, t=40, b=60),
    legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
    colorway=BOSCH_COLORS,
    height=300,
)


def make_bar_chart(df, x_col, y_col, title, orientation="v", color="#007BC0"):
    try:
        import plotly.graph_objects as go
        if df is None or df.empty:
            return _empty_chart(title)
        df = df.dropna(subset=[x_col, y_col])
        if orientation == "h":
            fig = go.Figure(go.Bar(
                x=df[y_col], y=df[x_col], orientation="h",
                marker_color=color, text=df[y_col], textposition="outside"
            ))
        else:
            fig = go.Figure(go.Bar(
                x=df[x_col], y=df[y_col], orientation="v",
                marker_color=color, text=df[y_col], textposition="outside"
            ))
        fig.update_layout(title_text=title, title_font_size=13, **PLOTLY_LAYOUT)
        if orientation == "h":
            fig.update_xaxes(showgrid=False)
            fig.update_yaxes(tickfont_size=10)
        else:
            fig.update_yaxes(showgrid=True, gridcolor="#e8e8e8")
            fig.update_xaxes(tickfont_size=10)
        return fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})
    except Exception as e:
        return f'<div class="chart-error">Chart error: {e}</div>'


def make_grouped_bar(df, group_col, category_col, value_col, title):
    try:
        import plotly.graph_objects as go
        if df is None or df.empty:
            return _empty_chart(title)
        categories = df[category_col].dropna().unique()
        fig = go.Figure()
        for i, cat in enumerate(categories):
            sub = df[df[category_col] == cat]
            fig.add_trace(go.Bar(
                name=str(cat),
                x=sub[group_col],
                y=sub[value_col],
                marker_color=BOSCH_COLORS[i % len(BOSCH_COLORS)],
            ))
        fig.update_layout(title_text=title, title_font_size=13,
                          barmode="group", **PLOTLY_LAYOUT)
        fig.update_yaxes(showgrid=True, gridcolor="#e8e8e8")
        return fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})
    except Exception as e:
        return f'<div class="chart-error">Chart error: {e}</div>'


def make_risk_bar(df, title):
    try:
        import plotly.graph_objects as go
        if df is None or df.empty:
            return _empty_chart(title)
        row = df.iloc[0]
        labels = ["ADA", "CAA", "CCA", "SEA", "CUA"]
        values = [int(row.get(l, 0) or 0) for l in labels]
        colors = ["#E20015", "#F2A900", "#007BC0", "#47B8AF", "#6D2077"]
        fig = go.Figure(go.Bar(
            x=labels, y=values,
            marker_color=colors,
            text=values, textposition="outside"
        ))
        fig.update_layout(title_text=title, title_font_size=13, **PLOTLY_LAYOUT)
        fig.update_yaxes(showgrid=True, gridcolor="#e8e8e8")
        return fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})
    except Exception as e:
        return f'<div class="chart-error">Chart error: {e}</div>'


def make_mfa_gauge(total, mfa_count, title):
    try:
        import plotly.graph_objects as go
        if total is None or mfa_count is None or total == 0:
            return _empty_chart(title)
        pct = round(mfa_count / total * 100, 1)
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=pct,
            number={"suffix": "%", "font": {"size": 36}},
            delta={"reference": 100, "decreasing": {"color": "#E20015"}},
            gauge={
                "axis": {"range": [0, 100], "tickcolor": "#666"},
                "bar": {"color": "#007BC0"},
                "bgcolor": "white",
                "steps": [
                    {"range": [0, 60], "color": "#fdecea"},
                    {"range": [60, 80], "color": "#fff8e1"},
                    {"range": [80, 100], "color": "#e8f5e9"},
                ],
                "threshold": {"line": {"color": "#E20015", "width": 3}, "value": 80},
            },
            title={"text": title, "font": {"size": 13}},
        ))
        fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", margin=dict(l=20, r=20, t=40, b=10),
                          height=220)
        return fig.to_html(full_html=False, include_plotlyjs=False, config={"displayModeBar": False})
    except Exception as e:
        return f'<div class="chart-error">Chart error: {e}</div>'


def _empty_chart(title):
    return f'<div class="chart-empty"><span>No data available</span><br><small>{title}</small></div>'


# ─────────────────────────────────────────────────────────────────────────────
# HTML TEMPLATE
# ─────────────────────────────────────────────────────────────────────────────

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>IT Governance Dashboard — {scope}</title>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<style>
  :root {{
    --bosch-blue: #007BC0;
    --bosch-dark: #003054;
    --bosch-red: #E20015;
    --bosch-light: #f4f6f9;
    --card-bg: #ffffff;
    --text-main: #1a1a2e;
    --text-muted: #6c757d;
    --border: #e0e6ee;
    --shadow: 0 2px 12px rgba(0,60,100,0.08);
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: "Segoe UI", Arial, sans-serif; background: var(--bosch-light); color: var(--text-main); }}

  /* ── HEADER ── */
  .header {{
    background: linear-gradient(135deg, var(--bosch-dark) 0%, var(--bosch-blue) 100%);
    color: white; padding: 18px 32px;
    display: flex; align-items: center; justify-content: space-between;
    box-shadow: 0 3px 12px rgba(0,0,0,0.25);
  }}
  .header-left {{ display: flex; align-items: center; gap: 20px; }}
  .header-logo {{ height: 36px; filter: brightness(0) invert(1); }}
  .header-divider {{ width: 1px; height: 40px; background: rgba(255,255,255,0.35); }}
  .header-title h1 {{ font-size: 20px; font-weight: 700; letter-spacing: 0.3px; }}
  .header-title p  {{ font-size: 12px; opacity: 0.75; margin-top: 3px; }}
  .header-meta {{ text-align: right; font-size: 12px; opacity: 0.8; }}
  .header-meta strong {{ display: block; font-size: 15px; opacity: 1; }}

  /* ── MAIN LAYOUT ── */
  .main {{ max-width: 1400px; margin: 0 auto; padding: 24px 24px 40px; }}

  /* ── SECTION LABELS ── */
  .section-label {{
    font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 1px;
    color: var(--bosch-blue); border-left: 3px solid var(--bosch-blue);
    padding-left: 10px; margin: 28px 0 14px;
  }}

  /* ── KPI CARDS ── */
  .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; }}
  .kpi-card {{
    background: var(--card-bg); border-radius: 10px; padding: 20px 18px;
    box-shadow: var(--shadow); border-top: 4px solid var(--bosch-blue);
    transition: transform 0.15s; position: relative; overflow: hidden;
  }}
  .kpi-card:hover {{ transform: translateY(-2px); }}
  .kpi-card.red   {{ border-top-color: var(--bosch-red); }}
  .kpi-card.amber {{ border-top-color: #F2A900; }}
  .kpi-card.green {{ border-top-color: #47B8AF; }}
  .kpi-card.teal  {{ border-top-color: #47B8AF; }}
  .kpi-icon {{ font-size: 28px; position: absolute; top: 14px; right: 16px; opacity: 0.15; }}
  .kpi-value {{ font-size: 34px; font-weight: 800; color: var(--bosch-dark); line-height: 1; }}
  .kpi-value.na {{ font-size: 20px; color: var(--text-muted); }}
  .kpi-label {{ font-size: 12px; color: var(--text-muted); margin-top: 8px; font-weight: 500; }}
  .kpi-sub   {{ font-size: 11px; color: #aaa; margin-top: 4px; }}

  /* ── CHART GRID ── */
  .chart-grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 18px; }}
  .chart-grid-3 {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 18px; }}
  .chart-grid-1 {{ display: grid; grid-template-columns: 1fr; gap: 18px; }}
  .chart-card {{
    background: var(--card-bg); border-radius: 10px; padding: 18px 14px 10px;
    box-shadow: var(--shadow); min-height: 320px;
  }}
  .chart-card.wide {{ grid-column: span 2; }}
  .chart-empty {{
    height: 200px; display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    color: var(--text-muted); font-size: 13px;
  }}
  .chart-error {{ color: #c0392b; font-size: 11px; padding: 8px; }}

  /* ── STATUS BANNER ── */
  .status-banner {{
    background: #e8f4fd; border: 1px solid #b3d8f0; border-radius: 8px;
    padding: 12px 18px; font-size: 12px; color: #0069aa;
    margin-bottom: 18px; display: flex; align-items: center; gap: 10px;
  }}
  .status-banner.warn {{ background: #fff8e1; border-color: #ffe082; color: #856404; }}
  .status-dot {{ width: 8px; height: 8px; border-radius: 50%; background: currentColor; flex-shrink: 0; }}

  /* ── FOOTER ── */
  .footer {{
    text-align: center; font-size: 11px; color: #aaa;
    padding: 18px; border-top: 1px solid var(--border); margin-top: 40px;
  }}

  /* ── RESPONSIVE ── */
  @media (max-width: 900px) {{
    .chart-grid-2, .chart-grid-3 {{ grid-template-columns: 1fr; }}
    .chart-card.wide {{ grid-column: span 1; }}
  }}
</style>
</head>
<body>

<!-- HEADER -->
<div class="header">
  <div class="header-left">
    <img class="header-logo" src="{logo_data}" alt="Bosch"/>
    <div class="header-divider"></div>
    <div class="header-title">
      <h1>IT Governance Dashboard</h1>
      <p>Identity &amp; Access Management — {scope}</p>
    </div>
  </div>
  <div class="header-meta">
    <strong>{report_date}</strong>
    Weekly Management Report
  </div>
</div>

<div class="main">

{status_banner}

<!-- ── KPIs ─────────────────────────────────────────────────────────────── -->
<div class="section-label">Key Performance Indicators</div>
<div class="kpi-grid">

  <div class="kpi-card">
    <div class="kpi-icon">👤</div>
    <div class="kpi-value {jv_accounts_cls}">{jv_accounts}</div>
    <div class="kpi-label">JV Accounts</div>
    <div class="kpi-sub">Active Directory users</div>
  </div>

  <div class="kpi-card green">
    <div class="kpi-icon">🔐</div>
    <div class="kpi-value {mfa_cls}">{mfa_assigned}</div>
    <div class="kpi-label">MFA Assigned</div>
    <div class="kpi-sub">Multi-factor authentication</div>
  </div>

  <div class="kpi-card teal">
    <div class="kpi-icon">🎭</div>
    <div class="kpi-value {idm_cls}">{idm_roles}</div>
    <div class="kpi-label">IdM Roles</div>
    <div class="kpi-sub">Distinct roles in use</div>
  </div>

  <div class="kpi-card amber">
    <div class="kpi-icon">💻</div>
    <div class="kpi-value {vm_cls}">{vms}</div>
    <div class="kpi-label">Virtual Machines (WTS)</div>
    <div class="kpi-sub">Managed compute assets</div>
  </div>

  <div class="kpi-card red">
    <div class="kpi-icon">⚠️</div>
    <div class="kpi-value {vr_cls}">{violated_roles}</div>
    <div class="kpi-label">Risky IdM Roles</div>
    <div class="kpi-sub">Roles with risk flags (ADA/CAA/CCA/SEA/CUA)</div>
  </div>

  <div class="kpi-card red">
    <div class="kpi-icon">🚦</div>
    <div class="kpi-value {va_cls}">{violated_apps}</div>
    <div class="kpi-label">Applications with Risk</div>
    <div class="kpi-sub">Apps with at least one risky role</div>
  </div>

</div>

<!-- ── MFA GAUGE + COMPANY BREAKDOWN ─────────────────────────────────────── -->
<div class="section-label">MFA Adoption &amp; User Distribution</div>
<div class="chart-grid-3">
  <div class="chart-card">{chart_mfa_gauge}</div>
  <div class="chart-card">{chart_by_company}</div>
  <div class="chart-card">{chart_by_user_type}</div>
</div>

<!-- ── DEPARTMENT + DEVICE ───────────────────────────────────────────────── -->
<div class="section-label">Department &amp; Device Distribution</div>
<div class="chart-grid-2">
  <div class="chart-card">{chart_by_department}</div>
  <div class="chart-card">{chart_by_device}</div>
</div>

<!-- ── VIOLATIONS ────────────────────────────────────────────────────────── -->
<div class="section-label">Governance Violations</div>
<div class="chart-grid-2">
  <div class="chart-card wide">{chart_violations}</div>
</div>

<!-- ── IdM RISK + TOP 10 ROLES ───────────────────────────────────────────── -->
<div class="section-label">IdM Role Risk &amp; Assignment</div>
<div class="chart-grid-2">
  <div class="chart-card">{chart_idm_risk}</div>
  <div class="chart-card">{chart_top10_roles}</div>
</div>

</div><!-- /main -->

<div class="footer">
  Generated: {generated_ts} &nbsp;·&nbsp; IT Governance Dashboard &nbsp;·&nbsp; {scope} &nbsp;·&nbsp; Bosch Group
</div>

</body>
</html>
"""


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def fmt_kpi(val):
    """Format a KPI number with thousands separator, or 'N/A'."""
    if val is None:
        return ("N/A", "na")
    return (f"{val:,}", "")


def load_logo_base64(path):
    try:
        with open(path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        return f"data:image/png;base64,{b64}"
    except Exception:
        return ""


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    now = datetime.now()
    print(f"[{now:%Y-%m-%d %H:%M:%S}] IT Governance Dashboard Generator starting...")

    # ── Connect to database ──────────────────────────────────────────────────
    conn = get_connection()
    connected = conn is not None

    status_banner = ""
    data = {}

    if connected:
        print("[INFO] Database connected. Fetching data...")
        try:
            data = fetch_all_data(conn)
            conn.close()
            status_banner = '<div class="status-banner"><div class="status-dot"></div>Data refreshed live from SQL database — all metrics current as of {}</div>'.format(
                now.strftime("%d %b %Y %H:%M")
            )
        except Exception as e:
            traceback.print_exc()
            status_banner = f'<div class="status-banner warn"><div class="status-dot"></div>Data fetch partially failed: {e}</div>'
    else:
        status_banner = '<div class="status-banner warn"><div class="status-dot"></div>⚠ Could not connect to database. Update the SQL connection string in CONFIG. Displaying N/A placeholders.</div>'

    # ── KPI values ───────────────────────────────────────────────────────────
    jv_accounts   = safe_kpi(data.get("kpi_jv_accounts"))
    mfa_assigned  = safe_kpi(data.get("kpi_mfa"))
    idm_roles     = safe_kpi(data.get("kpi_idm_roles"))
    vms           = safe_kpi(data.get("kpi_vms"))
    violated_roles = safe_kpi(data.get("kpi_violated_roles"))
    violated_apps  = safe_kpi(data.get("kpi_violated_apps"))

    jv_v, jv_c   = fmt_kpi(jv_accounts)
    mfa_v, mfa_c = fmt_kpi(mfa_assigned)
    idm_v, idm_c = fmt_kpi(idm_roles)
    vm_v, vm_c   = fmt_kpi(vms)
    vr_v, vr_c   = fmt_kpi(violated_roles)
    va_v, va_c   = fmt_kpi(violated_apps)

    # ── Charts ───────────────────────────────────────────────────────────────
    chart_mfa_gauge  = make_mfa_gauge(jv_accounts, mfa_assigned, "MFA Adoption Rate")
    chart_by_company = make_bar_chart(data.get("by_company"),  "label", "value", "JV Accounts by Company", color="#007BC0")
    chart_by_user_type = make_bar_chart(data.get("by_user_type"), "label", "value", "JV Accounts by IT User Type", color="#00A0E3")
    chart_by_dept    = make_bar_chart(data.get("by_department"), "label", "value", "Top Departments by User Count", orientation="h", color="#47B8AF")
    chart_by_device  = make_bar_chart(data.get("by_device_type"), "label", "value", "JV Accounts by Device Type", color="#F2A900")
    chart_violations = make_grouped_bar(data.get("violation_by_company"), "label", "vtype", "value", "Governance Violation Type by Company")
    chart_idm_risk   = make_risk_bar(data.get("idm_by_risk"), "IdM Roles by Risk Category")
    chart_top10_roles = make_bar_chart(data.get("top10_idm_roles"), "label", "value", "Top 10 IdM Roles Assigned", orientation="h", color="#E20015")

    # ── Logo ─────────────────────────────────────────────────────────────────
    logo_data = load_logo_base64(CONFIG["logo_path"])

    # ── Render HTML ──────────────────────────────────────────────────────────
    html = HTML_TEMPLATE.format(
        scope          = CONFIG["report_scope"],
        report_date    = now.strftime("%d %B %Y"),
        generated_ts   = now.strftime("%Y-%m-%d %H:%M:%S"),
        logo_data      = logo_data,
        status_banner  = status_banner,

        jv_accounts    = jv_v,   jv_accounts_cls = jv_c,
        mfa_assigned   = mfa_v,  mfa_cls         = mfa_c,
        idm_roles      = idm_v,  idm_cls         = idm_c,
        vms            = vm_v,   vm_cls          = vm_c,
        violated_roles = vr_v,   vr_cls          = vr_c,
        violated_apps  = va_v,   va_cls          = va_c,

        chart_mfa_gauge    = chart_mfa_gauge,
        chart_by_company   = chart_by_company,
        chart_by_user_type = chart_by_user_type,
        chart_by_department = chart_by_dept,
        chart_by_device    = chart_by_device,
        chart_violations   = chart_violations,
        chart_idm_risk     = chart_idm_risk,
        chart_top10_roles  = chart_top10_roles,
    )

    # ── Save output ──────────────────────────────────────────────────────────
    out_folder = Path(CONFIG["output_folder"])
    out_folder.mkdir(parents=True, exist_ok=True)
    out_file = out_folder / f"IT_Governance_Dashboard_{now:%Y_%m_%d}.html"
    out_file.write_text(html, encoding="utf-8")

    print(f"[INFO] Dashboard saved: {out_file}")
    print(f"[INFO] Open in browser to view.")
    return str(out_file)


if __name__ == "__main__":
    output_path = main()
    # Print the output path on its own line so Power Automate can parse it
    print(f"OUTPUT_FILE={output_path}")
