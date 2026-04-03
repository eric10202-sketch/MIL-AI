#!/usr/bin/env python3
"""
generate_bravo_monthly_report.py

Monthly Executive Status Report — Project Bravo
Output: Bravo/Bravo_Monthly_Status_Report_{MMM_YYYY}.pdf

Usage:
    python generate_bravo_monthly_report.py

No interactive prompts. Output filename auto-refreshes from runtime date.
"""

import sys, os, datetime

sys.path.insert(0, os.path.join(os.path.expanduser("~"), "py_packages"))

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from pathlib import Path

HERE = Path(__file__).parent

# ── PROJECT CONFIGURATION ─────────────────────────────────────────────────────
PROJECT_CODE     = "BRAVO"
PROJECT_TITLE    = "Project Bravo — Executive Status Report"
PROJECT_SUBTITLE = "IT Carve-Out: Bosch BGSW AI Business  →  50/50 JV with Tata  |  Bosch Leadership Control"
SELLER           = "Bosch BGSW"
BUYER            = "Tata (50% shareholder; Bosch leadership control)"
PMO              = "Riyaz Ahmed Syed Ahmed (BD/MIL-PSM4)"
MODEL            = "Combination — Bosch-led JV (no TSA, no Merger Zone)"
SCOPE_TEXT       = "70 users · 2 India sites · 17 AI apps · No ERP"
LABOUR_BUDGET    = "EUR 554,040"
DURATION         = "7 months (Apr–Oct 2026)"

ANTITRUST        = "Not applicable — Bosch leadership control"
TSA              = "None required"

# Key milestone dates (ISO)
DATES = {
    "kickoff":    "2026-04-01",
    "signing":    "2026-04-07",
    "qg0":        "2026-04-30",
    "qg123":      "2026-06-26",
    "golive":     "2026-07-01",
    "qg4":        "2026-07-02",
    "hypercare":  "2026-09-25",
    "qg5":        "2026-10-30",
}

PHASE_ROWS = [
    ("Phase 1: Initialization",           "01 Apr – 30 Apr 2026", "In Progress", "Governance, 17-app IT inventory, JV legal setup, QG0"),
    ("Phase 2: Concept & Architecture",   "01 May – 29 May 2026", "Planned",     "As-is analysis, JV IT architecture, migration strategy"),
    ("Phase 3: Build & Test",             "01 Jun – 26 Jun 2026", "Planned",     "JV infra build, 17-app migration, 70-device reimage, UAT"),
    ("Phase 4: Cutover & GoLive",         "29 Jun – 02 Jul 2026", "Planned",     "Cutover execution, Day 1 GoLive (01 Jul), QG4"),
    ("Phase 5: Hypercare & Closure",      "06 Jul – 30 Oct 2026", "Planned",     "60-day hypercare, programme closure, QG5"),
]

RISK_ROWS = [
    ("BGSW/JV bandwidth (Apr–Jun sprint)",   "16", "Amber", "Dedicated lead; SteerCo weekly"),
    ("JV MCA registration delay (India)",    "15", "Green", "India counsel engaged; MCA week 1"),
    ("Key BGSW IT staff unavailable (Jun)",  "12", "Amber", "Jun leave freeze; backups by 29 May"),
    ("AD JV forest separation incomplete",   "10", "Green", "Build Jun 1; dress rehearsal Jun 24"),
    ("Tata JV team not staffed for May",     "9",  "Amber", "Confirm nominees at kickoff Apr 1"),
]

BUDGET_ROWS = [
    ("Programme Management",   "EUR 164,000", "On target"),
    ("IT Project Management",  "EUR 113,600", "On target"),
    ("Hypercare & Closure",    "EUR  85,800", "On target"),
    ("Infrastructure & Cloud", "EUR  58,720", "On target"),
    ("Application Migration",  "EUR  58,240", "On target"),
    ("Architecture & Design",  "EUR  31,680", "On target"),
    ("Legal & Compliance",     "EUR  24,000", "On target"),
    ("Client Workplace",       "EUR  14,400", "On target"),
    ("HR IT",                  "EUR   3,600", "On target"),
]

# ── PATHS ─────────────────────────────────────────────────────────────────────
LOGO_PATH = HERE / "Bosch.png"
OUT_DIR   = HERE / "Bravo"
OUT_DIR.mkdir(exist_ok=True)

# ── COLOURS ──────────────────────────────────────────────────────────────────
NAVY    = colors.HexColor("#003b6e")
MID     = colors.HexColor("#005199")
ACC     = colors.HexColor("#0066CC")
LT      = colors.HexColor("#e4edf9")
GOOD    = colors.HexColor("#007A33")
WARN    = colors.HexColor("#E8A000")
BAD     = colors.HexColor("#CC0000")
WHITE   = colors.white
BLACK   = colors.black
BGRAY   = colors.HexColor("#f4f6f9")
MUTED   = colors.HexColor("#5a6478")


def days_to(date_str: str) -> int:
    target = datetime.date.fromisoformat(date_str)
    return (target - datetime.date.today()).days


def rag_color(status: str):
    s = status.lower()
    if "green" in s:  return GOOD
    if "amber" in s:  return WARN
    if "red"   in s:  return BAD
    return MID


def draw_report(c_: canvas.Canvas, w: float, h: float):
    today = datetime.date.today()
    report_month = today.strftime("%B %Y")
    y = h

    # ── HEADER BAND ──────────────────────────────────────────────────────────
    c_.setFillColor(NAVY)
    c_.rect(0, h - 80, w, 80, fill=1, stroke=0)

    # Logo
    if LOGO_PATH.exists():
        logo = ImageReader(str(LOGO_PATH))
        # White background behind logo
        c_.setFillColor(WHITE)
        c_.rect(28, h - 65, 90, 40, fill=1, stroke=0)
        c_.drawImage(logo, 30, h - 63, height=36, preserveAspectRatio=True, mask="auto")

    # Title
    c_.setFillColor(WHITE)
    c_.setFont("Helvetica-Bold", 14)
    c_.drawString(135, h - 34, PROJECT_TITLE)
    c_.setFont("Helvetica", 9)
    c_.drawString(135, h - 49, PROJECT_SUBTITLE)
    c_.setFont("Helvetica", 8)
    c_.drawString(135, h - 62, f"Report Month: {report_month}  |  PM: {PMO}")

    # Date top-right
    c_.setFont("Helvetica-Bold", 18)
    c_.setFillColor(colors.HexColor("#ffd700"))
    dtg = days_to(DATES["golive"])
    c_.drawRightString(w - 20, h - 38, f"{dtg}")
    c_.setFont("Helvetica", 8)
    c_.setFillColor(WHITE)
    c_.drawRightString(w - 20, h - 52, "days to GoLive (01 Jul 2026)")

    y = h - 90

    # ── KEY FACTS STRIP ──────────────────────────────────────────────────────
    facts = [
        ("Seller", SELLER),
        ("Buyer", "Tata (Bosch leadership)"),
        ("Sites", "2 India"),
        ("Users", "70"),
        ("Apps", "17 AI (no ERP)"),
        ("TSA", TSA),
    ]
    col_w = (w - 40) / len(facts)
    c_.setFillColor(LT)
    c_.rect(20, y - 36, w - 40, 34, fill=1, stroke=0)
    for i, (lbl, val) in enumerate(facts):
        x0 = 20 + i * col_w
        c_.setFillColor(MID)
        c_.setFont("Helvetica-Bold", 7)
        c_.drawString(x0 + 6, y - 14, lbl.upper())
        c_.setFillColor(BLACK)
        c_.setFont("Helvetica-Bold", 8)
        c_.drawString(x0 + 6, y - 27, val)
    y -= 44

    # ── PHASES TABLE ─────────────────────────────────────────────────────────
    c_.setFillColor(MID)
    c_.rect(20, y - 14, w - 40, 14, fill=1, stroke=0)
    c_.setFillColor(WHITE)
    c_.setFont("Helvetica-Bold", 8)
    c_.drawString(24, y - 10, "PHASE PROGRESS")
    y -= 16

    cols_phase = [160, 130, 72, 0]  # widths; last is remainder
    hdrs_phase = ["Phase", "Window", "Status", "Highlights"]
    x_phase = [20, 20 + cols_phase[0], 20 + cols_phase[0] + cols_phase[1],
               20 + cols_phase[0] + cols_phase[1] + cols_phase[2]]

    c_.setFillColor(LT)
    c_.rect(20, y - 14, w - 40, 14, fill=1, stroke=0)
    c_.setFillColor(NAVY)
    c_.setFont("Helvetica-Bold", 8)
    for xi, hd in zip(x_phase, hdrs_phase):
        c_.drawString(xi + 4, y - 10, hd)
    y -= 16

    for i, row in enumerate(PHASE_ROWS):
        bg = colors.HexColor("#f4f6f9") if i % 2 == 0 else WHITE
        c_.setFillColor(bg)
        c_.rect(20, y - 13, w - 40, 13, fill=1, stroke=0)
        c_.setFillColor(BLACK)
        c_.setFont("Helvetica", 8)
        c_.drawString(x_phase[0] + 4, y - 9, row[0])
        c_.drawString(x_phase[1] + 4, y - 9, row[1])
        # Status pill
        sc = rag_color(row[2])
        c_.setFillColor(sc)
        c_.roundRect(x_phase[2] + 4, y - 11, 62, 11, 4, fill=1, stroke=0)
        c_.setFillColor(WHITE if row[2] != "Amber" else BLACK)
        c_.setFont("Helvetica-Bold", 7)
        c_.drawCentredString(x_phase[2] + 35, y - 7, row[2].upper())
        c_.setFillColor(MUTED)
        c_.setFont("Helvetica", 7)
        c_.drawString(x_phase[3] + 4, y - 9, row[3])
        y -= 14

    y -= 8

    # ── RISK TABLE ───────────────────────────────────────────────────────────
    c_.setFillColor(MID)
    c_.rect(20, y - 14, w - 40, 14, fill=1, stroke=0)
    c_.setFillColor(WHITE)
    c_.setFont("Helvetica-Bold", 8)
    c_.drawString(24, y - 10, "TOP RISKS")
    y -= 16

    cols_risk = [220, 36, 58, 0]
    hdrs_risk = ["Risk Description", "P×I", "Status", "Mitigation"]
    x_risk = [20, 240, 276, 334]

    c_.setFillColor(LT)
    c_.rect(20, y - 14, w - 40, 14, fill=1, stroke=0)
    c_.setFillColor(NAVY)
    c_.setFont("Helvetica-Bold", 8)
    for xi, hd in zip(x_risk, hdrs_risk):
        c_.drawString(xi + 4, y - 10, hd)
    y -= 16

    for i, row in enumerate(RISK_ROWS):
        bg = colors.HexColor("#f4f6f9") if i % 2 == 0 else WHITE
        c_.setFillColor(bg)
        c_.rect(20, y - 13, w - 40, 13, fill=1, stroke=0)
        c_.setFillColor(BLACK)
        c_.setFont("Helvetica", 8)
        c_.drawString(x_risk[0] + 4, y - 9, row[0])
        c_.drawCentredString(x_risk[1] + 18, y - 9, row[1])
        sc = rag_color(row[2])
        c_.setFillColor(sc)
        c_.roundRect(x_risk[2] + 2, y - 11, 54, 11, 4, fill=1, stroke=0)
        c_.setFillColor(WHITE if row[2] != "Amber" else BLACK)
        c_.setFont("Helvetica-Bold", 7)
        c_.drawCentredString(x_risk[2] + 29, y - 7, row[2].upper())
        c_.setFillColor(MUTED)
        c_.setFont("Helvetica", 7)
        c_.drawString(x_risk[3] + 4, y - 9, row[3])
        y -= 14

    y -= 8

    # ── BUDGET TABLE ─────────────────────────────────────────────────────────
    c_.setFillColor(MID)
    c_.rect(20, y - 14, w - 40, 14, fill=1, stroke=0)
    c_.setFillColor(WHITE)
    c_.setFont("Helvetica-Bold", 8)
    c_.drawString(24, y - 10, f"BUDGET SUMMARY  (Labour Total: {LABOUR_BUDGET}  |  CAPEX: TBC – to be approved at QG0)")
    y -= 16

    col_ww = (w - 40) / 3
    for i, row in enumerate(BUDGET_ROWS):
        bg = colors.HexColor("#f4f6f9") if i % 2 == 0 else WHITE
        c_.setFillColor(bg)
        c_.rect(20, y - 12, w - 40, 12, fill=1, stroke=0)
        c_.setFillColor(BLACK)
        c_.setFont("Helvetica", 8)
        c_.drawString(24, y - 8, row[0])
        c_.drawString(20 + col_ww + 4, y - 8, row[1])
        c_.setFillColor(GOOD)
        c_.setFont("Helvetica", 8)
        c_.drawString(20 + 2 * col_ww + 4, y - 8, row[2])
        y -= 13

    y -= 8

    # ── KEY MILESTONE COUNTDOWN ───────────────────────────────────────────────
    c_.setFillColor(MID)
    c_.rect(20, y - 14, w - 40, 14, fill=1, stroke=0)
    c_.setFillColor(WHITE)
    c_.setFont("Helvetica-Bold", 8)
    c_.drawString(24, y - 10, "MILESTONE COUNTDOWN")
    y -= 16

    milestones = [
        ("QG0 – Initialization Gate", DATES["qg0"]),
        ("QG1/2/3 – Combined Gate",   DATES["qg123"]),
        ("Day 1 GoLive",              DATES["golive"]),
        ("QG4 – GoLive Gate",         DATES["qg4"]),
        ("QG5 – Programme Closure",   DATES["qg5"]),
    ]
    n_ms = len(milestones)
    ms_w = (w - 40) / n_ms
    c_.setFillColor(LT)
    c_.rect(20, y - 44, w - 40, 44, fill=1, stroke=0)
    for j, (name, dt) in enumerate(milestones):
        x0 = 20 + j * ms_w
        d = days_to(dt)
        c_.setFillColor(NAVY)
        c_.setFont("Helvetica-Bold", 16)
        c_.drawCentredString(x0 + ms_w / 2, y - 22, str(d))
        c_.setFillColor(ACC)
        c_.setFont("Helvetica-Bold", 7)
        c_.drawCentredString(x0 + ms_w / 2, y - 33, "days")
        c_.setFillColor(MUTED)
        c_.setFont("Helvetica", 6.5)
        c_.drawCentredString(x0 + ms_w / 2, y - 41, name)
    y -= 52

    # ── JV MODEL NOTE ────────────────────────────────────────────────────────
    c_.setFillColor(LT)
    c_.rect(20, y - 36, w - 40, 34, fill=1, stroke=0)
    c_.setFillColor(NAVY)
    c_.setFont("Helvetica-Bold", 8)
    c_.drawString(24, y - 12, "KEY PROGRAMME NOTE — BOSCH JV LEADERSHIP CONTROL")
    c_.setFillColor(MUTED)
    c_.setFont("Helvetica", 7.5)
    c_.drawString(24, y - 23, (
        "Bosch retains operational/leadership control of the 50/50 Tata–Bosch JV. "
        "Post-GoLive, the JV operates as a near-Bosch entity. No TSA, no Merger Zone, "
        "no antitrust filings required."
    ))
    c_.drawString(24, y - 33, (
        "IT separation complexity is significantly lower than a Stand Alone model. "
        "BGSW governance and escalation channels remain available to the JV post-GoLive."
    ))
    y -= 42

    # ── FOOTER ───────────────────────────────────────────────────────────────
    c_.setFillColor(LT)
    c_.rect(0, 0, w, 22, fill=1, stroke=0)
    c_.setFillColor(MUTED)
    c_.setFont("Helvetica", 7)
    c_.drawString(20, 8, f"Project Bravo | Monthly Status Report | {report_month} | PM: {PMO} | Confidential — Internal Use Only")
    c_.drawRightString(w - 20, 8, f"Seller: {SELLER}  |  Model: {MODEL}")


def main():
    today = datetime.date.today()
    fname = f"Bravo_Monthly_Status_Report_{today.strftime('%b_%Y')}.pdf"
    out_path = OUT_DIR / fname

    w, h = A4
    c_ = canvas.Canvas(str(out_path), pagesize=A4)
    draw_report(c_, w, h)
    c_.save()
    print(f"Monthly status report written: {out_path}")


if __name__ == "__main__":
    main()
