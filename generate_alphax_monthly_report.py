#!/usr/bin/env python3
"""
generate_alphax_monthly_report.py

Monthly Executive Status Report — Project AlphaX
Output: AlphaX/AlphaX_Monthly_Status_Report_{MMM_YYYY}.pdf

Designed for scheduled monthly execution via Power Automate (or any scheduler).
  - All "days to gate" values auto-calculate from today's date.
  - No interactive prompts. No arguments required.
  - Output filename includes the current month and year.

Usage:
    python generate_alphax_monthly_report.py

Power Automate:
    Use a "Run script" / "Run Python script" action pointing to this file.
    The embedded Python path below (C:/Program Files/px/python.exe) matches
    the Bosch-provisioned embedded Python. Adjust if your environment differs.
"""

import sys, os, datetime

sys.path.insert(0, os.path.join(os.path.expanduser("~"), "py_packages"))

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# ─────────────────────────────────────────────────────────────────────────────
# PROJECT CONFIGURATION — update these for a new engagement
# ─────────────────────────────────────────────────────────────────────────────

PROJECT_CODE     = "ALPHAX"
PROJECT_TITLE    = "Project AlphaX — Executive Status Report"
PROJECT_SUBTITLE = "IT Carve-Out: Bosch Battery Division  →  NewCo Battery Division  |  Stand Alone Model"
SELLER           = "Robert Bosch GmbH"
BUYER            = "NewCo Battery Division"
PMO              = "KPMG"
MODEL            = "Stand Alone — No Merger Zone"
SCOPE_TEXT       = "3,000 users · 35 sites"
LABOUR_BUDGET    = "EUR 2,969,520"
DURATION         = "19 months"

# Key programme milestone dates (ISO format YYYY-MM-DD)
DATES = {
    "kickoff":  "2026-06-01",
    "signing":  "2026-06-15",
    "qg1":      "2026-08-31",
    "qg2":      "2026-11-30",
    "qg3":      "2027-06-30",
    "qg4":      "2027-09-15",
    "day1":     "2027-10-01",
    "closure":  "2027-12-31",
}

PHASE_ROWS = [
    ("Phase 1: Initialization",    "Jun – Aug 2026",       "Upcoming", "Governance, IT inventory (35 sites), TSA scope definition"),
    ("Phase 2: Concept",           "Sep – Nov 2026",       "Planned",  "IT architecture, migration strategy, WAN / co-lo vendor"),
    ("Phase 3: Development",       "Dec 2026 – Jun 2027",  "Planned",  "NewCo WAN, AD forest, SAP shell copy, M365 tenant, IAM"),
    ("Phase 4: Implementation",    "Jul – Sep 2027",       "Planned",  "Wave cutovers, 2,500 device reimage, app migrations"),
    ("Phase 5: GoLive & Closure",  "Oct – Dec 2027",       "CRITICAL", "Day 1 GoLive (Oct 1), 90-day hypercare, TSA exit, closure"),
]

TOP_RISKS = [
    ("R1: NewCo Legal Entity TBC",          "16", "HIGH"),
    ("R2: WAN Lead Time (4–6 months)",       "15", "HIGH"),
    ("R3: SAP ERP Strategy Unresolved",      "12", "HIGH"),
    ("R4: No CISO Appointed for NewCo",      "12", "MEDIUM"),
    ("R5: Compressed 19-month Timeline",     "12", "MEDIUM"),
]

BUDGET_ROWS = [
    ("Governance / PMO",           "EUR 453,600",    "15.3%"),
    ("Bosch IT Team (Seller)",     "EUR 1,382,400",  "46.6%"),
    ("NewCo Battery (Buyer)",      "EUR 124,800",    "4.2%"),
    ("External Partners (KPMG)",   "EUR 378,000",    "12.7%"),
    ("Cross-Functional Teams",     "EUR 351,000",    "11.8%"),
    ("Regional Teams EMEA/AP/AM",  "EUR 214,920",    "7.2%"),
    ("Executive Leadership",       "EUR 64,800",     "2.2%"),
    ("TOTAL",                      "EUR 2,969,520",  "100%"),
]

PROGRAMME_NOTES = [
    "NewCo legal entity registration must begin immediately — Mexico and China require 12+ weeks lead time.",
    "WAN vendor orders for 35 sites must be placed at QG2 (Nov 2026) — 4–6 month lead time to Apr 2027 delivery.",
    "ERP strategy (SAP shell copy vs. greenfield) must be decided at QG1 — any delay impacts Phase 3 by 3+ months.",
    "KPMG engagement scope must be confirmed by Apr 10, 2026 to staff Phase 1 PMO from Jun 1 kickoff.",
    "NewCo CISO appointment and IAM platform selection must begin Phase 1; ISO 27001 baseline assessment by QG2.",
    f"Budget baseline ({LABOUR_BUDGET} labour) to be formally approved and locked at QG1 — 31 Aug 2026.",
]

# ─────────────────────────────────────────────────────────────────────────────
# AUTO-CALCULATED FROM TODAY'S DATE
# ─────────────────────────────────────────────────────────────────────────────

TODAY        = datetime.date.today()
REPORT_DATE  = TODAY.strftime("%B %d, %Y")
FILE_MONTH   = TODAY.strftime("%b_%Y")          # e.g. Apr_2026
NEXT_REVIEW  = (TODAY + datetime.timedelta(days=14)).strftime("%B %d, %Y")

HERE      = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(HERE, "Bosch-Logo.png")
OUT_DIR   = os.path.join(HERE, "AlphaX")
os.makedirs(OUT_DIR, exist_ok=True)
OUT_PATH  = os.path.join(OUT_DIR, f"AlphaX_Monthly_Status_Report_{FILE_MONTH}.pdf")


def days_to(iso):
    return (datetime.date.fromisoformat(iso) - TODAY).days


def fmt_days(iso):
    d = days_to(iso)
    return "Complete" if d < 0 else str(d)


def auto_status(iso, force_critical=False):
    d = days_to(iso)
    if d < 0:    return "Complete"
    if d <= 30:  return "CRITICAL"
    if d <= 120: return "Upcoming"
    if force_critical: return "CRITICAL"
    return "Planned"


# ─────────────────────────────────────────────────────────────────────────────
# BOSCH DIGITAL COLOUR PALETTE
# ─────────────────────────────────────────────────────────────────────────────

C_RED     = colors.HexColor("#ED0007")   # Bosch Digital primary red
C_NAVY    = colors.HexColor("#004975")   # Bosch Digital dark navy
C_TEAL    = colors.HexColor("#0A4F4B")   # Bosch Digital dark teal
C_GREEN   = colors.HexColor("#00512A")   # Bosch Digital dark green
C_BLUE    = colors.HexColor("#007BC0")   # Bosch Digital blue
C_MID     = colors.HexColor("#43464A")   # body text
C_MUTED   = colors.HexColor("#71767C")   # muted / labels
C_LGRAY   = colors.HexColor("#F5F5F5")   # alternating row background
C_LINE    = colors.HexColor("#D4D4D4")   # borders
C_WHITE   = colors.white
C_HI_BG   = colors.HexColor("#FEE8EA")
C_HI_TX   = colors.HexColor("#9B0006")
C_MD_BG   = colors.HexColor("#FFF3E0")
C_MD_TX   = colors.HexColor("#C66000")
C_GN_BG   = colors.HexColor("#E6F4EC")
C_GN_TX   = colors.HexColor("#00512A")
C_BL_BG   = colors.HexColor("#E8F4FD")
C_BL_TX   = colors.HexColor("#1565C0")
C_NVLT    = colors.HexColor("#AACCEE")   # light navy (header sub-text)
C_METLBL  = colors.HexColor("#FFBBBB")   # meta bar labels

STATUS_STYLES = {
    "CRITICAL": (C_HI_BG, C_HI_TX),
    "Upcoming": (C_BL_BG, C_BL_TX),
    "Planned":  (C_LGRAY, C_MUTED),
    "Complete": (C_GN_BG, C_GN_TX),
}

# ─────────────────────────────────────────────────────────────────────────────
# PAGE GEOMETRY
# ─────────────────────────────────────────────────────────────────────────────

W, H   = A4           # 595.27 x 841.89 pt
LM     = 22           # left margin
RM     = 22           # right margin
TM     = 22           # top margin
BM     = 22           # bottom margin
CW     = W - LM - RM  # content width ≈ 551 pt


# ─────────────────────────────────────────────────────────────────────────────
# DRAWING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def section_header(c, y, label, bg=C_NAVY, height=16):
    """Draw a coloured section header bar. Returns pts consumed (incl. gap)."""
    c.setFillColor(bg)
    c.rect(LM, y - height, CW, height, fill=1, stroke=0)
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(LM + 6, y - height + 4, label.upper())
    return height + 3   # height + 3 pt gap


def table_rows(c, y, headers, rows, col_widths, row_h=15,
               hdr_bg=C_NAVY, status_col=None, bold_last=False):
    """Draw header + data rows. Returns total height consumed."""
    hdr_h = row_h + 1
    x0 = [LM]
    for w in col_widths[:-1]:
        x0.append(x0[-1] + w)

    # Header row
    c.setFillColor(hdr_bg)
    c.rect(LM, y - hdr_h, CW, hdr_h, fill=1, stroke=0)
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 7)
    for h_, x_ in zip(headers, x0):
        c.drawString(x_ + 3, y - hdr_h + 4, str(h_))

    # Data rows
    for ri, row in enumerate(rows):
        ry   = y - hdr_h - (ri + 1) * row_h
        last = (ri == len(rows) - 1)
        bg   = colors.HexColor("#DDE8F0") if (bold_last and last) \
               else (C_LGRAY if ri % 2 == 0 else C_WHITE)
        c.setFillColor(bg)
        c.rect(LM, ry, CW, row_h, fill=1, stroke=0)
        c.setStrokeColor(C_LINE)
        c.setLineWidth(0.3)
        c.rect(LM, ry, CW, row_h, fill=0, stroke=1)

        for ci, (cell, x_) in enumerate(zip(row, x0)):
            if status_col is not None and ci == status_col:
                sc   = str(cell)
                bg_s, tx_s = STATUS_STYLES.get(sc, (C_LGRAY, C_MUTED))
                pw   = col_widths[ci] - 6
                c.setFillColor(bg_s)
                c.roundRect(x_ + 2, ry + 2.5, pw, row_h - 5, 2, fill=1, stroke=0)
                c.setFillColor(tx_s)
                c.setFont("Helvetica-Bold", 6.5)
                c.drawCentredString(x_ + pw / 2 + 2, ry + 4, sc)
            else:
                c.setFont("Helvetica-Bold" if (bold_last and last) else "Helvetica", 7)
                c.setFillColor(C_NAVY if (bold_last and last) else C_MID)
                c.drawString(x_ + 3, ry + 4, str(cell))

    # Outer border
    total_h = hdr_h + len(rows) * row_h
    c.setStrokeColor(C_LINE)
    c.setLineWidth(0.5)
    c.rect(LM, y - total_h, CW, total_h, fill=0, stroke=1)
    return total_h


# ─────────────────────────────────────────────────────────────────────────────
# REPORT BUILDER
# ─────────────────────────────────────────────────────────────────────────────

def build_report():
    c = canvas.Canvas(OUT_PATH, pagesize=A4)
    y = H - TM   # y-cursor starts at top, decrements downward

    # ── HEADER ───────────────────────────────────────────────────────────────
    HDR_H = 52
    c.setFillColor(C_NAVY)
    c.rect(LM, y - HDR_H, CW, HDR_H, fill=1, stroke=0)
    # Red left accent bar
    c.setFillColor(C_RED)
    c.rect(LM, y - HDR_H, 4, HDR_H, fill=1, stroke=0)
    # Bosch logo
    if os.path.exists(LOGO_PATH):
        try:
            c.drawImage(ImageReader(LOGO_PATH), LM + 8, y - HDR_H + 7,
                        width=62, height=40, preserveAspectRatio=True, mask="auto")
        except Exception:
            pass
    # Title
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(W / 2, y - 21, PROJECT_TITLE)
    c.setFont("Helvetica", 7.5)
    c.setFillColor(C_NVLT)
    c.drawCentredString(W / 2, y - 34, PROJECT_SUBTITLE)
    # Right meta
    c.setFont("Helvetica", 7)
    c.setFillColor(C_NVLT)
    c.drawRightString(W - RM - 4, y - 19, f"Report Date: {REPORT_DATE}")
    c.drawRightString(W - RM - 4, y - 30, f"PMO: {PMO}")
    c.drawRightString(W - RM - 4, y - 41, f"Next Review: {NEXT_REVIEW}")
    y -= HDR_H

    # ── METADATA BAR ─────────────────────────────────────────────────────────
    META_H = 26
    c.setFillColor(C_RED)
    c.rect(LM, y - META_H, CW, META_H, fill=1, stroke=0)
    segments = [
        ("PROGRAMME STATUS",  "PRE-LAUNCH"),
        ("DURATION",          DURATION),
        ("DAY 1 GO-LIVE",     "01 Oct 2027"),
        ("SCOPE",             SCOPE_TEXT),
        ("LABOUR BUDGET",     LABOUR_BUDGET),
        ("DAYS TO DAY 1",     str(days_to(DATES["day1"]))),
    ]
    seg_w = CW / len(segments)
    for i, (lbl, val) in enumerate(segments):
        x = LM + i * seg_w
        if i > 0:
            c.setStrokeColor(colors.HexColor("#FF5566"))
            c.setLineWidth(0.5)
            c.line(x, y - META_H + 4, x, y - 4)
        c.setFont("Helvetica", 5.5)
        c.setFillColor(C_METLBL)
        c.drawCentredString(x + seg_w / 2, y - 8, lbl)
        c.setFont("Helvetica-Bold", 8)
        c.setFillColor(C_WHITE)
        c.drawCentredString(x + seg_w / 2, y - 19, val)
    y -= (META_H + 5)

    # ── PHASE STATUS & TIMELINE ───────────────────────────────────────────────
    y -= section_header(c, y, "Phase Status & Timeline")
    ph_cols = [140, 88, 58, CW - 140 - 88 - 58]
    ph_hdrs = ["Phase", "Timeline", "Status", "Key Deliverable"]
    y -= (table_rows(c, y, ph_hdrs, PHASE_ROWS, ph_cols, row_h=17, status_col=2) + 4)

    # ── KEY RISKS & IMMEDIATE ACTIONS ────────────────────────────────────────
    y -= section_header(c, y, "Key Risks & Immediate Actions", bg=C_TEAL)
    col_w  = (CW - 4) / 2
    rx     = LM + col_w + 4
    sub_h  = 14
    rrow_h = 17
    n      = len(TOP_RISKS)
    next_d = (TODAY + datetime.timedelta(days=14)).strftime("%b %d, %Y")

    # Sub-column headers
    c.setFillColor(colors.HexColor("#003355"))
    c.rect(LM, y - sub_h, col_w, sub_h, fill=1, stroke=0)
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(LM + 4, y - sub_h + 4, "TOP RISKS  (Rating)")

    c.setFillColor(colors.HexColor("#063333"))
    c.rect(rx, y - sub_h, col_w, sub_h, fill=1, stroke=0)
    c.setFillColor(C_WHITE)
    c.drawString(rx + 4, y - sub_h + 4, f"IMMEDIATE ACTIONS  (by {next_d})")

    decisions = [
        "Approve Project Charter & Governance Model",
        f"Authorize {LABOUR_BUDGET} Labour Budget (QG1)",
        "Confirm KPMG engagement scope (by Apr 10, 2026)",
        "Confirm Phase 1 Kickoff — Jun 1, 2026",
        "Initiate NewCo legal entity registration immediately",
    ]
    for ri in range(n):
        ry = y - sub_h - (ri + 1) * rrow_h
        lbl, rating, lvl = TOP_RISKS[ri]
        bg_r = C_HI_BG if lvl == "HIGH" else C_MD_BG
        tx_r = C_HI_TX if lvl == "HIGH" else C_MD_TX
        # Risk row (left)
        c.setFillColor(bg_r)
        c.rect(LM, ry, col_w, rrow_h, fill=1, stroke=0)
        c.setStrokeColor(tx_r); c.setLineWidth(3)
        c.line(LM, ry, LM, ry + rrow_h)
        c.setStrokeColor(C_LINE); c.setLineWidth(0.3)
        c.rect(LM, ry, col_w, rrow_h, fill=0, stroke=1)
        c.setFillColor(tx_r)
        c.setFont("Helvetica-Bold", 7)
        c.drawString(LM + 6, ry + 5, f"({rating} {lvl})  {lbl}")
        # Decision row (right)
        c.setFillColor(C_LGRAY if ri % 2 == 0 else C_WHITE)
        c.rect(rx, ry, col_w, rrow_h, fill=1, stroke=0)
        c.setStrokeColor(C_LINE); c.setLineWidth(0.3)
        c.rect(rx, ry, col_w, rrow_h, fill=0, stroke=1)
        c.setFillColor(C_GREEN); c.setFont("Helvetica-Bold", 8)
        c.drawString(rx + 4, ry + 5, "▶")
        c.setFillColor(C_MID); c.setFont("Helvetica", 7)
        c.drawString(rx + 14, ry + 5, decisions[ri])

    y -= (sub_h + n * rrow_h + 5)

    # ── BUDGET SNAPSHOT ───────────────────────────────────────────────────────
    y -= section_header(c, y, "Budget Snapshot — Labour Only  (excl. WAN, hardware, licences, co-lo)", bg=C_TEAL)
    bud_cols = [178, 95, CW - 178 - 95]
    bud_hdrs = ["Category", "Cost (EUR)", "Share"]
    y -= (table_rows(c, y, bud_hdrs, BUDGET_ROWS, bud_cols,
                     row_h=15, bold_last=True) + 5)

    # ── CRITICAL MILESTONES & QUALITY GATES ──────────────────────────────────
    y -= section_header(c, y, "Critical Milestones & Quality Gates")
    ms_rows = [
        ("Programme Kickoff",                   "01 Jun 2026", auto_status(DATES["kickoff"]),  fmt_days(DATES["kickoff"])),
        ("Signing — Frozen Zone Activated",      "15 Jun 2026", auto_status(DATES["signing"]),  fmt_days(DATES["signing"])),
        ("QG1 — Initialization Gate",            "31 Aug 2026", auto_status(DATES["qg1"]),      fmt_days(DATES["qg1"])),
        ("QG2 — Concept Gate",                   "30 Nov 2026", auto_status(DATES["qg2"]),      fmt_days(DATES["qg2"])),
        ("QG3 — Development Gate",               "30 Jun 2027", auto_status(DATES["qg3"]),      fmt_days(DATES["qg3"])),
        ("QG4 — Implementation Gate (Go/No-Go)", "15 Sep 2027", auto_status(DATES["qg4"]),      fmt_days(DATES["qg4"])),
        ("★  DAY 1 GO-LIVE — NewCo Live",        "01 Oct 2027", "CRITICAL",                     fmt_days(DATES["day1"])),
        ("QG5 — Programme Closure",              "31 Dec 2027", auto_status(DATES["closure"]),   fmt_days(DATES["closure"])),
    ]
    ms_cols = [192, 76, 60, CW - 192 - 76 - 60]
    ms_hdrs = ["Milestone", "Target Date", "Status", "Days to Gate"]
    y -= (table_rows(c, y, ms_hdrs, ms_rows, ms_cols, row_h=14, status_col=2) + 5)

    # ── PROGRAMME NOTES & NEXT STEPS ─────────────────────────────────────────
    y -= section_header(c, y, "Programme Notes & Key Actions — Next Period", bg=C_NAVY)
    note_h = 12.5
    for i, note in enumerate(PROGRAMME_NOTES):
        ny = y - (i + 1) * note_h
        c.setFillColor(C_LGRAY if i % 2 == 0 else C_WHITE)
        c.rect(LM, ny, CW, note_h, fill=1, stroke=0)
        c.setStrokeColor(C_LINE); c.setLineWidth(0.3)
        c.line(LM, ny, W - RM, ny)
        c.setFillColor(C_RED); c.setFont("Helvetica-Bold", 7.5)
        c.drawString(LM + 5, ny + 3, "▸")
        c.setFillColor(C_MID); c.setFont("Helvetica", 7)
        c.drawString(LM + 14, ny + 3, note)
    # Border around notes block
    c.setStrokeColor(C_LINE); c.setLineWidth(0.5)
    c.rect(LM, y - len(PROGRAMME_NOTES) * note_h, CW,
           len(PROGRAMME_NOTES) * note_h, fill=0, stroke=1)
    y -= (len(PROGRAMME_NOTES) * note_h + 4)

    # ── BOSCH DIGITAL ACCENT LINE ─────────────────────────────────────────────
    c.setFillColor(C_WHITE)
    c.rect(LM, BM + 23, CW, 3, fill=1, stroke=0)
    # Rainbow gradient approximated with 5 rect segments
    accent_colors = [C_NAVY, C_BLUE, C_TEAL, C_RED, C_GREEN]
    seg = CW / len(accent_colors)
    for i, ac in enumerate(accent_colors):
        c.setFillColor(ac)
        c.rect(LM + i * seg, BM + 23, seg, 3, fill=1, stroke=0)

    # ── FOOTER ────────────────────────────────────────────────────────────────
    FTR_H = 22
    c.setFillColor(C_NAVY)
    c.rect(LM, BM, CW, FTR_H, fill=1, stroke=0)
    c.setFillColor(C_WHITE); c.setFont("Helvetica-Bold", 7)
    c.drawString(LM + 5, BM + 13,
                 f"{PROJECT_CODE}  |  IT Carve-Out  |  {SELLER}  →  {BUYER}")
    c.setFillColor(C_NVLT); c.setFont("Helvetica", 6.5)
    c.drawString(LM + 5, BM + 4,
                 "Distribution: Bosch Executive Leadership · NewCo Board · Steering Committee · IT Leadership"
                 "  |  Confidentiality: Internal Use Only")
    c.drawRightString(W - RM - 4, BM + 4, f"Generated: {REPORT_DATE}")

    c.save()
    print(f"[OK] Monthly status report written:")
    print(f"     {OUT_PATH}")


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        build_report()
    except Exception as e:
        print(f"[ERROR] Report generation failed: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
