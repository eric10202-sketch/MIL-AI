#!/usr/bin/env python3
"""
generate_hamburger_monthly_report.py

Monthly Executive Status Report — Project Hamburger
Output: Hamburger/Hamburger_Monthly_Status_Report_{MMM_YYYY}.pdf

Scheduler-ready: all "days to gate" auto-calculate from today's date.
No interactive prompts. No arguments required.

Usage:
    C:/Program Files/px/python.exe generate_hamburger_monthly_report.py
"""

import sys, os, datetime

sys.path.insert(0, os.path.join(os.path.expanduser("~"), "py_packages"))

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# ─────────────────────────────────────────────────────────────────────────────
# PROJECT CONFIGURATION — Project Hamburger
# ─────────────────────────────────────────────────────────────────────────────

PROJECT_CODE     = "HAMBURGER"
PROJECT_TITLE    = "Project Hamburger — Executive Status Report"
PROJECT_SUBTITLE = "IT Carve-Out: Robert Bosch GmbH Solar Energy  →  Undisclosed Buyer  |  Stand Alone Model"
SELLER           = "Robert Bosch GmbH"
BUYER            = "Undisclosed Buyer (legal confidentiality)"
PMO              = "Erik Ho (BD/MIL-ICC)"
MODEL            = "Stand Alone — No Merger Zone"
SCOPE_TEXT       = "2,600 users · 17 sites"
LABOUR_BUDGET    = "EUR 2,649,600"
DURATION         = "14 months"

# Key programme milestone dates (ISO format YYYY-MM-DD)
DATES = {
    "kickoff":       "2026-04-01",
    "qg0":           "2026-04-01",
    "signing":       "2026-04-09",
    "qg1":           "2026-06-30",
    "qg2":           "2026-08-31",
    "qg3":           "2026-10-30",
    "qg4":           "2026-11-30",
    "day1":          "2026-12-01",
    "stab_milestone": "2027-03-01",  # NOT a QG — just marks 90-day hypercare end
    "qg5":           "2027-05-30",  # Bosch QG5 = Programme Closure quality gate
}

PHASE_ROWS = [
    ("Phase 1: Initialization",     "Apr – Jun 2026",  "Upcoming",
     "Governance, IT inventory (17 sites / 2600 users), WAN RFQ, TSA scope"),
    ("Phase 2: Concept",            "Jul – Aug 2026",  "Planned",
     "IT architecture freeze, ERP strategy decision, WAN contracts signed"),
    ("Phase 3: Development & Build","Sep – Oct 2026",  "Planned",
     "WAN/AD/M365 built, ERP separated, 2600 devices configured, SIT1 passed"),
    ("Phase 4: Testing & Cutover",  "Nov 2026",        "Planned",
     "UAT all apps/ERP, cutover plan, Go/No-Go, Help Desk ready globally"),
    ("Phase 5: GoLive & Hypercare",  "Dec 26 – Mar 27", "CRITICAL",
     "Day 1 = GoLive (01 Dec), hypercare 90 calendar days (ends 01 Mar 2027), TSA services active"),
    ("Phase 6: Programme Closure",  "Mar – May 2027",  "Planned",
     "Hypercare Milestone (01 Mar), doc archiving, QG5 Programme Closure (30 May 2027)"),
]

TOP_RISKS = [
    ("R1: GoLive complexity — 17 sites / 2600 users hard date",  "20", "HIGH"),
    ("R2: SAP/ERP separation — shared mandants, custom ABAP",    "20", "HIGH"),
    ("R3: Buyer unknown — scope change risk post-reveal",         "15", "HIGH"),
    ("R4: AD migration 17 sites — hard GoLive dependency",        "12", "HIGH"),
    ("R5: WAN circuit delays — 8–12 week lead per region",        "12", "HIGH"),
]

BUDGET_ROWS = [
    ("Governance / PMO",          "EUR 561,600",    "21.2%"),
    ("Bosch IT Team (Seller)",    "EUR 1,161,600",  "43.8%"),
    ("Buyer / Solar Team",        "EUR 84,000",     "3.2%"),
    ("External Partners",         "EUR 324,000",    "12.2%"),
    ("Cross-Functional Teams",    "EUR 360,000",    "13.6%"),
    ("Regional Teams (17 sites)", "EUR 158,400",    "6.0%"),
    ("TOTAL",                     "EUR 2,649,600",  "100%"),
]

PROGRAMME_NOTES = [
    "WAN RFQ for 17 global sites must be issued by QG1 (30 Jun 2026) — 8-12 week lead time per region.",
    "ERP/SAP strategy (shell copy vs. greenfield) must be decided at QG2 (31 Aug) — delay impacts Phase 3 by 3+ months.",
    "Buyer identity expected to be revealed before QG2 — scope change risk remains open until then (Risk R3).",
    "Active Directory migration (17 sites) is a hard GoLive dependency — must start Phase 3 (Sep 2026).",
    f"Budget baseline ({LABOUR_BUDGET} labour) to be formally approved and locked at QG1 — 30 Jun 2026.",
    "Multi-jurisdiction compliance (GDPR, PIPL China, LGPD Brazil) legal reviews must start Phase 1 immediately.",
]

# ─────────────────────────────────────────────────────────────────────────────
# AUTO-CALCULATED FROM TODAY'S DATE
# ─────────────────────────────────────────────────────────────────────────────

TODAY       = datetime.date.today()
REPORT_DATE = TODAY.strftime("%B %d, %Y")
FILE_MONTH  = TODAY.strftime("%b_%Y")       # e.g. Apr_2026
NEXT_REVIEW = (TODAY + datetime.timedelta(days=14)).strftime("%B %d, %Y")

HERE      = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(HERE, "Bosch.png")
OUT_DIR   = os.path.join(HERE, "Hamburger")
os.makedirs(OUT_DIR, exist_ok=True)
OUT_PATH  = os.path.join(OUT_DIR, f"Hamburger_Monthly_Status_Report_{FILE_MONTH}.pdf")


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

C_RED    = colors.HexColor("#ED0007")
C_NAVY   = colors.HexColor("#004975")
C_TEAL   = colors.HexColor("#0A4F4B")
C_GREEN  = colors.HexColor("#00512A")
C_BLUE   = colors.HexColor("#007BC0")
C_MID    = colors.HexColor("#43464A")
C_MUTED  = colors.HexColor("#71767C")
C_LGRAY  = colors.HexColor("#F5F5F5")
C_LINE   = colors.HexColor("#D4D4D4")
C_WHITE  = colors.white
C_HI_BG  = colors.HexColor("#FEE8EA")
C_HI_TX  = colors.HexColor("#9B0006")
C_MD_BG  = colors.HexColor("#FFF3E0")
C_MD_TX  = colors.HexColor("#C66000")
C_GN_BG  = colors.HexColor("#E6F4EC")
C_GN_TX  = colors.HexColor("#00512A")
C_BL_BG  = colors.HexColor("#E8F4FD")
C_BL_TX  = colors.HexColor("#1565C0")
C_NVLT   = colors.HexColor("#AACCEE")
C_METLBL = colors.HexColor("#FFBBBB")

STATUS_STYLES = {
    "CRITICAL": (C_HI_BG, C_HI_TX),
    "Upcoming": (C_BL_BG, C_BL_TX),
    "Planned":  (C_LGRAY, C_MUTED),
    "Complete": (C_GN_BG, C_GN_TX),
}

# ─────────────────────────────────────────────────────────────────────────────
# PAGE GEOMETRY
# ─────────────────────────────────────────────────────────────────────────────

W, H = A4
LM   = 22
RM   = 22
TM   = 22
BM   = 22
CW   = W - LM - RM


# ─────────────────────────────────────────────────────────────────────────────
# DRAWING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def section_header(c, y, label, bg=C_NAVY, height=16):
    c.setFillColor(bg)
    c.rect(LM, y - height, CW, height, fill=1, stroke=0)
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(LM + 6, y - height + 4, label.upper())
    return height + 3


def table_rows(c, y, headers, rows, col_widths, row_h=15,
               hdr_bg=C_NAVY, status_col=None, bold_last=False):
    hdr_h = row_h + 1
    x0 = [LM]
    for w in col_widths[:-1]:
        x0.append(x0[-1] + w)

    c.setFillColor(hdr_bg)
    c.rect(LM, y - hdr_h, CW, hdr_h, fill=1, stroke=0)
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 7)
    for h_, x_ in zip(headers, x0):
        c.drawString(x_ + 3, y - hdr_h + 4, str(h_))

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
                sc        = str(cell)
                bg_s, tx_s = STATUS_STYLES.get(sc, (C_LGRAY, C_MUTED))
                pw        = col_widths[ci] - 6
                c.setFillColor(bg_s)
                c.roundRect(x_ + 2, ry + 2.5, pw, row_h - 5, 2, fill=1, stroke=0)
                c.setFillColor(tx_s)
                c.setFont("Helvetica-Bold", 6.5)
                c.drawCentredString(x_ + pw / 2 + 2, ry + 4, sc)
            else:
                c.setFont("Helvetica-Bold" if (bold_last and last) else "Helvetica", 7)
                c.setFillColor(C_NAVY if (bold_last and last) else C_MID)
                c.drawString(x_ + 3, ry + 4, str(cell))

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
    y = H - TM

    # ── HEADER ───────────────────────────────────────────────────────────────
    HDR_H = 52
    c.setFillColor(C_NAVY)
    c.rect(LM, y - HDR_H, CW, HDR_H, fill=1, stroke=0)
    c.setFillColor(C_RED)
    c.rect(LM, y - HDR_H, 4, HDR_H, fill=1, stroke=0)
    if os.path.exists(LOGO_PATH):
        try:
            c.drawImage(ImageReader(LOGO_PATH), LM + 8, y - HDR_H + 7,
                        width=62, height=40, preserveAspectRatio=True, mask="auto")
        except Exception:
            pass
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(W / 2, y - 21, PROJECT_TITLE)
    c.setFont("Helvetica", 7.5)
    c.setFillColor(C_NVLT)
    c.drawCentredString(W / 2, y - 34, PROJECT_SUBTITLE)
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
        ("PROGRAMME STATUS",  auto_status(DATES["day1"])),
        ("DURATION",          DURATION),
        ("DAY 1 GO-LIVE",     "01 Dec 2026"),
        ("SCOPE",             SCOPE_TEXT),
        ("LABOUR BUDGET",     LABOUR_BUDGET),
        ("DAYS TO DAY 1",     fmt_days(DATES["day1"])),
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

    # ── COUNTDOWN STRIP ──────────────────────────────────────────────────────
    CTD_H = 22
    c.setFillColor(C_NAVY)
    c.rect(LM, y - CTD_H, CW, CTD_H, fill=1, stroke=0)
    gates = [
        ("QG0", DATES["qg0"]),
        ("QG1", DATES["qg1"]),
        ("QG2", DATES["qg2"]),
        ("QG3", DATES["qg3"]),
        ("QG4 / Pre-Cutover", DATES["qg4"]),
        ("Day 1 = GoLive", DATES["day1"]),
        ("Hypercare Milestone", DATES["stab_milestone"]),
        ("QG5 Closure", DATES["qg5"]),
    ]
    gw = CW / len(gates)
    for gi, (glbl, giso) in enumerate(gates):
        gx = LM + gi * gw
        if gi > 0:
            c.setStrokeColor(colors.HexColor("#0066AA"))
            c.setLineWidth(0.4)
            c.line(gx, y - CTD_H + 3, gx, y - 3)
        gd = days_to(giso)
        gval = "Complete" if gd < 0 else f"{gd}d"
        gcol = C_RED if ("Day 1" in glbl or gd <= 30) else C_NVLT
        c.setFont("Helvetica", 5.5)
        c.setFillColor(C_NVLT)
        c.drawCentredString(gx + gw / 2, y - 8, glbl)
        c.setFont("Helvetica-Bold", 8)
        c.setFillColor(gcol)
        c.drawCentredString(gx + gw / 2, y - 18, gval)
    y -= (CTD_H + 5)

    # ── PHASE STATUS & TIMELINE ───────────────────────────────────────────────
    y -= section_header(c, y, "Phase Status & Timeline")
    ph_cols = [140, 82, 52, CW - 140 - 82 - 52]
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

    c.setFillColor(colors.HexColor("#003355"))
    c.rect(LM, y - sub_h, col_w, sub_h, fill=1, stroke=0)
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica-Bold", 7)
    c.drawString(LM + 4, y - sub_h + 4, "TOP RISKS  (P\u00d7I Rating)")

    c.setFillColor(colors.HexColor("#063333"))
    c.rect(rx, y - sub_h, col_w, sub_h, fill=1, stroke=0)
    c.setFillColor(C_WHITE)
    c.drawString(rx + 4, y - sub_h + 4, f"IMMEDIATE ACTIONS  (by {next_d})")

    decisions = [
        "Establish Steering Committee & appoint all WS leads",
        "Issue WAN RFQ — all 17 sites (EMEA priority first)",
        "Engage SAP migration partner — shortlist by Apr 30",
        f"Confirm budget baseline process — lock at QG1 (30 Jun 2026)",
        "Initiate legal review — buyer identity confidentiality plan",
    ]
    for ri in range(n):
        ry        = y - sub_h - (ri + 1) * rrow_h
        lbl, rating, lvl = TOP_RISKS[ri]
        bg_r      = C_HI_BG if lvl == "HIGH" else C_MD_BG
        tx_r      = C_HI_TX if lvl == "HIGH" else C_MD_TX
        c.setFillColor(bg_r)
        c.rect(LM, ry, col_w, rrow_h, fill=1, stroke=0)
        c.setStrokeColor(tx_r); c.setLineWidth(3)
        c.line(LM, ry, LM, ry + rrow_h)
        c.setStrokeColor(C_LINE); c.setLineWidth(0.3)
        c.rect(LM, ry, col_w, rrow_h, fill=0, stroke=1)
        c.setFillColor(tx_r)
        c.setFont("Helvetica-Bold", 7)
        c.drawString(LM + 6, ry + 5, f"({rating} {lvl})  {lbl}")

        c.setFillColor(C_WHITE)
        c.rect(rx, ry, col_w, rrow_h, fill=1, stroke=0)
        c.setStrokeColor(C_NAVY); c.setLineWidth(3)
        c.line(rx, ry, rx, ry + rrow_h)
        c.setStrokeColor(C_LINE); c.setLineWidth(0.3)
        c.rect(rx, ry, col_w, rrow_h, fill=0, stroke=1)
        c.setFillColor(C_NAVY)
        c.setFont("Helvetica", 7)
        c.drawString(rx + 6, ry + 5, decisions[ri])

    y -= (sub_h + n * rrow_h + 4)

    # ── BUDGET SUMMARY ───────────────────────────────────────────────────────
    y -= section_header(c, y, "Budget Summary (Labour — TBC at QG1)", bg=C_BLUE)
    bud_cols = [170, CW - 170 - 70, 70]
    bud_hdrs = ["Cost Category", "Estimated Amount", "% of Total"]
    y -= (table_rows(c, y, bud_hdrs, BUDGET_ROWS, bud_cols,
                     row_h=14, bold_last=True) + 4)

    # ── PROGRAMME NOTES ───────────────────────────────────────────────────────
    NOTE_H = 10 + len(PROGRAMME_NOTES) * 11 + 4
    y -= section_header(c, y, "Programme Notes & Open Decisions", bg=C_TEAL)
    c.setFillColor(C_LGRAY)
    c.rect(LM, y - NOTE_H, CW, NOTE_H, fill=1, stroke=0)
    c.setStrokeColor(C_LINE); c.setLineWidth(0.4)
    c.rect(LM, y - NOTE_H, CW, NOTE_H, fill=0, stroke=1)
    ny = y - 8
    for note in PROGRAMME_NOTES:
        c.setFillColor(C_RED); c.setFont("Helvetica-Bold", 9)
        c.drawString(LM + 6, ny, "\u25ba")
        c.setFillColor(C_MID); c.setFont("Helvetica", 7)
        c.drawString(LM + 16, ny, note)
        ny -= 11
    y -= (NOTE_H + 4)

    # ── FOOTER ───────────────────────────────────────────────────────────────
    FOOT_H = 18
    c.setFillColor(C_NAVY)
    c.rect(LM, BM, CW, FOOT_H, fill=1, stroke=0)
    c.setFillColor(C_RED)
    c.rect(LM, BM, 4, FOOT_H, fill=1, stroke=0)
    c.setFillColor(C_WHITE)
    c.setFont("Helvetica", 6.5)
    c.drawString(LM + 10, BM + 6,
                 f"Project Hamburger  |  {SELLER} \u2192 {BUYER}  |  "
                 f"Solar Energy Business  |  Stand Alone  |  {MODEL}")
    c.setFont("Helvetica-Bold", 6.5)
    c.drawRightString(W - RM - 4, BM + 6, "CONFIDENTIAL")

    c.save()
    print(f"[OK] Report written to: {OUT_PATH}")


if __name__ == "__main__":
    from datetime import datetime as _dt
    _t0 = _dt.now()
    print(f"Started : {_t0.strftime('%Y-%m-%d %H:%M:%S')}")
    try:
        build_report()
    finally:
        _t1 = _dt.now()
        print(f"Finished: {_t1.strftime('%Y-%m-%d %H:%M:%S')}  ({(_t1-_t0).total_seconds():.1f}s elapsed)")
