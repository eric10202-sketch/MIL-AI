#!/usr/bin/env python3
"""
generate_technical_paper.py

Generates: AI-Assisted_PM_Technical_Paper.pdf

Technical paper on AI-Assisted Project Management for IT Carve-Outs.
All engagement-specific project names, buyer/seller identities, and
personnel names have been removed in compliance with confidentiality
obligations.

Usage:
    C:/Program Files/px/python.exe generate_technical_paper.py
"""

import sys, os, datetime

sys.path.insert(0, os.path.join(os.path.expanduser("~"), "py_packages"))

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

TODAY      = datetime.date.today()
PUB_DATE   = TODAY.strftime("%B %Y")
HERE       = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH  = os.path.join(HERE, "Bosch.png")
OUT_PATH   = os.path.join(HERE, "AI-Assisted_PM_Technical_Paper.pdf")

W, H  = A4
LM    = 32   # left margin
RM    = 32   # right margin
TM    = 32   # top margin
BM    = 32   # bottom margin
CW    = W - LM - RM

# ─────────────────────────────────────────────────────────────────────────────
# COLOUR PALETTE  (Bosch blue theme)
# ─────────────────────────────────────────────────────────────────────────────

C_HERO    = colors.HexColor("#003b6e")   # deep navy — headers
C_ACCENT  = colors.HexColor("#0066CC")   # Bosch mid-blue
C_SECT    = colors.HexColor("#005199")   # section bars
C_RULE    = colors.HexColor("#0077BB")   # horizontal rules
C_LGRAY   = colors.HexColor("#f4f6f9")   # body background / alt rows
C_LINE    = colors.HexColor("#ccddee")   # table grid lines
C_WHITE   = colors.white
C_BODY    = colors.HexColor("#1a1a1a")   # body text
C_MUTED   = colors.HexColor("#555555")   # captions / secondary text
C_HILIGHT = colors.HexColor("#e8f0fa")   # light blue row highlight
C_TBLHDR  = colors.HexColor("#003b6e")   # table header rows
C_AMBER   = colors.HexColor("#E8A000")
C_RED     = colors.HexColor("#CC0000")
C_GREEN   = colors.HexColor("#007A33")

FONT_BODY  = "Helvetica"
FONT_BOLD  = "Helvetica-Bold"
FONT_OBLI  = "Helvetica-Oblique"

# ─────────────────────────────────────────────────────────────────────────────
# LOW-LEVEL DRAWING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

class Doc:
    """Wrapper around reportlab canvas that tracks the cursor and handles
    page breaks automatically."""

    def __init__(self, path):
        self.c   = canvas.Canvas(path, pagesize=A4)
        self.y   = H - TM
        self.page = 1
        self._new_page_callbacks = []   # functions called after each new page
        self._draw_footer()

    # ── cursor helpers ───────────────────────────────────────────────────────

    def advance(self, pts):
        """Move cursor down by pts; trigger page break if needed."""
        self.y -= pts
        if self.y < BM + 60:
            self.new_page()

    def need(self, pts):
        """Ensure at least pts points remain; break if not."""
        if self.y - pts < BM + 60:
            self.new_page()

    # ── page management ──────────────────────────────────────────────────────

    def new_page(self):
        self._draw_footer()
        self.c.showPage()
        self.page += 1
        self.y = H - TM
        self._draw_page_header_bar()

    def _draw_footer(self):
        c = self.c
        c.setStrokeColor(C_RULE)
        c.setLineWidth(0.5)
        c.line(LM, BM + 12, W - RM, BM + 12)
        c.setFont(FONT_BODY, 6.5)
        c.setFillColor(C_MUTED)
        c.drawString(LM, BM + 4,
                     "AI-Assisted Project Management for IT Carve-Outs  |  Bosch Group — Internal")
        c.drawRightString(W - RM, BM + 4, f"Page {self.page}  |  {PUB_DATE}")

    def _draw_page_header_bar(self):
        """Thin navy rule at top of continuation pages."""
        c = self.c
        c.setFillColor(C_HERO)
        c.rect(LM, H - TM - 8, CW, 8, fill=1, stroke=0)
        c.setFillColor(C_WHITE)
        c.setFont(FONT_BOLD, 6)
        c.drawString(LM + 5, H - TM - 5.5,
                     "AI-ASSISTED PROJECT MANAGEMENT FOR IT CARVE-OUTS")

    # ── text primitives ──────────────────────────────────────────────────────

    def text(self, s, font=FONT_BODY, size=8.5, color=C_BODY,
             indent=0, center=False, right=False):
        c = self.c
        c.setFont(font, size)
        c.setFillColor(color)
        x = LM + indent
        if center:
            c.drawCentredString(W / 2, self.y, s)
        elif right:
            c.drawRightString(W - RM, self.y, s)
        else:
            c.drawString(x, self.y, s)

    def wrapped_text(self, text_str, font=FONT_BODY, size=8.5,
                     color=C_BODY, indent=0, leading=12, max_width=None):
        """Draw wrapped paragraph; advances cursor."""
        mw   = (max_width if max_width else CW) - indent
        style = ParagraphStyle(
            "body",
            fontName=font,
            fontSize=size,
            textColor=color,
            leading=leading,
            alignment=TA_JUSTIFY,
        )
        p    = Paragraph(text_str, style)
        pw, ph = p.wrap(mw, 9999)
        self.need(ph + 4)
        p.drawOn(self.c, LM + indent, self.y - ph)
        self.y -= (ph + 4)

    def spacer(self, pts=6):
        self.advance(pts)

    # ── structural elements ──────────────────────────────────────────────────

    def section_bar(self, number, title):
        """Full-width navy section heading bar."""
        bar_h = 18
        self.need(bar_h + 16)
        self.spacer(8)
        c = self.c
        c.setFillColor(C_SECT)
        c.rect(LM, self.y - bar_h, CW, bar_h, fill=1, stroke=0)
        c.setFillColor(C_ACCENT)
        c.rect(LM, self.y - bar_h, 4, bar_h, fill=1, stroke=0)
        c.setFillColor(C_WHITE)
        c.setFont(FONT_BOLD, 8.5)
        label = f"{number}   {title.upper()}"
        c.drawString(LM + 10, self.y - bar_h + 5, label)
        self.y -= (bar_h + 8)

    def subsection(self, number, title):
        """Bold underlined sub-heading."""
        self.need(24)
        self.spacer(6)
        c = self.c
        c.setFillColor(C_HERO)
        c.setFont(FONT_BOLD, 8.5)
        label = f"{number}  {title}"
        c.drawString(LM, self.y, label)
        w = c.stringWidth(label, FONT_BOLD, 8.5)
        c.setStrokeColor(C_ACCENT)
        c.setLineWidth(0.8)
        c.line(LM, self.y - 1.5, LM + w, self.y - 1.5)
        self.y -= 14

    def bullet(self, text_str, indent=12, size=8.2):
        """Single bullet point (wraps)."""
        c = self.c
        style = ParagraphStyle(
            "bul",
            fontName=FONT_BODY,
            fontSize=size,
            textColor=C_BODY,
            leading=11.5,
            alignment=TA_LEFT,
        )
        full = f"•&nbsp;&nbsp;{text_str}"
        p    = Paragraph(full, style)
        pw, ph = p.wrap(CW - indent, 9999)
        self.need(ph + 3)
        p.drawOn(self.c, LM + indent, self.y - ph)
        self.y -= (ph + 3)

    def info_box(self, lines, bg=C_LGRAY, border=C_RULE):
        """Shaded info box with multiple text lines."""
        line_h = 12
        box_h  = len(lines) * line_h + 10
        self.need(box_h + 8)
        self.spacer(4)
        c = self.c
        c.setFillColor(bg)
        c.roundRect(LM, self.y - box_h, CW, box_h, 4, fill=1, stroke=0)
        c.setStrokeColor(border)
        c.setLineWidth(0.5)
        c.roundRect(LM, self.y - box_h, CW, box_h, 4, fill=0, stroke=1)
        c.setFont(FONT_BODY, 8)
        c.setFillColor(C_BODY)
        ty = self.y - 10
        for line in lines:
            c.setFont(FONT_BOLD if line.startswith("**") else FONT_BODY, 8)
            txt = line.replace("**", "")
            c.drawString(LM + 10, ty, txt)
            ty -= line_h
        self.y -= (box_h + 4)

    def code_block(self, lines):
        """Monospaced code block rendering."""
        lh    = 10.5
        box_h = len(lines) * lh + 10
        self.need(box_h + 8)
        self.spacer(4)
        c = self.c
        c.setFillColor(colors.HexColor("#1e1e2e"))
        c.roundRect(LM, self.y - box_h, CW, box_h, 4, fill=1, stroke=0)
        c.setFont("Courier", 7)
        c.setFillColor(colors.HexColor("#cdd6f4"))
        ty = self.y - 10
        for line in lines:
            c.drawString(LM + 10, ty, line)
            ty -= lh
        self.y -= (box_h + 4)

    def rule(self):
        c = self.c
        c.setStrokeColor(C_LINE)
        c.setLineWidth(0.4)
        c.line(LM, self.y, W - RM, self.y)
        self.y -= 6

    # ── table ────────────────────────────────────────────────────────────────

    def table(self, headers, rows, col_widths, row_h=13, zebra=True,
              status_col=None, bold_last=False, note=None):
        """Draw a bordered table. col_widths must sum to CW."""
        hdr_h     = row_h + 2
        total_rows = len(rows)
        total_h   = hdr_h + total_rows * row_h
        self.need(total_h + (14 if note else 6))
        self.spacer(4)

        c   = self.c
        x0s = [LM]
        for w in col_widths[:-1]:
            x0s.append(x0s[-1] + w)

        # header row
        c.setFillColor(C_TBLHDR)
        c.rect(LM, self.y - hdr_h, CW, hdr_h, fill=1, stroke=0)
        c.setFillColor(C_WHITE)
        c.setFont(FONT_BOLD, 7)
        for h_, x_ in zip(headers, x0s):
            c.drawString(x_ + 4, self.y - hdr_h + 4, str(h_))

        # data rows
        for ri, row in enumerate(rows):
            ry   = self.y - hdr_h - (ri + 1) * row_h
            last = ri == total_rows - 1
            if bold_last and last:
                bg = colors.HexColor("#dde8f0")
            elif zebra:
                bg = C_HILIGHT if ri % 2 == 0 else C_WHITE
            else:
                bg = C_WHITE
            c.setFillColor(bg)
            c.rect(LM, ry, CW, row_h, fill=1, stroke=0)
            c.setStrokeColor(C_LINE)
            c.setLineWidth(0.25)
            c.rect(LM, ry, CW, row_h, fill=0, stroke=1)

            for ci, (cell, x_) in enumerate(zip(row, x0s)):
                cell_str = str(cell)
                if bold_last and last:
                    c.setFont(FONT_BOLD, 7)
                    c.setFillColor(C_HERO)
                else:
                    c.setFont(FONT_BOLD if ci == 0 else FONT_BODY, 7)
                    c.setFillColor(C_BODY)

                if status_col is not None and ci == status_col:
                    sc_map = {
                        "HIGH":   (colors.HexColor("#fde8e8"), C_RED),
                        "MEDIUM": (colors.HexColor("#fff5e0"), C_AMBER),
                        "LOW":    (colors.HexColor("#e6f4ec"), C_GREEN),
                        "YES":    (colors.HexColor("#e6f4ec"), C_GREEN),
                        "NO":     (colors.HexColor("#fde8e8"), C_RED),
                    }
                    bg_s, tx_s = sc_map.get(cell_str.upper(), (C_LGRAY, C_MUTED))
                    pw = col_widths[ci] - 6
                    c.setFillColor(bg_s)
                    c.roundRect(x_ + 2, ry + 2, pw, row_h - 4, 2, fill=1, stroke=0)
                    c.setFillColor(tx_s)
                    c.setFont(FONT_BOLD, 6.5)
                    c.drawCentredString(x_ + pw / 2 + 2, ry + 3.5, cell_str)
                else:
                    # Wrap long cells
                    style = ParagraphStyle(
                        "tc",
                        fontName=FONT_BOLD if (ci == 0 or (bold_last and last)) else FONT_BODY,
                        fontSize=7,
                        textColor=C_HERO if (bold_last and last) else C_BODY,
                        leading=8.5,
                    )
                    pw   = col_widths[ci] - 6
                    para = Paragraph(cell_str, style)
                    pw2, ph2 = para.wrap(pw, row_h)
                    para.drawOn(c, x_ + 4, ry + (row_h - ph2) / 2)

        # outer border
        c.setStrokeColor(C_RULE)
        c.setLineWidth(0.6)
        c.rect(LM, self.y - hdr_h - total_rows * row_h, CW, total_h, fill=0, stroke=1)

        self.y -= (total_h + 4)

        if note:
            c.setFont(FONT_OBLI, 6.5)
            c.setFillColor(C_MUTED)
            c.drawString(LM + 2, self.y, note)
            self.y -= 10

    def save(self):
        self._draw_footer()
        self.c.save()


# ─────────────────────────────────────────────────────────────────────────────
# COVER PAGE
# ─────────────────────────────────────────────────────────────────────────────

def draw_cover(doc):
    c = doc.c

    # Full-page gradient-like hero band
    c.setFillColor(C_HERO)
    c.rect(0, H - 220, W, 220, fill=1, stroke=0)

    # Accent stripe
    c.setFillColor(C_ACCENT)
    c.rect(0, H - 220, W, 6, fill=1, stroke=0)

    # Bosch logo on cover
    if os.path.exists(LOGO_PATH):
        try:
            logo_reader = ImageReader(LOGO_PATH)
            c.drawImage(logo_reader, LM, H - TM - 52,
                        width=80, height=50,
                        preserveAspectRatio=True, mask="auto")
        except Exception:
            pass

    # Classification badge
    c.setFillColor(C_AMBER)
    c.roundRect(W - RM - 100, H - TM - 30, 100, 18, 4, fill=1, stroke=0)
    c.setFillColor(colors.HexColor("#1a1a1a"))
    c.setFont(FONT_BOLD, 7)
    c.drawCentredString(W - RM - 50, H - TM - 18, "INTERNAL — BOSCH GROUP")

    # Title block
    c.setFillColor(C_WHITE)
    c.setFont(FONT_BOLD, 22)
    c.drawCentredString(W / 2, H - 95, "AI-Assisted Project Management")
    c.setFont(FONT_BOLD, 16)
    c.drawCentredString(W / 2, H - 118, "for IT Carve-Outs")

    c.setFillColor(colors.HexColor("#AACCEE"))
    c.setFont(FONT_BODY, 9.5)
    c.drawCentredString(W / 2, H - 140,
        "Architecture, Methodology and Tooling for AI-Driven Deliverable Generation")

    # Meta strip below hero
    c.setFillColor(C_LGRAY)
    c.rect(0, H - 280, W, 54, fill=1, stroke=0)
    c.setStrokeColor(C_RULE)
    c.setLineWidth(0.5)
    c.line(0, H - 227, W, H - 227)
    c.line(0, H - 280, W, H - 280)

    meta = [
        ("Practice", "Mergers, Integrations & Liquidations (MIL)"),
        ("Organisation", "Bosch Global Business Services — IT Strategy"),
        ("Published", PUB_DATE),
        ("Classification", "Internal — Bosch Group"),
    ]
    cx = LM + 10
    cy = H - 244
    for label, val in meta:
        c.setFont(FONT_BOLD, 7.5)
        c.setFillColor(C_HERO)
        c.drawString(cx, cy, label + ":")
        c.setFont(FONT_BODY, 7.5)
        c.setFillColor(C_BODY)
        c.drawString(cx + 72, cy, val)
        cy -= 13

    # Abstract box
    c.setFillColor(C_WHITE)
    c.roundRect(LM, H - 440, CW, 148, 5, fill=1, stroke=0)
    c.setStrokeColor(C_RULE)
    c.setLineWidth(0.8)
    c.roundRect(LM, H - 440, CW, 148, 5, fill=0, stroke=1)
    c.setFillColor(C_SECT)
    c.roundRect(LM, H - 301, CW, 14, 5, fill=1, stroke=0)
    c.setFillColor(C_WHITE)
    c.setFont(FONT_BOLD, 7)
    c.drawString(LM + 10, H - 310, "ABSTRACT")

    abstract = (
        "This paper describes the design, architecture, and operational practice of an AI-assisted project-management "
        "toolkit developed within Bosch's MIL practice. The toolkit uses GitHub Copilot (powered by Claude Sonnet) "
        "running inside Visual Studio Code in agent mode to automate generation of all mandatory IT carve-out "
        "deliverables — project schedule, risk register, cost plan, project charter, executive dashboard, management "
        "KPI dashboard, and monthly status reports. "
        "The system combines a structured repository, a domain-specific instruction layer, modular skill files that "
        "encode methodology logic, Python generator scripts, and a persistent memory framework. The result is a "
        "significant reduction in manual document-production effort while enforcing methodological consistency "
        "across all engagements. All project-specific identifiers have been removed from this paper in compliance "
        "with confidentiality obligations."
    )
    abs_style = ParagraphStyle("abs",
        fontName=FONT_BODY, fontSize=8, textColor=C_BODY,
        leading=12, alignment=TA_JUSTIFY)
    p = Paragraph(abstract, abs_style)
    pw, ph = p.wrap(CW - 20, 9999)
    p.drawOn(c, LM + 10, H - 430)

    # Cover footer
    c.setStrokeColor(C_LINE)
    c.setLineWidth(0.5)
    c.line(LM, BM + 26, W - RM, BM + 26)
    c.setFont(FONT_BODY, 6.5)
    c.setFillColor(C_MUTED)
    c.drawString(LM, BM + 16,
        "Confidential business methodology paper. Project identifiers removed per confidentiality obligations.")
    c.drawRightString(W - RM, BM + 16, f"Published: {PUB_DATE}")
    c.drawString(LM, BM + 6,
        "© Bosch Group — For internal circulation only. Not for external distribution.")

    doc.c.showPage()
    doc.page += 1
    doc.y = H - TM
    doc._draw_page_header_bar()


# ─────────────────────────────────────────────────────────────────────────────
# TABLE OF CONTENTS
# ─────────────────────────────────────────────────────────────────────────────

TOC = [
    ("1",   "Background and Motivation",                      3),
    ("2",   "Platform Architecture",                          3),
    ("2.1", "Tool Stack",                                     3),
    ("2.2", "Repository Structure",                           4),
    ("3",   "The AI Instruction Layer",                       4),
    ("3.1", "Workspace-Level Guardrails",                     4),
    ("3.2", "CLAUDE.md — Redundant Guardrail",                5),
    ("3.3", "Skill Files — Domain Knowledge Encoding",        5),
    ("4",   "The Compliance Gate",                            5),
    ("5",   "Schedule Generation",                            6),
    ("5.1", "Two-File Output Contract",                       6),
    ("5.2", "The Canonical XML Generator",                    6),
    ("5.3", "Per-Project Schedule Generator Pattern",         7),
    ("6",   "Risk Register Generation",                       7),
    ("7",   "Cost Plan Generation",                           8),
    ("8",   "HTML Deliverables",                              8),
    ("8.1", "Design Principles",                              8),
    ("8.2", "Executive Dashboard Layout",                     9),
    ("8.3", "Management KPI Dashboard",                       9),
    ("9",   "Monthly Status Report Generation",               10),
    ("10",  "Memory Architecture",                            10),
    ("10.1","Three-Tier Memory",                              10),
    ("10.2","Memory as Institutional Knowledge",              11),
    ("11",  "Generator Script Patterns",                      11),
    ("11.1","Timing Wrapper",                                 11),
    ("11.2","Subprocess Chaining",                            12),
    ("12",  "Deliverable Orchestration",                      12),
    ("13",  "Version Control and Collaboration",              13),
    ("14",  "Observed Outcomes and Lessons",                  13),
    ("15",  "Future Directions",                              14),
    ("16",  "Conclusion",                                     15),
    ("A",   "Appendix A — Mandatory Engagement Intake Fields",16),
    ("B",   "Appendix B — Deliverable Generation Sequence",   16),
    ("C",   "Appendix C — Generator Scripts Inventory",       17),
]

def draw_toc(doc):
    doc.need(400)
    c = doc.c
    # TOC header
    c.setFillColor(C_SECT)
    c.rect(LM, doc.y - 18, CW, 18, fill=1, stroke=0)
    c.setFillColor(C_ACCENT)
    c.rect(LM, doc.y - 18, 4, 18, fill=1, stroke=0)
    c.setFillColor(C_WHITE)
    c.setFont(FONT_BOLD, 9)
    c.drawString(LM + 10, doc.y - 13, "TABLE OF CONTENTS")
    doc.y -= 26

    for num, title, pg in TOC:
        doc.need(15)
        indent = 14 if "." in num or (len(num) == 1 and num.isalpha()) else 0
        is_main = ("." not in num) and not (len(num) == 1 and num.isalpha() and num in "ABC")
        is_app  = len(num) == 1 and num.isalpha()

        font   = FONT_BOLD if (is_main or is_app) else FONT_BODY
        fsize  = 8.5 if (is_main or is_app) else 7.5
        color  = C_HERO if (is_main or is_app) else C_BODY

        c.setFont(font, fsize)
        c.setFillColor(color)
        label = f"{num}  {title}"
        c.drawString(LM + indent, doc.y, label)

        # dot leaders
        lw = c.stringWidth(label, font, fsize)
        pg_text = str(pg)
        pw = c.stringWidth(pg_text, font, fsize)
        dot_start = LM + indent + lw + 4
        dot_end   = W - RM - pw - 6
        if dot_end > dot_start:
            c.setFont(FONT_BODY, 6)
            c.setFillColor(C_MUTED)
            x = dot_start
            while x < dot_end:
                c.drawString(x, doc.y, ".")
                x += 5

        c.setFont(font, fsize)
        c.setFillColor(color)
        c.drawRightString(W - RM, doc.y, pg_text)
        doc.y -= (13 if is_main else 11)

    doc.spacer(10)
    doc.rule()


# ─────────────────────────────────────────────────────────────────────────────
# CONTENT SECTIONS
# ─────────────────────────────────────────────────────────────────────────────

def section_1(doc):
    doc.section_bar("1", "Background and Motivation")
    doc.wrapped_text(
        "IT carve-outs are among the most document-intensive activities in mergers and acquisitions. "
        "A typical engagement requires a project manager to produce and keep in sync seven major artifacts — "
        "project schedule, risk register, cost plan, project charter, two HTML dashboards, and a monthly "
        "report — each with interdependencies, complex formatting requirements, and data derived from the "
        "same core scope facts (buyer, seller, sites, users, applications)."
    )
    doc.spacer(5)
    doc.wrapped_text(
        "Producing these documents manually is time-consuming, error-prone, and inconsistent across "
        "engagements. Small differences in assumptions — a phase date that drifts between the schedule and "
        "the cost plan, a resource label that changes, a cost contingency line not reconciled with the risk "
        "register — compound into governance problems at quality gates."
    )
    doc.spacer(5)
    doc.wrapped_text(
        "The initiative described here arose from real project experience across multiple IT carve-out "
        "engagements of varying scope — from focused 70-user JV separations to large 2,600-user multi-site "
        "Stand Alone carve-outs. The goal was to build a toolkit where a project manager could supply the "
        "eleven mandatory engagement parameters and receive a complete, consistent, methodologically correct "
        "set of deliverables within minutes — with all cross-checks enforced automatically, not by human review."
    )


def section_2(doc):
    doc.section_bar("2", "Platform Architecture")

    doc.subsection("2.1", "Tool Stack")
    doc.wrapped_text(
        "The toolkit runs on a combination of commercially available and enterprise-grade tools. "
        "All components are either open-source or covered by existing Bosch enterprise licences."
    )
    doc.spacer(4)
    doc.table(
        headers=["Layer", "Tool", "Role"],
        rows=[
            ("AI Inference",        "GitHub Copilot (Claude Sonnet 4.6)",  "Reasoning, code generation, content synthesis"),
            ("IDE",                 "Visual Studio Code",                   "Primary workspace and chat interface"),
            ("Version Control",     "Git + GitHub (Bosch DevCloud)",        "Repository storage, history, team sharing"),
            ("Python Runtime",      "CPython 3.x",                         "Script execution and file generation"),
            ("Document Generation", "Python stdlib + openpyxl + reportlab", "CSV, MSPDI XML, HTML, PDF output"),
            ("Collaboration",       "Microsoft SharePoint / OneDrive",      "File distribution and offline HTML rendering"),
        ],
        col_widths=[80, 130, CW - 210],
        row_h=14,
    )

    doc.subsection("2.2", "Repository Structure")
    doc.wrapped_text(
        "The repository follows a four-zone content model that enforces strict separation between "
        "methodology reference material, reusable templates, active engagement deliverables, and archived work."
    )
    doc.spacer(4)
    doc.code_block([
        "Carveout/",
        "├── .claude/skills/               ← Domain-skill instruction files (8 SKILL.md files)",
        "├── .github/copilot-instructions.md  ← Auto-loaded workspace guardrails",
        "├── references/                   ← Methodology-only assets (read-only for AI)",
        "├── templates/                    ← Blank project-agnostic templates",
        "├── active-projects/              ← Current engagement metadata",
        "├── archive/                      ← Closed engagement artifacts",
        "├── {ProjectName}/                ← Per-project output folder",
        "├── generate_msp_xml.py           ← Canonical shared CSV→XML converter",
        "├── generate_{project}_schedule.py   ← Per-project generator scripts",
        "├── generate_{project}_risk_register.py",
        "├── generate_{project}_monthly_report.py",
        "├── Bosch.png                     ← Logo asset (base64-embedded in HTML/PDF)",
        "└── CLAUDE.md                     ← Always-on AI guardrails",
    ])
    doc.spacer(4)
    doc.bullet("references/ and archive/ are methodology-only. The AI is instructed to never copy "
               "reference party names, dates, or scope into active engagement deliverables.")
    doc.bullet("Every project output lives in its own named folder, isolated from all other engagements.")
    doc.bullet("Generator scripts are project-specific — one script per engagement — so parameters "
               "are hardcoded and auditable, not interpolated at runtime from unvalidated input.")


def section_3(doc):
    doc.section_bar("3", "The AI Instruction Layer")

    doc.subsection("3.1", "Workspace-Level Guardrails (copilot-instructions.md)")
    doc.wrapped_text(
        "Visual Studio Code's GitHub Copilot extension automatically loads .github/copilot-instructions.md "
        "for any workspace that contains it. This file is the always-on constitution of the AI's behaviour "
        "in this repository. It contains four categories of rules:"
    )
    doc.spacer(4)
    doc.bullet("<b>Mandatory intake fields:</b> The AI is blocked from generating any deliverable unless "
               "all eleven engagement parameters are confirmed (see Appendix A).")
    doc.bullet("<b>Derivation rules:</b> Once seller and buyer are confirmed, several downstream fields "
               "are automatically derived — Sponsor Customer = Buyer; Sponsor Contractor = Seller; "
               "IT flow direction = Seller IT → Merger Zone (Integration model) → Buyer IT.")
    doc.bullet("<b>Deliverable orchestration sequence:</b> Schedule → Risk Register → Cost Plan → "
               "Project Charter → Executive Dashboard → Management KPI Dashboard → Monthly Status Report. "
               "Each deliverable depends on all preceding ones and the AI cannot skip steps.")
    doc.bullet("<b>Skill routing table:</b> A lookup table maps each task type to the corresponding "
               "skill file under .claude/skills/, instructing the AI to read the full skill before acting.")

    doc.subsection("3.2", "CLAUDE.md — Redundant Guardrail")
    doc.wrapped_text(
        "CLAUDE.md at the workspace root is a second copy of the global guardrails, formatted for the "
        "Claude AI's native instruction format. Having both files ensures the rules are loaded regardless "
        "of which entry point is used — native Copilot chat or direct Claude API. The two files are kept "
        "in sync as part of the repository governance routine."
    )

    doc.subsection("3.3", "Skill Files — Domain Knowledge Encoding")
    doc.wrapped_text(
        "The .claude/skills/ directory contains eight SKILL.md files, each encoding the complete "
        "methodology for one deliverable type. A skill file is not a prompt template — it is a structured "
        "procedural specification that the AI reads and executes deterministically."
    )
    doc.spacer(4)
    doc.table(
        headers=["Skill File", "Deliverable", "Key Rules Encoded"],
        rows=[
            ("intake-compliance-gate",          "Engagement validation",   "11 mandatory fields, blocking rules, budget TBC exception"),
            ("schedule-generation",             "CSV + MSPDI XML",         "Task schema, XML element ordering, date format, subprocess pattern"),
            ("cost-plan-generation",            "Labour cost plan CSV",    "3 cross-checks: schedule, risk register, resource name consistency"),
            ("risk-register-generation",        "Risk register .xlsx",     "Column schema, RZ formula, template path, openpyxl rules"),
            ("executive-dashboard-generation",  "Executive dashboard HTML","3-page Trinity layout, colour palette, logo embedding rules"),
            ("management-kpi-dashboard-generation","KPI dashboard HTML",   "12-column card grid, SPI/CPI/readiness KPIs, 90-day forecast"),
            ("monthly-status-report-generation","Status report PDF",       "A4 layout, auto-dated filename, reportlab pattern"),
            ("repository-governance-updates",   "Repo metadata",           "Last-reviewed dates, inventory sync, naming conventions"),
        ],
        col_widths=[108, 78, CW - 186],
        row_h=14,
    )
    doc.spacer(4)
    doc.wrapped_text(
        "This separation of <i>knowledge</i> (skill files) from <i>instructions</i> (copilot-instructions.md) "
        "is a deliberate architectural choice. It keeps the global instruction file short enough to be reliably "
        "loaded within the AI's context window, while allowing rich domain detail in skill files that are "
        "loaded on demand for each specific task."
    )


def section_4(doc):
    doc.section_bar("4", "The Compliance Gate")
    doc.wrapped_text(
        "The intake-compliance-gate skill operationalises a strict validation step that must pass before any "
        "document generation begins. The gate enforces a binary contract:"
    )
    doc.spacer(5)
    doc.info_box([
        "**PASS:  All 11 mandatory fields confirmed → emit validation summary → proceed to generation",
        "**FAIL:  Any field missing → block ALL generation → list every missing field in one response",
        "**RULE:  No fielding of incomplete data. No substitutions from reference projects.",
        "**EXCEPTION: If budget is explicitly unknown → substitute literal: TBC — to be approved at QG1",
    ], bg=C_HILIGHT, border=C_ACCENT)
    doc.spacer(4)
    doc.wrapped_text(
        "The buyer/seller derivation step follows immediately after a successful gate pass. The AI applies "
        "the derivation rules automatically — it does not ask for derived fields separately. This keeps the "
        "user interaction lean: eleven fields in, full engagement context out."
    )
    doc.spacer(4)
    doc.wrapped_text(
        "The compliance gate design reflects a real project risk: AI tools can generate plausible-looking "
        "content from incomplete inputs — scope inventories with estimated user counts, phase dates derived "
        "from a similar engagement, resource labels borrowed from a reference project. The gate pattern "
        "forces the AI to surface incompleteness explicitly rather than fill gaps silently."
    )


def section_5(doc):
    doc.section_bar("5", "Schedule Generation")

    doc.subsection("5.1", "Two-File Output Contract")
    doc.wrapped_text(
        "Every schedule generation produces exactly two files. The CSV is the source of truth and is "
        "human-readable, importable directly into Excel. The XML is always derived from the CSV via "
        "generate_msp_xml.py — it is never hand-written."
    )
    doc.spacer(4)
    doc.info_box([
        "OUTPUT 1:  {ProjectName}_Project_Schedule.csv   (source of truth — human readable)",
        "OUTPUT 2:  {ProjectName}_Project_Schedule.xml   (MSPDI format — derived from CSV only)",
        "RULE:  XML must always be generated via generate_msp_xml.py. Never hand-write or hardcode.",
    ])

    doc.subsection("5.2", "The Canonical XML Generator (generate_msp_xml.py)")
    doc.wrapped_text(
        "generate_msp_xml.py is a pure Python stdlib script (zero non-standard dependencies) that reads any "
        "project's CSV schedule and emits a standards-conformant Microsoft Project Data Interchange (MSPDI) "
        "XML file. Its significance lies in encoding a set of XML rules that are not clearly documented in "
        "the MS Project specification but are critical for correct import behaviour. These rules were "
        "discovered empirically through iteration on real project imports."
    )
    doc.spacer(4)
    doc.table(
        headers=["XML Rule", "Consequence if Violated"],
        rows=[
            ("<TaskMode>1</TaskMode> must appear immediately after <Name>",
             "MS Project silently ignores the manual-task flag; all dates recalculate from predecessor chain on import"),
            ("<ManualStart> and <ManualFinish> must accompany <Start> and <Finish>",
             "MS Project displays auto-calculated dates rather than the authored dates — GoLive may appear months late"),
            ("<ConstraintType>2 + <ConstraintDate> on every non-summary task",
             "Without 'Must Start On' pin, task dates drift when the file is reopened or refreshed"),
            ("Project header: <NewTasksAreManual>1 and <DefaultTaskType>1",
             "New tasks added in MS Project default to auto-scheduling, undermining date integrity"),
            ("Duration: PT{n×8}H0M0S (working hours), not ISO P{n}D",
             "ISO calendar-day duration causes MS Project to miscalculate working-day task lengths"),
        ],
        col_widths=[CW // 2, CW // 2],
        row_h=16,
        note="These rules are captured in generate_msp_xml.py, the schedule-generation SKILL.md, and the persistent memory system.",
    )

    doc.subsection("5.3", "Per-Project Schedule Generator Pattern")
    doc.wrapped_text(
        "Each engagement has its own generate_{project}_schedule.py containing the complete TASKS list "
        "as a Python tuple structure, with schema: (ID, OutlineLevel, Name, Duration, Start, Finish, "
        "Predecessors, ResourceNames, Notes, Milestone). The main block writes the CSV, then calls "
        "generate_msp_xml.py as a subprocess."
    )
    doc.spacer(4)
    doc.wrapped_text(
        "The per-project hardcoded pattern — rather than a single parametric script — is deliberate. "
        "Hardcoded task lists are auditable: a project manager can read the script and see exactly what "
        "was generated and why. There is no runtime variable substitution that could produce unexpected "
        "output from slightly different input."
    )


def section_6(doc):
    doc.section_bar("6", "Risk Register Generation")
    doc.wrapped_text(
        "Risk registers are generated as .xlsx files using Risk_analysis_template.xlsx as the base workbook. "
        "The template contains pre-built RAG formula cells and the authoritative column structure used "
        "across all Bosch MIL engagements. The openpyxl library populates cell values; Excel calculates "
        "the formulas on file open."
    )
    doc.spacer(4)
    doc.table(
        headers=["Column", "Field", "Notes"],
        rows=[
            ("A", "Risk number",          "Sequential index"),
            ("B", "Sub-project",          "Optional subdivision"),
            ("C", "Entry date",           "ISO date of risk identification"),
            ("D", "Risk category",        "Taxonomy: ScR, SR, RR, BtR, QR, BR, LR, CR"),
            ("E", "Risk description",     "Concrete, engagement-specific statement"),
            ("F", "Effects",              "Business/schedule consequence if risk materialises"),
            ("G", "Causes",               "Root cause or trigger condition"),
            ("H", "Probability (W)",      "1–5 scale"),
            ("I", "Impact (T)",           "1–5 scale"),
            ("J", "Risk Rating (RZ)",     "Formula: W × T — pre-built in template; threshold 12 = HIGH"),
            ("K", "Mitigation actions",   "Preventive and corrective measures"),
            ("L", "Owner",                "Named individual responsible"),
            ("M", "Deadline",             "Target resolution date"),
            ("N", "Action status",        "Open / In Progress / Closed"),
            ("O", "Remarks",              "Supplementary notes — NOT the same field as Effects (col F)"),
        ],
        col_widths=[22, 90, CW - 112],
        row_h=13,
        note="Column F (Effects) and Column O (Remarks) serve different purposes and must not be swapped.",
    )
    doc.spacer(4)
    doc.wrapped_text(
        "High-priority risks (RZ ≥ 12) are surfaced at script end for immediate review. The cost plan "
        "generation skill mandates a cross-check against these risks: any risk whose mitigation involves "
        "external cost must have a corresponding contingency line in the cost plan."
    )


def section_7(doc):
    doc.section_bar("7", "Cost Plan Generation")
    doc.wrapped_text(
        "The cost plan is the third deliverable in the mandatory sequence and can only be generated after "
        "both the schedule and risk register are confirmed complete. The skill enforces three mandatory "
        "pre-generation cross-checks before any cost figures are written."
    )
    doc.spacer(4)
    doc.table(
        headers=["Cross-Check", "Rule", "Consequence if Skipped"],
        rows=[
            ("1 — Schedule alignment",
             "Phase names and date ranges must match the schedule exactly. Resource names must map to schedule Resource Names column.",
             "Cost plan presents a different project story from the schedule — governance gap at QG review"),
            ("2 — Risk register alignment",
             "Every Amber/Red or RZ ≥ 12 risk whose mitigation involves external spend must have a CAPEX/contingency line referencing the risk.",
             "Budget baseline understates true cost exposure — financial risk at approval gate"),
            ("3 — Resource name consistency",
             "Each +-separated token in the schedule's resource column must be traceable to at least one cost plan line. No invented labels.",
             "Cost plan resources cannot be reconciled with schedule resource assignments"),
        ],
        col_widths=[60, CW // 2, CW - 60 - CW // 2],
        row_h=20,
    )
    doc.spacer(4)
    doc.wrapped_text(
        "The output is a structured CSV with mandatory sections: category blocks with subtotals, an overall "
        "project total, cost breakdowns by category and by phase, and a CAPEX/additional costs section "
        "(excluded from the labour total) containing risk-driven contingency lines with explicit references "
        "to the relevant risk register entries."
    )


def section_8(doc):
    doc.section_bar("8", "HTML Deliverables — Executive and KPI Dashboards")

    doc.subsection("8.1", "Design Principles")
    doc.wrapped_text(
        "Both HTML deliverables follow a strict set of non-negotiable design rules enforced globally "
        "through the instruction layer and individual skill files."
    )
    doc.spacer(4)
    doc.bullet("<b>Self-contained HTML:</b> No external CDN links, no web font requests. The file must render "
               "correctly offline — SharePoint, local file system, email attachment.")
    doc.bullet("<b>Bosch logo embedded as base64 PNG:</b> Bosch.png is read at generation time, base64-encoded, "
               "and injected as a data: URI. This eliminates broken-image problems when the file is moved.")
    doc.bullet("<b>Blue as primary colour:</b> #003b6e (deep navy) for headers/hero bands; #0066CC (Bosch "
               "mid-blue) for accents. Bosch Red (#E20015) is reserved exclusively for RAG badges and "
               "critical-path indicators.")
    doc.bullet("<b>Print-safe layout:</b> page-break-before:always on section breaks enables clean A4 PDF "
               "printing from the browser.")
    doc.bullet("<b>CSS container rule for logo:</b> .bosch-logo must use display:flex; align-items:center — "
               "display:grid or fixed dimensions clip the logo image at certain viewport sizes.")

    doc.subsection("8.2", "Executive Dashboard — Three-Page A4 Layout")
    doc.wrapped_text(
        "The canonical layout is derived from the Project Trinity Executive Dashboard (reference PDF). "
        "All new dashboards replicate this structure with engagement-specific data."
    )
    doc.spacer(4)
    doc.table(
        headers=["Page", "Sections"],
        rows=[
            ("Page 1",
             "Header band with logo and countdown · Day-1 event strip · Programme overview (narrative + carve-out model + key parties + budget + governance) · "
             "Stats row (6 metrics: sites, users, devices, apps, duration, TSA) · Phase timeline bar · Milestones and QG table · Budget distribution"),
            ("Page 2",
             "Continued milestones (if overflow) · IT Workstream Coverage (3×3 grid, WS1–WS9, confidence tags) · "
             "Quality Gate Tracker (criteria per gate) · Regional site distribution · Key risk indicators (HIGH / MEDIUM / LOW with P×I scores)"),
            ("Page 3",
             "Application migration waves (bar chart) · Country complexity hotspots grid · "
             "Stats strip (total tasks, resources, hours, DC hubs) · Critical path (4-column grid) · Footer with data sources and confidentiality notice"),
        ],
        col_widths=[38, CW - 38],
        row_h=26,
    )

    doc.subsection("8.3", "Management KPI Dashboard")
    doc.wrapped_text(
        "The KPI dashboard is the operational steering view for the PMO and SteerCo. It uses a 12-column "
        "CSS grid card layout and renders the following KPI areas: Schedule Performance Index (SPI), "
        "Cost Performance Index (CPI), Day-1 readiness, Stand Alone / TSA confidence, workstream confidence "
        "bars, milestone gate control timeline, top risk table, carve-out model key differences, and a "
        "90-day action forecast."
    )


def section_9(doc):
    doc.section_bar("9", "Monthly Status Report Generation")
    doc.wrapped_text(
        "Monthly status reports are PDF documents generated by per-project Python scripts using the "
        "reportlab library. Each script produces a single-page A4 report with: programme overview header "
        "band with embedded Bosch logo, days-to-gate countdown strip, phase RAG status table, risk summary, "
        "budget burn, programme notes, and upcoming actions."
    )
    doc.spacer(5)
    doc.wrapped_text(
        "The output filename includes the runtime month and year "
        "({ProjectName}_Monthly_Status_Report_{MMM_YYYY}.pdf), so the script auto-dates itself on each "
        "run. No manual filename editing is needed. All days-to-gate values are calculated from "
        "datetime.date.today() at runtime, not hardcoded."
    )


def section_10(doc):
    doc.section_bar("10", "Memory Architecture")

    doc.subsection("10.1", "Three-Tier Memory")
    doc.wrapped_text(
        "The toolkit uses a three-tier memory system that persists knowledge beyond individual "
        "conversation sessions. Each tier has a different scope and lifetime."
    )
    doc.spacer(4)
    doc.table(
        headers=["Tier", "Scope", "Lifetime", "Contents"],
        rows=[
            ("User Memory\n(/memories/)",
             "Cross-workspace",
             "Permanent",
             "User preferences, common patterns, learned rules — e.g. XML element ordering rules, logo CSS rules"),
            ("Session Memory\n(/memories/session/)",
             "Current conversation",
             "Session only",
             "Task-specific working state, in-progress notes, intermediate results"),
            ("Repository Memory\n(/memories/repo/)",
             "This repository",
             "Repository-scoped",
             "GitHub remote URL, canonical branch, Python interpreter path, repo facts"),
        ],
        col_widths=[64, 66, 50, CW - 180],
        row_h=22,
    )

    doc.subsection("10.2", "Memory as Institutional Knowledge")
    doc.wrapped_text(
        "The memory system serves a function analogous to onboarding documentation in a traditional team. "
        "When a new AI session begins, reading the memory is equivalent to a new team member reading the "
        "project handbook. The difference is that AI memory is precise (stored as structured rules, "
        "not narrative prose), always consulted (not skipped under time pressure), and always current "
        "(updated immediately when a new lesson is learned)."
    )
    doc.spacer(5)
    doc.wrapped_text(
        "Each time a significant lesson is learned — the XML <TaskMode> element ordering rule, "
        "the .bosch-logo CSS display:grid clipping bug, the openpyxl .xlsx-only constraint — it is "
        "written into the user memory system and enforced from that point forward. The system becomes "
        "more reliable with each engagement, rather than regressing to previously resolved bugs."
    )
    doc.spacer(5)
    doc.info_box([
        "EXAMPLE — User memory entry (carveout-output-standards.md):",
        "  · .bosch-logo CSS: display:flex; align-items:center — do NOT use display:grid or fixed width/height",
        "  · HTML colour theme: Blue primary (#003b6e hero, #0066CC accent). Red only for RAG badges.",
        "  · Logo: embed Bosch.png as data:image/png;base64,... at height:36px",
    ], bg=C_HILIGHT)


def section_11(doc):
    doc.section_bar("11", "Generator Script Patterns")

    doc.subsection("11.1", "Timing Wrapper")
    doc.wrapped_text(
        "All generator scripts follow a consistent __main__ pattern that records and displays start time, "
        "finish time, and elapsed duration. The try/finally structure ensures the timing line always "
        "prints, even when the generation step raises an exception — so the operator always knows how "
        "long the run took and whether it completed."
    )
    doc.spacer(4)
    doc.code_block([
        "if __name__ == '__main__':",
        "    from datetime import datetime as _dt",
        "    _t0 = _dt.now()",
        "    print(f'Started : {_t0.strftime(\"%Y-%m-%d %H:%M:%S\")}')",
        "    try:",
        "        main()",
        "    finally:",
        "        _t1 = _dt.now()",
        "        elapsed = (_t1 - _t0).total_seconds()",
        "        print(f'Finished: {_t1.strftime(\"%Y-%m-%d %H:%M:%S\")}  ({elapsed:.1f}s elapsed)')",
    ])
    doc.spacer(4)
    doc.wrapped_text(
        "This pattern was applied uniformly across all fourteen generator scripts in a single AI-assisted "
        "refactoring operation, demonstrating another use case of the toolkit: bulk code modifications "
        "guided by natural language instruction."
    )

    doc.subsection("11.2", "Subprocess Chaining")
    doc.wrapped_text(
        "Schedule generators call generate_msp_xml.py as a subprocess, not by importing it as a module. "
        "This design keeps the XML generation logic centralised in one place while allowing each "
        "project's schedule script to invoke it. Improvements to the XML generator (such as the "
        "TaskMode element ordering fix) apply immediately to all projects without any per-project "
        "script changes."
    )
    doc.spacer(4)
    doc.code_block([
        "result = subprocess.run(",
        "    [sys.executable,",
        "     str(HERE / 'generate_msp_xml.py'),",
        "     '--csv',     str(CSV_PATH),",
        "     '--out',     str(XML_PATH),",
        "     '--project', 'Project {Name}'],",
        "    capture_output=True, text=True",
        ")",
        "if result.returncode != 0:",
        "    print(result.stderr); sys.exit(1)",
    ])


def section_12(doc):
    doc.section_bar("12", "Deliverable Orchestration")
    doc.wrapped_text(
        "The deliverable generation sequence is enforced at the AI instruction level — it is not merely "
        "a recommendation. The ordering exists because each deliverable has hard data dependencies on "
        "its predecessors."
    )
    doc.spacer(4)
    doc.table(
        headers=["Step", "Deliverable", "Depends On", "Key Dependency Rationale"],
        rows=[
            ("1", "Schedule (CSV + XML)",       "—",              "Foundation for all downstream deliverables"),
            ("2", "Risk Register (.xlsx)",       "Schedule",       "Risks must reference schedule phases and quality gates"),
            ("3", "Cost Plan (.csv)",            "Schedule + Risk","Resources from schedule; contingency lines from risk register"),
            ("4", "Project Charter (.html)",     "Cost Plan",      "Charter presents approved schedule, budget baseline, and risks"),
            ("5", "Executive Dashboard (.html)", "Cost Plan + Risk","Dashboard visualises data from all three source deliverables"),
            ("6", "Management KPI Dashboard (.html)", "Cost Plan + Risk", "SPI/CPI/readiness metrics derived from schedule and cost"),
            ("7", "Monthly Status Report (.pdf)","All above",      "Snapshot of current state across all sources"),
        ],
        col_widths=[20, 110, 80, CW - 210],
        row_h=14,
    )
    doc.spacer(4)
    doc.wrapped_text(
        "If any mandatory predecessor is missing when a downstream deliverable is requested, the AI "
        "blocks generation and reports the missing dependency explicitly. This prevents a common failure "
        "mode where dashboards or charters are generated with placeholder or stale data."
    )


def section_13(doc):
    doc.section_bar("13", "Version Control and Collaboration")
    doc.wrapped_text(
        "Generated deliverables and generator scripts are both tracked in Git. The repository is hosted "
        "on Bosch DevCloud, with change sharing managed through GitHub. The standard workflow is:"
    )
    doc.spacer(4)
    doc.bullet("Open workspace in VS Code.")
    doc.bullet("Use Copilot Chat in Agent mode to generate or update deliverables.")
    doc.bullet("Review generated files for engagement-specific accuracy.")
    doc.bullet("Commit and push to persist the work: git add . ; git commit -m '...' ; git push")
    doc.spacer(5)
    doc.wrapped_text(
        "Git provides a complete audit trail: every generated file has a commit history showing when it "
        "was created or modified, by whom, and under what AI instruction. For carve-out engagements — "
        "which carry regulatory and legal implications — this auditability is a significant governance "
        "benefit. A reviewer can trace any cell value in a cost plan back to the schedule resource "
        "assignment that produced it, and then to the commit that recorded it."
    )
    doc.spacer(5)
    doc.wrapped_text(
        "The repository-governance-updates skill captures ongoing maintenance responsibilities: whenever "
        "a new project folder is created, a generator script is added, a template changes, or a project "
        "status moves from active to closed, the repository metadata documents must be updated — including "
        "the CLAUDE.md last-reviewed date and the content structure inventory."
    )


def section_14(doc):
    doc.section_bar("14", "Observed Outcomes and Lessons")

    doc.subsection("14.1", "Speed")
    doc.wrapped_text(
        "A complete deliverable set for a new engagement — which previously required several days of "
        "manual work — is now generated in minutes. Generation times measured across multiple engagements "
        "range from under one second for a schedule CSV to approximately two to three seconds for a "
        "full HTML dashboard with embedded base64 logo. The primary time cost is now in reviewing the "
        "output for engagement-specific accuracy, not in document construction."
    )

    doc.subsection("14.2", "Consistency")
    doc.wrapped_text(
        "Because the AI always reads the skill files before generating, and skill files encode canonical "
        "layouts and schemas, outputs are structurally identical across engagements. A stakeholder "
        "familiar with one project charter immediately understands the next — same section structure, "
        "same visual language, same quality gate structure."
    )

    doc.subsection("14.3", "Cross-Check Enforcement")
    doc.wrapped_text(
        "The mandatory cross-checks between schedule, risk register, and cost plan are enforced by "
        "the AI's instruction layer, not by human memory. In practice, this has caught real "
        "inconsistencies: resource names that changed between schedule and cost plan revisions, "
        "phase dates updated in the schedule but not reflected in dashboard countdowns."
    )

    doc.subsection("14.4", "Institutional Memory Accumulation")
    doc.wrapped_text(
        "Each time a significant lesson is learned — the XML TaskMode rule, the CSS logo clipping bug, "
        "the openpyxl .xlsx restriction — it is written into the memory system and enforced from that "
        "point forward. The system becomes more reliable with each engagement, rather than regressing "
        "to previously resolved issues when a new session begins."
    )

    doc.subsection("14.5", "Limitations")
    doc.bullet("<b>Data accuracy is the PM's responsibility:</b> The AI generates structure and formatting "
               "from given facts. If scope facts in the intake are incorrect, the deliverables will be "
               "internally consistent but externally wrong.")
    doc.bullet("<b>Complex Excel formula execution:</b> openpyxl writes formula strings; Excel calculates "
               "them on file open. The risk rating column (=H×I) is written as a formula string, not "
               "pre-calculated — this requires Excel to be available for consumption.")
    doc.bullet("<b>Session restarts:</b> Each Copilot chat session begins fresh. The memory architecture "
               "mitigates this for known rules, but a very long multi-step session benefits from "
               "periodic memory reads to refresh working context.")
    doc.bullet("<b>Scope verification:</b> The AI cannot verify whether the user-supplied facts reflect "
               "the actual engagement reality. The intake gate validates completeness, not accuracy.")


def section_15(doc):
    doc.section_bar("15", "Future Directions")
    doc.wrapped_text(
        "Several extensions to the toolkit are under consideration based on observed usage patterns "
        "across live engagements."
    )
    doc.spacer(5)
    doc.table(
        headers=["Initiative", "Description"],
        rows=[
            ("Dashboard auto-refresh",
             "A script that reads actual progress data and regenerates dashboards automatically without AI intervention, for routine weekly updates"),
            ("Deliverable diff reporting",
             "A tool comparing two versions of a schedule or cost plan that summarises what changed — useful for change-control presentations at SteerCo"),
            ("Portfolio view",
             "An aggregated multi-project dashboard showing status across all active engagements, drawing from the individual project output folders"),
            ("Natural language status entry",
             "A Copilot chat flow where the PM types a brief status update and the AI propagates it to the relevant fields in the report and dashboards"),
            ("Template versioning",
             "Formal versioning of Risk_analysis_template.xlsx and dashboard layout specs, with the AI instructed to use the version tagged at engagement start"),
            ("Quality gate checklist automation",
             "AI-assisted generation of gate-readiness checklists from the schedule's QG milestones and the risk register's open high-priority items"),
        ],
        col_widths=[110, CW - 110],
        row_h=18,
    )


def section_16(doc):
    doc.section_bar("16", "Conclusion")
    doc.wrapped_text(
        "The Carveout AI Toolkit demonstrates that AI-assisted project management, when properly "
        "structured, can go well beyond text generation. By combining a domain-specific instruction "
        "layer, modular skill files encoding methodology, an enforced deliverable sequence, mandatory "
        "cross-check rules, and a memory architecture that accumulates institutional knowledge, the "
        "system produces governance-quality deliverables with high structural consistency and low "
        "manual effort."
    )
    doc.spacer(6)
    doc.wrapped_text(
        "The key architectural insight is that the value of the AI is not in the sophistication of any "
        "single prompt but in the layer of structured knowledge — skills, instructions, memory — that "
        "wraps and guides its inference. Without that layer, the AI produces plausible output. With it, "
        "the AI produces correct, consistent, cross-checked, audit-ready output."
    )
    doc.spacer(6)
    doc.wrapped_text(
        "For IT carve-out practitioners, this translates to a practical shift in where project managers "
        "invest their energy. Document construction is automated. The judgment about which risks matter, "
        "what the real critical path is, and how to communicate with the Steering Committee — that "
        "remains human. The toolkit removes the mechanical work so the practitioner can focus "
        "on the strategic work."
    )
    doc.spacer(8)
    doc.info_box([
        "TOOLKIT SUMMARY",
        "  AI Engine:      GitHub Copilot (Claude Sonnet 4.6) — VS Code Agent mode",
        "  Repository:     Git (Bosch DevCloud) — four-zone content model",
        "  Skills:         8 domain SKILL.md files covering all deliverable types",
        "  Outputs:        7 mandatory deliverables per engagement (schedule, risk, cost, charter, 2 dashboards, report)",
        "  Generator scripts: 14 Python scripts (stdlib + openpyxl + reportlab only)",
        "  Memory:         3-tier (user / session / repository) — accumulates institutional rules",
    ], bg=C_HILIGHT, border=C_ACCENT)


def appendix_a(doc):
    doc.section_bar("A", "Appendix A — Mandatory Engagement Intake Fields")
    doc.wrapped_text(
        "All eleven fields below must be confirmed before any deliverable generation begins. "
        "Missing fields block generation. Estimates and reference-project placeholders are not accepted."
    )
    doc.spacer(4)
    doc.table(
        headers=["Field", "Description", "Example Value"],
        rows=[
            ("Project name",         "Engagement codename",                           "Project Phoenix"),
            ("Seller",               "Entity divesting the business",                  "Robert Bosch GmbH"),
            ("Buyer",                "Entity acquiring the business",                  "TBC / Disclosed Buyer"),
            ("Business carved out",  "Specific business or division in scope",         "Solar Energy Business Unit"),
            ("Carve-out model",      "Stand Alone / Integration / Combination",        "Stand Alone"),
            ("PMO lead",             "Named methodology lead",                         "Named PM (department)"),
            ("Worldwide sites",      "Number of locations in scope",                   "17"),
            ("IT users",             "Target population count",                        "2,600"),
            ("Project start date",   "Phase 0 kickoff date (DD MMM YYYY)",             "01 April 2026"),
            ("GoLive / Day-1 date",  "Operational cutover date (DD MMM YYYY)",         "01 December 2026"),
            ("Completion date",      "QG5 / programme closure date (DD MMM YYYY)",     "30 May 2027"),
        ],
        col_widths=[88, 130, CW - 218],
        row_h=14,
        note="Budget exception: if budget is explicitly unknown, use 'TBC — to be approved at QG1' rather than blocking generation.",
    )


def appendix_b(doc):
    doc.section_bar("B", "Appendix B — Deliverable Generation Sequence")
    doc.wrapped_text(
        "The diagram below shows the dependency chain enforced by the AI instruction layer. "
        "No step can be skipped; each deliverable must exist before the next is generated."
    )
    doc.spacer(8)

    # Draw ASCII-art style flow diagram using low-level canvas drawing
    c = doc.c
    box_w  = 160
    box_h  = 22
    col1_x = LM + (CW / 2 - box_w) / 2
    col2_x = LM + CW / 2 + (CW / 2 - box_w) / 2
    y_start = doc.y - 10

    steps = [
        ("Intake Compliance Gate",      C_AMBER,             True,  None),
        ("1. Schedule (CSV + XML)",      C_HERO,              False, None),
        ("2. Risk Register (.xlsx)",     C_HERO,              False, None),
        ("3. Cost Plan (.csv)",          C_HERO,              False, None),
        ("4. Project Charter (.html)",   colors.HexColor("#005199"), False, None),
        ("5. Executive Dashboard (.html)", colors.HexColor("#005199"), False, None),
        ("6. Management KPI Dashboard (.html)", colors.HexColor("#005199"), False, None),
        ("7. Monthly Status Report (.pdf)", C_ACCENT,         False, None),
    ]

    y = y_start
    for i, (label, bg, is_gate, _) in enumerate(steps):
        c.setFillColor(bg)
        c.roundRect(col1_x, y - box_h, box_w, box_h, 4, fill=1, stroke=0)
        c.setFillColor(C_WHITE)
        c.setFont(FONT_BOLD if is_gate else FONT_BODY, 7.5)
        c.drawCentredString(col1_x + box_w / 2, y - box_h / 2 - 2.5, label)
        if i < len(steps) - 1:
            arr_x = col1_x + box_w / 2
            arr_y = y - box_h - 2
            c.setStrokeColor(C_RULE)
            c.setLineWidth(1)
            c.line(arr_x, arr_y, arr_x, arr_y - 10)
            # draw arrow head as filled path
            from reportlab.graphics.shapes import Drawing
            c.setFillColor(C_RULE)
            p = c.beginPath()
            p.moveTo(arr_x - 4, arr_y - 10)
            p.lineTo(arr_x + 4, arr_y - 10)
            p.lineTo(arr_x,     arr_y - 16)
            p.close()
            c.drawPath(p, fill=1, stroke=0)
        y -= (box_h + 14)

    doc.y = y - 10


def appendix_c(doc):
    doc.section_bar("C", "Appendix C — Generator Scripts Inventory")
    doc.wrapped_text(
        "All generator scripts reside at the repository root. Per-project scripts are named "
        "generate_{project}_{type}.py. The shared canonical XML converter (generate_msp_xml.py) "
        "is invoked by all schedule scripts as a subprocess."
    )
    doc.spacer(4)
    doc.table(
        headers=["Script", "Scope", "Produces"],
        rows=[
            ("generate_msp_xml.py",               "All projects (shared)", "MSPDI XML from any schedule CSV"),
            ("generate_{project}_schedule.py",     "Per engagement",        "Schedule CSV + XML (calls generate_msp_xml.py)"),
            ("generate_{project}_risk_register.py","Per engagement",        "Risk register .xlsx (based on Risk_analysis_template.xlsx)"),
            ("generate_{project}_charter.py",      "Per engagement",        "Project charter self-contained .html"),
            ("generate_{project}_monthly_report.py","Per engagement",       "Monthly status report .pdf (auto-dated filename)"),
            ("generate_project_charter.py",        "Multi-project capable", "Project charter .html (parametric version)"),
            ("generate_risk_register.py",          "Reference / Trinity",   "Risk register .xlsx (reference implementation)"),
            ("generate_carveout_schedule_direct_xml.py","Legacy",           "Direct XML generation (legacy workflow)"),
        ],
        col_widths=[138, 80, CW - 218],
        row_h=14,
    )
    doc.spacer(8)
    doc.wrapped_text(
        "All fourteen scripts follow the standard timing wrapper pattern (see Section 11.1) and record "
        "start time, finish time, and elapsed seconds to stdout on every run."
    )

    # Final note
    doc.spacer(12)
    doc.rule()
    doc.spacer(6)
    doc.text("End of Paper", font=FONT_OBLI, size=8, color=C_MUTED, center=True)
    doc.spacer(4)
    doc.text(f"AI-Assisted Project Management for IT Carve-Outs  ·  Bosch MIL Practice  ·  {PUB_DATE}",
             font=FONT_BODY, size=7, color=C_MUTED, center=True)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    doc = Doc(OUT_PATH)

    # Cover page
    draw_cover(doc)

    # Table of Contents
    doc.text("TABLE OF CONTENTS", font=FONT_BOLD, size=10.5, color=C_HERO)
    doc.y -= 4
    draw_toc(doc)

    # Body sections
    section_1(doc)
    section_2(doc)
    section_3(doc)
    section_4(doc)
    section_5(doc)
    section_6(doc)
    section_7(doc)
    section_8(doc)
    section_9(doc)
    section_10(doc)
    section_11(doc)
    section_12(doc)
    section_13(doc)
    section_14(doc)
    section_15(doc)
    section_16(doc)

    # Appendices
    appendix_a(doc)
    appendix_b(doc)
    appendix_c(doc)

    doc.save()
    print(f"PDF written: {OUT_PATH}")
    sz = os.path.getsize(OUT_PATH)
    print(f"File size : {sz / 1024:.1f} KB  ({doc.page} pages)")


if __name__ == "__main__":
    from datetime import datetime as _dt
    _t0 = _dt.now()
    print(f"Started : {_t0.strftime('%Y-%m-%d %H:%M:%S')}")
    try:
        main()
    finally:
        _t1 = _dt.now()
        print(f"Finished: {_t1.strftime('%Y-%m-%d %H:%M:%S')}  ({(_t1-_t0).total_seconds():.1f}s elapsed)")
