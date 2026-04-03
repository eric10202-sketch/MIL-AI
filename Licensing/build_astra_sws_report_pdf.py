"""
Generate PDF Management Report for Astra Software License Shopping List (SWS)
Output: Astra_SWS_Management_Report.pdf
"""
import os, io, base64, collections
import openpyxl

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, HRFlowable, KeepTogether,
                                 PageBreak, Image, Flowable)
from reportlab.graphics.shapes import Drawing, Rect, String, Line, Circle
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.barcharts import HorizontalBarChart
from reportlab.graphics import renderPDF
from reportlab.lib.utils import ImageReader

# ── Bosch colours ──────────────────────────────────────────────────────────
C_RED   = colors.HexColor('#dc0000')
C_NAVY  = colors.HexColor('#003c64')
C_GREEN = colors.HexColor('#00783c')
C_ORANGE= colors.HexColor('#e65100')
C_GRAY  = colors.HexColor('#6c757d')
C_LGRAY = colors.HexColor('#f4f5f7')
C_DARK  = colors.HexColor('#1a1a1a')
C_WHITE = colors.white

BASE = r'c:\Users\kho1sgp\OneDrive - Bosch Group\My Work Documents\AI topics\BDMIL use case\Software_license_project'

# ── Load logo ──────────────────────────────────────────────────────────────
with open(os.path.join(BASE, 'Bosch_png_b64.txt')) as f:
    logo_bytes = base64.b64decode(f.read().strip())
logo_io = io.BytesIO(logo_bytes)

# ── Load data ──────────────────────────────────────────────────────────────
def clean(v):
    s = str(v or '').strip()
    return s if s not in ('None', '') else ''

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

# KPIs
status_cnt  = collections.Counter(r['status'].lower().strip() for r in records if r['status'])
done        = status_cnt.get('done', 0)
out_scope   = status_cnt.get('out of scope', 0)
out_dup     = status_cnt.get('out-duplicate', 0)
replaced    = status_cnt.get('replaced', 0)
in_clarif   = status_cnt.get('in clarification', 0)

purch_cnt   = collections.Counter(r['purch_status'].strip() for r in records if r['purch_status'])
ready_proc  = purch_cnt.get('ready for procurement', 0)
freeware    = purch_cnt.get('freeware', 0)
oss         = purch_cnt.get('Open Source Software (OSS)', 0)
no_proc     = purch_cnt.get('No procurment needed', 0)
bu_proc     = purch_cnt.get('BU procurement', 0)
sws_proc    = purch_cnt.get('SWS procurement', 0)
tsa         = purch_cnt.get('TSA', 0)

sws_cnt     = collections.Counter(r['sws'] for r in records if r['sws'])
bdbu_raw    = collections.Counter(r['bdbu'].strip().upper() for r in records if r['bdbu'])
bd_cnt      = bdbu_raw.get('BD', 0)
bu_cnt      = bdbu_raw.get('BU', 0)

pt_cnt      = collections.Counter((r['pay_type'] or 'blank').lower() for r in records)
sub_cnt     = sum(v for k, v in pt_cnt.items() if 'sub' in k)
perp_cnt    = sum(v for k, v in pt_cnt.items() if 'perp' in k)
free_cnt    = sum(v for k, v in pt_cnt.items() if 'free' in k or 'oss' in k)
tbd_cnt     = sum(v for k, v in pt_cnt.items() if 'tbd' in k or 'blank' in k or 'unk' in k)
seq_cnt     = sum(v for k, v in pt_cnt.items() if 'see' in k or 'paid' in k)

unique_vendors  = len(set(r['vendor'].lower() for r in records if r['vendor']))
with_req_date   = sum(1 for r in records if r['req_date'])
with_contact    = sum(1 for r in records if r['sw_contact'])
with_notes      = sum(1 for r in records if r['notes'])
with_paytype    = sum(1 for r in records if r['pay_type'])

vendor_cnt  = collections.Counter(r['vendor'] for r in records if r['vendor'])
top10       = vendor_cnt.most_common(10)
top10_rev   = list(reversed(top10))

# SWS colors
sws_colors_hex = ['#003c64','#00783c','#e65100','#dc0000','#6c757d','#795548']
sws_rl_colors  = [colors.HexColor(h) for h in sws_colors_hex]

# ── Helpers ────────────────────────────────────────────────────────────────
def pct(val, tot):
    return round(100 * val / tot) if tot else 0

def make_pie(parts_vals, parts_colors, size=130):
    """Return a Drawing with a pie chart."""
    d = Drawing(size, size)
    pie = Pie()
    pie.x = size // 2 - size // 2 + 10
    pie.y = 10
    pie.width  = size - 20
    pie.height = size - 20
    pie.data   = [v for v, _ in parts_vals]
    pie.slices.strokeWidth = 0.5
    pie.slices.strokeColor = colors.white
    pie.sideLabels = False
    pie.simpleLabels = False
    for i, (_, col) in enumerate(parts_vals):
        pie.slices[i].fillColor = col
    d.add(pie)
    return d

def make_hbar(items, bar_color=C_NAVY, width=200, height=None):
    """items = [(label, value)], returns a Drawing."""
    n = len(items)
    row_h = 16
    pad = 10
    if height is None:
        height = n * row_h + pad * 2
    d = Drawing(width, height)
    max_v = max(v for _, v in items) if items else 1
    bar_area = width - 120
    for i, (label, val) in enumerate(items):
        y = height - pad - (i + 1) * row_h + 4
        bar_w = round(bar_area * val / max_v)
        # label
        d.add(String(0, y + 2, label[:22], fontName='Helvetica', fontSize=7,
                     fillColor=C_DARK))
        # bar track
        d.add(Rect(110, y, bar_area, 10, fillColor=colors.HexColor('#f0f0f0'),
                   strokeColor=None))
        # bar fill
        if bar_w > 0:
            d.add(Rect(110, y, bar_w, 10, fillColor=bar_color, strokeColor=None))
        # value label
        d.add(String(112 + bar_w + 3, y + 2, str(val), fontName='Helvetica-Bold',
                     fontSize=7, fillColor=C_DARK))
    return d

def make_progress_bar(label, val, tot, bar_color, width=380):
    """Returns a Table row with a progress bar."""
    pct_v = pct(val, tot)
    bar_w = round(width * 0.5 * pct_v / 100)
    d = Drawing(width * 0.5, 12)
    d.add(Rect(0, 0, width * 0.5, 12, fillColor=colors.HexColor('#f0f0f0'),
               strokeColor=None))
    if bar_w > 0:
        d.add(Rect(0, 0, bar_w, 12, fillColor=bar_color, strokeColor=None))
    return [label, d, f'{val} ({pct_v}%)']

# ── Styles ─────────────────────────────────────────────────────────────────
styles = getSampleStyleSheet()

def S(name, **kw):
    return ParagraphStyle(name, **kw)

s_title   = S('title',   fontName='Helvetica-Bold', fontSize=18, textColor=C_WHITE,
              leading=22, alignment=TA_LEFT)
s_sub     = S('sub',     fontName='Helvetica',      fontSize=10, textColor=colors.HexColor('#bbbbbb'),
              leading=14, alignment=TA_LEFT)
s_section = S('section', fontName='Helvetica-Bold', fontSize=11, textColor=C_NAVY,
              leading=16, spaceBefore=18, spaceAfter=6,
              borderPad=4)
s_body    = S('body',    fontName='Helvetica',      fontSize=9,  textColor=C_DARK,
              leading=13, spaceAfter=4)
s_bold    = S('bold',    fontName='Helvetica-Bold', fontSize=9,  textColor=C_DARK,
              leading=13)
s_small   = S('small',   fontName='Helvetica',      fontSize=8,  textColor=C_GRAY,
              leading=11)
s_kpi_val = S('kpiv',    fontName='Helvetica-Bold', fontSize=22, textColor=C_DARK,
              leading=26, alignment=TA_CENTER)
s_kpi_lbl = S('kpil',    fontName='Helvetica',      fontSize=8,  textColor=C_GRAY,
              leading=11, alignment=TA_CENTER)
s_th      = S('th',      fontName='Helvetica-Bold', fontSize=8,  textColor=C_WHITE,
              leading=11, alignment=TA_CENTER)
s_td      = S('td',      fontName='Helvetica',      fontSize=8,  textColor=C_DARK,
              leading=11, alignment=TA_LEFT)
s_td_c    = S('tdc',     fontName='Helvetica',      fontSize=8,  textColor=C_DARK,
              leading=11, alignment=TA_CENTER)
s_bullet  = S('bullet',  fontName='Helvetica',      fontSize=9,  textColor=C_DARK,
              leading=13, spaceAfter=3, leftIndent=12, firstLineIndent=-12)
s_narr    = S('narr',    fontName='Helvetica',      fontSize=9.5,textColor=C_DARK,
              leading=14, spaceAfter=6,
              backColor=colors.HexColor('#e3f0fb'),
              borderPad=8)

# ── Page template helpers ──────────────────────────────────────────────────
PAGE_W, PAGE_H = A4

def on_first_page(canvas, doc):
    pass

def on_later_pages(canvas, doc):
    canvas.saveState()
    # footer
    canvas.setFont('Helvetica', 7)
    canvas.setFillColor(C_GRAY)
    canvas.drawString(2*cm, 1.2*cm,
        'Astra Software License Management Report  ·  BDMIL Use Case  ·  April 2, 2026')
    canvas.drawRightString(PAGE_W - 2*cm, 1.2*cm, f'Page {doc.page}')
    canvas.setStrokeColor(colors.HexColor('#e0e0e0'))
    canvas.line(2*cm, 1.5*cm, PAGE_W - 2*cm, 1.5*cm)
    canvas.restoreState()

# ── Cover header block ─────────────────────────────────────────────────────
class CoverHeader(Flowable):
    def __init__(self, logo_io, width):
        Flowable.__init__(self)
        self.logo_io = logo_io
        self.width   = width
        self.height  = 90

    def draw(self):
        c = self.canv
        w, h = self.width, self.height
        # dark background
        c.setFillColor(C_DARK)
        c.rect(0, 0, w, h, fill=1, stroke=0)
        # colour strip at bottom of header
        strip_h = 8
        for i, col in enumerate([C_RED, C_NAVY, C_GREEN]):
            c.setFillColor(col)
            seg_w = w / 3
            c.rect(i * seg_w, 0, seg_w, strip_h, fill=1, stroke=0)
        # logo
        logo_reader = ImageReader(io.BytesIO(self.logo_io.getvalue()))
        c.drawImage(logo_reader, 18, strip_h + 14, width=90, height=40,
                    preserveAspectRatio=True, mask='auto')
        # title text
        c.setFont('Helvetica-Bold', 14)
        c.setFillColor(C_WHITE)
        c.drawRightString(w - 18, strip_h + 38,
                          'Astra Software License Management Report')
        c.setFont('Helvetica', 9)
        c.setFillColor(colors.HexColor('#bbbbbb'))
        c.drawRightString(w - 18, strip_h + 22,
                          'BDMIL Use Case  ·  Shopping List with SWS  ·  April 2, 2026')

class SectionHeader(Flowable):
    def __init__(self, text, color=C_NAVY, width=None):
        Flowable.__init__(self)
        self.text   = text
        self.color  = color
        self.width  = width or (PAGE_W - 4*cm)
        self.height = 22

    def draw(self):
        c = self.canv
        c.setFillColor(self.color)
        c.rect(0, 0, 4, self.height, fill=1, stroke=0)
        c.setFillColor(colors.HexColor('#f4f5f7'))
        c.rect(4, 0, self.width - 4, self.height, fill=1, stroke=0)
        c.setFont('Helvetica-Bold', 10)
        c.setFillColor(self.color)
        c.drawString(12, 6, self.text.upper())

def kpi_table(kpis):
    """kpis = [(value, label, sub, color), ...]"""
    cells = []
    for val, lbl, sub, col in kpis:
        inner = Table([
            [Paragraph(str(val), ParagraphStyle('kv', fontName='Helvetica-Bold',
                fontSize=20, textColor=col, leading=24, alignment=TA_CENTER))],
            [Paragraph(lbl, ParagraphStyle('kl', fontName='Helvetica', fontSize=7.5,
                textColor=C_GRAY, leading=10, alignment=TA_CENTER))],
            [Paragraph(sub, ParagraphStyle('ks', fontName='Helvetica', fontSize=7,
                textColor=colors.HexColor('#999999'), leading=9, alignment=TA_CENTER))],
        ], colWidths=[3.5*cm])
        inner.setStyle(TableStyle([
            ('BACKGROUND', (0,0),(-1,-1), C_WHITE),
            ('TOPPADDING',  (0,0),(-1,-1), 6),
            ('BOTTOMPADDING',(0,0),(-1,-1), 6),
            ('LEFTPADDING', (0,0),(-1,-1), 4),
            ('RIGHTPADDING',(0,0),(-1,-1), 4),
            ('LINEABOVE',   (0,0),(-1,0),  2, col),
            ('ROUNDEDCORNERS', [4]),
        ]))
        cells.append(inner)
    n = len(cells)
    row = Table([cells], colWidths=[3.7*cm]*n)
    row.setStyle(TableStyle([
        ('LEFTPADDING',  (0,0),(-1,-1), 4),
        ('RIGHTPADDING', (0,0),(-1,-1), 4),
        ('TOPPADDING',   (0,0),(-1,-1), 0),
        ('BOTTOMPADDING',(0,0),(-1,-1), 0),
    ]))
    return row

# ── Build content ──────────────────────────────────────────────────────────
story = []

# ─ Cover header
story.append(CoverHeader(logo_io, PAGE_W - 4*cm))
story.append(Spacer(1, 10))

# ─ Executive summary narrative
story.append(SectionHeader('Executive Summary', C_RED))
story.append(Spacer(1, 6))

narr_text = (
    f'<b>Scope:</b> The Astra software license portfolio covers <b>{total} products</b> from '
    f'<b>{unique_vendors} unique vendors</b>, tracked across 6 SWS systems (SWS1 through SWS8). '
    f'Data is sourced from the Astra Shopping List with SWS as of April 2, 2026.<br/><br/>'
    f'<b>Progress:</b> <b>{done} items ({pct(done,total)}%)</b> have confirmed procurement status "Done". '
    f'<b>{ready_proc} items</b> are classified as ready for procurement and pending ordering. '
    f'<b>{in_clarif} items</b> remain in clarification and require immediate resolution. '
    f'{out_scope + out_dup + replaced} items are excluded from active procurement '
    f'({out_scope} out-of-scope, {out_dup} duplicates, {replaced} replaced).<br/><br/>'
    f'<b>License Model:</b> Of actively classified paid items, <b>{sub_cnt} are subscription-based</b> '
    f'and <b>{perp_cnt} are perpetual</b>. <b>{free_cnt} items</b> are classified as free or '
    f'Open Source Software. <b>{tbd_cnt} items ({pct(tbd_cnt,total)}%)</b> have no payment type '
    f'defined — this is the key data gap requiring resolution.<br/><br/>'
    f'<b>BD vs BU:</b> {bd_cnt} products ({pct(bd_cnt,total)}%) are attributed to BD; '
    f'{bu_cnt} products ({pct(bu_cnt,total)}%) to BU.'
)
story.append(Paragraph(narr_text, ParagraphStyle('narr2', fontName='Helvetica', fontSize=9,
    textColor=C_DARK, leading=13.5, spaceAfter=8,
    backColor=colors.HexColor('#eaf3fb'), borderPad=10,
    leftIndent=0, rightIndent=0)))
story.append(Spacer(1, 8))

# ─ KPI row 1
story.append(kpi_table([
    (total,         'Total Products',              'Astra Shopping List',       C_NAVY),
    (done,          f'Procurement Done',           f'{pct(done,total)}% of total', C_GREEN),
    (ready_proc,    'Ready for Procurement',       f'{pct(ready_proc,total)}% of total', C_NAVY),
    (in_clarif,     'In Clarification',            'Require immediate action',  C_ORANGE),
]))
story.append(Spacer(1, 8))
# KPI row 2
story.append(kpi_table([
    (unique_vendors,    'Unique Vendors',           'Across all SWS',           C_GRAY),
    (freeware + oss,    'Freeware / OSS',           f'{freeware} freeware · {oss} OSS', C_GREEN),
    (with_req_date,     'License Date Defined',     f'{pct(with_req_date,total)}% populated', C_NAVY),
    (out_scope+out_dup+replaced, 'Excluded Items',  f'{out_scope} OOS · {out_dup} dup · {replaced} repl', C_GRAY),
]))
story.append(Spacer(1, 16))

# ─ Status breakdown
story.append(SectionHeader('Item Status & Procurement Status', C_NAVY))
story.append(Spacer(1, 8))

# Two-column: status table left, purchasing status table right
status_rows = [
    [Paragraph('Status', s_th), Paragraph('Items', s_th), Paragraph('%', s_th)],
    [Paragraph('Done', s_td), Paragraph(str(done), s_td_c), Paragraph(f'{pct(done,total)}%', s_td_c)],
    [Paragraph('In Clarification', s_td), Paragraph(str(in_clarif), s_td_c), Paragraph(f'{pct(in_clarif,total)}%', s_td_c)],
    [Paragraph('Out of Scope', s_td), Paragraph(str(out_scope), s_td_c), Paragraph(f'{pct(out_scope,total)}%', s_td_c)],
    [Paragraph('Out-Duplicate', s_td), Paragraph(str(out_dup), s_td_c), Paragraph(f'{pct(out_dup,total)}%', s_td_c)],
    [Paragraph('Replaced', s_td), Paragraph(str(replaced), s_td_c), Paragraph(f'{pct(replaced,total)}%', s_td_c)],
    [Paragraph('<b>TOTAL</b>', s_bold), Paragraph(f'<b>{total}</b>', ParagraphStyle('bold_c', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER)), Paragraph('<b>100%</b>', ParagraphStyle('bold_c2', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER))],
]
t_status = Table(status_rows, colWidths=[4.5*cm, 2*cm, 1.5*cm])
t_status.setStyle(TableStyle([
    ('BACKGROUND',  (0,0), (-1,0),  C_NAVY),
    ('BACKGROUND',  (0,1), (-1,1),  colors.HexColor('#e6f4ea')),
    ('BACKGROUND',  (0,2), (-1,2),  colors.HexColor('#fff3e0')),
    ('BACKGROUND',  (0,6), (-1,6),  colors.HexColor('#f4f5f7')),
    ('GRID',        (0,0), (-1,-1), 0.5, colors.HexColor('#e0e0e0')),
    ('TOPPADDING',  (0,0), (-1,-1), 5),
    ('BOTTOMPADDING',(0,0),(-1,-1), 5),
    ('LEFTPADDING', (0,0), (-1,-1), 6),
    ('RIGHTPADDING',(0,0), (-1,-1), 6),
    ('FONTNAME',    (0,0), (-1,0),  'Helvetica-Bold'),
    ('FONTNAME',    (0,6), (-1,6),  'Helvetica-Bold'),
]))

purch_rows = [
    [Paragraph('Purchasing Status', s_th), Paragraph('Items', s_th), Paragraph('%', s_th)],
]
for ps_label, ps_val in [
    ('Ready for Procurement', ready_proc),
    ('Freeware',              freeware),
    ('Open Source (OSS)',     oss),
    ('No Procurement Needed', no_proc),
    ('BU Procurement',        bu_proc),
    ('SWS Procurement',       sws_proc),
    ('TSA',                   tsa),
]:
    purch_rows.append([
        Paragraph(ps_label, s_td),
        Paragraph(str(ps_val), s_td_c),
        Paragraph(f'{pct(ps_val,total)}%', s_td_c),
    ])
purch_rows.append([
    Paragraph('<b>TOTAL with status</b>', s_bold),
    Paragraph(f'<b>{sum(p[1] for p in [(ready_proc,1),(freeware,1),(oss,1),(no_proc,1),(bu_proc,1),(sws_proc,1),(tsa,1)])}</b>', ParagraphStyle('b_c', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER)),
    Paragraph('', s_td_c),
])
t_purch = Table(purch_rows, colWidths=[4.5*cm, 2*cm, 1.5*cm])
t_purch.setStyle(TableStyle([
    ('BACKGROUND',  (0,0), (-1,0),  C_NAVY),
    ('BACKGROUND',  (0,1), (-1,1),  colors.HexColor('#e3f0fb')),
    ('BACKGROUND',  (0,-1),(-1,-1), colors.HexColor('#f4f5f7')),
    ('GRID',        (0,0), (-1,-1), 0.5, colors.HexColor('#e0e0e0')),
    ('TOPPADDING',  (0,0), (-1,-1), 5),
    ('BOTTOMPADDING',(0,0),(-1,-1), 5),
    ('LEFTPADDING', (0,0), (-1,-1), 6),
    ('RIGHTPADDING',(0,0), (-1,-1), 6),
    ('FONTNAME',    (0,-1),(-1,-1), 'Helvetica-Bold'),
]))

# Side by side
story.append(Table([[t_status, Spacer(0.5*cm,1), t_purch]],
    colWidths=[8*cm, 0.5*cm, 8*cm]))
story.append(Spacer(1, 16))

# ─ Payment model
story.append(SectionHeader('Payment / License Model', C_NAVY))
story.append(Spacer(1, 8))

pay_items = [
    ('Subscription',         sub_cnt,  C_NAVY),
    ('Perpetual',            perp_cnt, C_GREEN),
    ('Free / OSS',           free_cnt, C_GRAY),
    ('TBD / Not Defined',    tbd_cnt,  C_RED),
    ('Other (see quote etc)',seq_cnt,  C_ORANGE),
]
pay_rows = [
    [Paragraph('License Model', s_th), Paragraph('Items', s_th),
     Paragraph('%', s_th), Paragraph('Visual', s_th)],
]
for label, val, col in pay_items:
    bar_w = round(100 * val / total) if total else 0
    bar_d = Drawing(80, 10)
    bar_d.add(Rect(0, 0, 80, 10, fillColor=colors.HexColor('#f0f0f0'), strokeColor=None))
    if bar_w > 0:
        bar_d.add(Rect(0, 0, bar_w * 0.8, 10, fillColor=col, strokeColor=None))
    pay_rows.append([
        Paragraph(label, s_td),
        Paragraph(str(val), s_td_c),
        Paragraph(f'{pct(val,total)}%', s_td_c),
        bar_d,
    ])
t_pay = Table(pay_rows, colWidths=[5*cm, 2*cm, 1.5*cm, 8*cm])
t_pay.setStyle(TableStyle([
    ('BACKGROUND',  (0,0), (-1,0),  C_NAVY),
    ('GRID',        (0,0), (-1,-1), 0.5, colors.HexColor('#e0e0e0')),
    ('TOPPADDING',  (0,0), (-1,-1), 5),
    ('BOTTOMPADDING',(0,0),(-1,-1), 5),
    ('LEFTPADDING', (0,0), (-1,-1), 6),
    ('RIGHTPADDING',(0,0), (-1,-1), 6),
    ('ROWBACKGROUNDS', (0,1),(-1,-1), [C_WHITE, colors.HexColor('#f9f9f9')]),
]))
story.append(t_pay)
story.append(Spacer(1, 16))

# ─ SWS Breakdown
story.append(KeepTogether([
    SectionHeader('SWS System Distribution', C_NAVY),
    Spacer(1, 8),
]))

sws_detail_rows = [
    [Paragraph(h, s_th) for h in ['SWS', 'Total', 'Done', 'Clarif.', 'OOS/Dup/Repl', 'Ready Proc.', 'Completion']],
]
for sws_key, sws_val in sws_cnt.most_common():
    s_done  = sum(1 for r in records if r['sws']==sws_key and r['status'].lower().strip()=='done')
    s_clar  = sum(1 for r in records if r['sws']==sws_key and 'clarif' in r['status'].lower())
    s_excl  = sum(1 for r in records if r['sws']==sws_key
                  and any(x in r['status'].lower() for x in ['scope','duplic','replac']))
    s_ready = sum(1 for r in records if r['sws']==sws_key and r['purch_status']=='ready for procurement')
    comp    = pct(s_done, sws_val)
    sws_detail_rows.append([
        Paragraph(f'<b>{sws_key}</b>', s_td),
        Paragraph(str(sws_val), s_td_c),
        Paragraph(str(s_done),  s_td_c),
        Paragraph(str(s_clar) if s_clar else '–',  s_td_c),
        Paragraph(str(s_excl),  s_td_c),
        Paragraph(str(s_ready), s_td_c),
        Paragraph(f'{comp}%',   ParagraphStyle('comp_c', fontName='Helvetica-Bold', fontSize=8,
            textColor=C_GREEN if comp>=50 else C_ORANGE, alignment=TA_CENTER)),
    ])
sws_detail_rows.append([
    Paragraph('<b>TOTAL</b>', s_bold),
    Paragraph(f'<b>{total}</b>', ParagraphStyle('tc', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER)),
    Paragraph(f'<b>{done}</b>',  ParagraphStyle('tc', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, textColor=C_GREEN)),
    Paragraph(f'<b>{in_clarif}</b>', ParagraphStyle('tc', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, textColor=C_ORANGE)),
    Paragraph(f'<b>{out_scope+out_dup+replaced}</b>', ParagraphStyle('tc', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER)),
    Paragraph(f'<b>{ready_proc}</b>', ParagraphStyle('tc', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER)),
    Paragraph(f'<b>{pct(done,total)}%</b>', ParagraphStyle('tc', fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER, textColor=C_GREEN)),
])
t_sws = Table(sws_detail_rows, colWidths=[2*cm, 1.5*cm, 1.5*cm, 1.5*cm, 2.5*cm, 2.5*cm, 2.5*cm])
t_sws.setStyle(TableStyle([
    ('BACKGROUND',  (0,0), (-1,0),  C_NAVY),
    ('BACKGROUND',  (0,-1),(-1,-1), colors.HexColor('#f4f5f7')),
    ('GRID',        (0,0), (-1,-1), 0.5, colors.HexColor('#e0e0e0')),
    ('TOPPADDING',  (0,0), (-1,-1), 5),
    ('BOTTOMPADDING',(0,0),(-1,-1), 5),
    ('LEFTPADDING', (0,0), (-1,-1), 6),
    ('RIGHTPADDING',(0,0), (-1,-1), 6),
    ('ROWBACKGROUNDS', (0,1),(-1,-2), [C_WHITE, colors.HexColor('#f9f9f9')]),
    ('FONTNAME',    (0,-1),(-1,-1), 'Helvetica-Bold'),
]))
story.append(t_sws)
story.append(Spacer(1, 6))

bd_bu_text = (f'<b>BD / BU Attribution:</b> {bd_cnt} products ({pct(bd_cnt,total)}%) attributed to BD  |  '
              f'{bu_cnt} products ({pct(bu_cnt,total)}%) attributed to BU.')
story.append(Paragraph(bd_bu_text, s_body))
story.append(Spacer(1, 14))

# ─ Top 10 Vendors
story.append(KeepTogether([
    SectionHeader('Top 10 Vendors by Product Count', C_NAVY),
    Spacer(1, 8),
]))

vendor_hbar = make_hbar(top10, bar_color=C_NAVY, width=420, height=190)
story.append(vendor_hbar)
story.append(Spacer(1, 8))

# Full vendor table (top 10)
v_rows = [[Paragraph(h, s_th) for h in ['#', 'Vendor', 'Products', '% of Total']]]
for rank, (vendor, cnt_v) in enumerate(top10, 1):
    v_rows.append([
        Paragraph(str(rank), s_td_c),
        Paragraph(vendor,    s_td),
        Paragraph(str(cnt_v),s_td_c),
        Paragraph(f'{pct(cnt_v, total)}%', s_td_c),
    ])
t_vendors = Table(v_rows, colWidths=[1*cm, 8*cm, 2.5*cm, 2.5*cm])
t_vendors.setStyle(TableStyle([
    ('BACKGROUND',  (0,0), (-1,0),  C_NAVY),
    ('GRID',        (0,0), (-1,-1), 0.5, colors.HexColor('#e0e0e0')),
    ('TOPPADDING',  (0,0), (-1,-1), 5),
    ('BOTTOMPADDING',(0,0),(-1,-1), 5),
    ('LEFTPADDING', (0,0), (-1,-1), 6),
    ('RIGHTPADDING',(0,0), (-1,-1), 6),
    ('ROWBACKGROUNDS', (0,1),(-1,-1), [C_WHITE, colors.HexColor('#f9f9f9')]),
]))
story.append(t_vendors)
story.append(Spacer(1, 14))

# ─ Data quality
story.append(KeepTogether([
    SectionHeader('Data Completeness', C_NAVY),
    Spacer(1, 8),
]))

dq_items = [
    ('Status defined',          sum(1 for r in records if r['status']),     C_GREEN),
    ('Purchasing status set',   sum(1 for r in records if r['purch_status']),C_GREEN),
    ('License date defined',    with_req_date,                               C_NAVY),
    ('SW Key Contact assigned', with_contact,                                C_NAVY),
    ('Notes / comments filled', with_notes,                                  C_GRAY),
    ('Payment type set',        with_paytype,                                C_GRAY),
    ('TBD / blank payment type',tbd_cnt,                                     C_RED),
]
dq_rows = [[Paragraph(h, s_th) for h in ['Field', 'Items Populated', '% Populated', 'Progress']]]
for label, val, col in dq_items:
    pct_v = pct(val, total)
    bw = round(pct_v * 0.8)
    bd = Drawing(80, 10)
    bd.add(Rect(0, 0, 80, 10, fillColor=colors.HexColor('#f0f0f0'), strokeColor=None))
    if bw > 0:
        bd.add(Rect(0, 0, bw, 10, fillColor=col, strokeColor=None))
    dq_rows.append([
        Paragraph(label, s_td),
        Paragraph(str(val), s_td_c),
        Paragraph(f'{pct_v}%', s_td_c),
        bd,
    ])
t_dq = Table(dq_rows, colWidths=[5*cm, 2.5*cm, 2.5*cm, 7.5*cm])
t_dq.setStyle(TableStyle([
    ('BACKGROUND',  (0,0), (-1,0),  C_NAVY),
    ('BACKGROUND',  (0,-1),(-1,-1), colors.HexColor('#fff0f0')),
    ('GRID',        (0,0), (-1,-1), 0.5, colors.HexColor('#e0e0e0')),
    ('TOPPADDING',  (0,0), (-1,-1), 5),
    ('BOTTOMPADDING',(0,0),(-1,-1), 5),
    ('LEFTPADDING', (0,0), (-1,-1), 6),
    ('RIGHTPADDING',(0,0), (-1,-1), 6),
    ('ROWBACKGROUNDS', (0,1),(-1,-2), [C_WHITE, colors.HexColor('#f9f9f9')]),
]))
story.append(t_dq)
story.append(Spacer(1, 16))

# ═══════════════════════════════════════════════════════════════
# PAGE 3 – NEXT STEPS
# ═══════════════════════════════════════════════════════════════
story.append(PageBreak())
story.append(CoverHeader(logo_io, PAGE_W - 4*cm))
story.append(Spacer(1, 10))

story.append(SectionHeader('Recommended Next Steps & Actions', C_RED))
story.append(Spacer(1, 8))

intro_text = (
    'Based on the analysis of the Astra Software License Shopping List (563 products, '
    f'{unique_vendors} vendors, April 2026), the following prioritised action plan is recommended '
    'to complete procurement, resolve data gaps, and ensure license compliance.'
)
story.append(Paragraph(intro_text, s_body))
story.append(Spacer(1, 10))

P_HEX = {'HIGH': '#dc0000', 'MEDIUM': '#e65100', 'LOW': '#6c757d'}

def action_section(title, color, actions):
    """Render a coloured action block."""
    rows = [[Paragraph(f'<b>{title}</b>',
                       ParagraphStyle('ah', fontName='Helvetica-Bold', fontSize=10,
                                      textColor=C_WHITE, leading=14))]]
    t_hdr = Table(rows, colWidths=[PAGE_W - 4*cm])
    t_hdr.setStyle(TableStyle([
        ('BACKGROUND',  (0,0),(-1,-1), color),
        ('TOPPADDING',  (0,0),(-1,-1), 7),
        ('BOTTOMPADDING',(0,0),(-1,-1), 7),
        ('LEFTPADDING', (0,0),(-1,-1), 10),
    ]))
    block = [t_hdr]
    for priority, action, detail in actions:
        ph = P_HEX[priority]
        p_text = (f'<font color="{ph}" size="7">'
                  f'<b>[{priority}]</b></font>  <b>{action}</b><br/>'
                  f'<font size="8" color="#444444">{detail}</font>')
        row_t = Table([[Paragraph(p_text, ParagraphStyle('act', fontName='Helvetica',
                                 fontSize=9, leading=13, leftIndent=0))]],
                      colWidths=[PAGE_W - 4*cm])
        row_t.setStyle(TableStyle([
            ('BACKGROUND',  (0,0),(-1,-1), colors.HexColor('#fafafa')),
            ('TOPPADDING',  (0,0),(-1,-1), 6),
            ('BOTTOMPADDING',(0,0),(-1,-1), 6),
            ('LEFTPADDING', (0,0),(-1,-1), 12),
            ('RIGHTPADDING',(0,0),(-1,-1), 10),
            ('LINEBELOW',   (0,0),(-1,-1), 0.5, colors.HexColor('#eeeeee')),
        ]))
        block.append(row_t)
    return block

# IMMEDIATE actions
immediate = [
    ('HIGH',
     f'Resolve {in_clarif} items currently in clarification',
     f'Each of the {in_clarif} "in clarification" items must have a responsible owner assigned '
     f'and a resolution target date. Escalate unresolved items to the SWS workstream leads.'),
    ('HIGH',
     f'Classify {tbd_cnt} items with blank / TBD payment type',
     f'{tbd_cnt} items ({pct(tbd_cnt,total)}% of total) have no payment type defined. '
     f'Procurement cannot be completed for these items until the license model '
     f'(subscription, perpetual, freeware, OSS) is confirmed with the vendor.'),
    ('HIGH',
     f'Advance {ready_proc} "Ready for Procurement" items to active ordering',
     f'{ready_proc} items are fully classified and approved but have not yet entered the '
     f'ordering process. Trigger procurement workflow for all SWS workstreams without delay.'),
    ('HIGH',
     f'Confirm {bu_proc} BU-procured items with local business units',
     f'{bu_proc} items are delegated to BU procurement. Validate status and expected completion '
     f'dates with each BU to avoid gaps at cutover.'),
]

story.extend(action_section('🔴  Immediate Actions (Complete within 2–4 weeks)', C_RED, immediate))
story.append(Spacer(1, 12))

# SHORT-TERM actions
short_term = [
    ('MEDIUM',
     f'Populate SW Key Contact for {total - with_contact} items without an assigned contact',
     f'Currently {with_contact} of {total} items ({pct(with_contact,total)}%) have an SW Key Contact. '
     f'The remaining {total - with_contact} items are unowned. Assign contacts across SWS teams '
     f'to ensure accountability for each product.'),
    ('MEDIUM',
     f'Set Required License Date for {total - with_req_date} items',
     f'Only {with_req_date} ({pct(with_req_date,total)}%) of items have a required date. '
     f'All procurement-relevant items should have a target date defined to support '
     f'scheduling and cutover planning.'),
    ('MEDIUM',
     'Standardise payment type classification across all vendors',
     'Ensure all products have a valid payment type (subscription, perpetual, freeware, OSS). '
     'Use a controlled vocabulary to prevent re-occurrence of "blank/TBD" entries. '
     'Align classification with the Bosch license procurement standards.'),
    ('MEDIUM',
     'Review and consolidate duplicate / replaced entries',
     f'{out_dup} entries are marked as duplicates and {replaced} as replaced. '
     f'Archive these in a separate tab and ensure the master list reflects only active items '
     f'to prevent confusion during procurement.'),
]

story.extend(action_section('🟠  Short-term Actions (Complete within 1–2 months)', C_ORANGE, short_term))
story.append(Spacer(1, 12))

# ONGOING actions
ongoing = [
    ('LOW',
     f'Audit {freeware + oss} Freeware / OSS products for compliance obligations',
     f'{freeware} products are classified as freeware and {oss} as Open Source Software. '
     f'Conduct a licence compliance review (GPL, MIT, Apache, etc.) to identify any '
     f'copyleft obligations or commercial-use restrictions before use in production.'),
    ('LOW',
     'Establish quarterly review cycle for the Astra license inventory',
     'Implement a recurring quarterly review to update product status, vendor changes, '
     'quantity adjustments, and new procurement requirements. Assign a DRI (Directly '
     'Responsible Individual) per SWS system.'),
    ('LOW',
     'Migrate Astra list to Microsoft List with structured procurement workflow',
     'The current Excel-based list lacks procurement tracking columns (owner, due date, '
     'approval status). Migrating to Microsoft List will enable filtering, alerting, '
     'and governance tracking across all 563 items.'),
    ('LOW',
     f'Increase notes / comments coverage (currently {pct(with_notes,total)}%)',
     f'Only {with_notes} of {total} items have notes. Critical context for procurement '
     f'decisions (special terms, negotiations, version restrictions) should be captured '
     f'systematically to support knowledge transfer and audits.'),
]

story.extend(action_section('🟡  Ongoing / Governance Actions', C_GREEN, ongoing))
story.append(Spacer(1, 16))

# ─ Summary action table
story.append(SectionHeader('Action Summary', C_DARK))
story.append(Spacer(1, 8))

sum_rows = [
    [Paragraph(h, s_th) for h in ['Priority', 'Action', 'Items Affected', 'Owner']],
    [Paragraph('HIGH',   ParagraphStyle('hp', fontName='Helvetica-Bold', fontSize=8, textColor=C_WHITE, alignment=TA_CENTER)),
     Paragraph('Resolve items in clarification', s_td), Paragraph(str(in_clarif), s_td_c), Paragraph('SWS Workstream Lead', s_td)],
    [Paragraph('HIGH',   ParagraphStyle('hp', fontName='Helvetica-Bold', fontSize=8, textColor=C_WHITE, alignment=TA_CENTER)),
     Paragraph('Classify blank/TBD payment types', s_td), Paragraph(str(tbd_cnt), s_td_c), Paragraph('Product Owner / Vendor Mgr', s_td)],
    [Paragraph('HIGH',   ParagraphStyle('hp', fontName='Helvetica-Bold', fontSize=8, textColor=C_WHITE, alignment=TA_CENTER)),
     Paragraph('Trigger procurement for "Ready" items', s_td), Paragraph(str(ready_proc), s_td_c), Paragraph('SWS Procurement Team', s_td)],
    [Paragraph('MEDIUM', ParagraphStyle('mp', fontName='Helvetica-Bold', fontSize=8, textColor=C_WHITE, alignment=TA_CENTER)),
     Paragraph('Assign SW Key Contact', s_td), Paragraph(str(total-with_contact), s_td_c), Paragraph('SWS Team Leads', s_td)],
    [Paragraph('MEDIUM', ParagraphStyle('mp', fontName='Helvetica-Bold', fontSize=8, textColor=C_WHITE, alignment=TA_CENTER)),
     Paragraph('Set Required License Dates', s_td), Paragraph(str(total-with_req_date), s_td_c), Paragraph('Project Manager', s_td)],
    [Paragraph('LOW',    ParagraphStyle('lp', fontName='Helvetica-Bold', fontSize=8, textColor=C_WHITE, alignment=TA_CENTER)),
     Paragraph('OSS compliance audit', s_td), Paragraph(str(freeware+oss), s_td_c), Paragraph('License / Legal Team', s_td)],
    [Paragraph('LOW',    ParagraphStyle('lp', fontName='Helvetica-Bold', fontSize=8, textColor=C_WHITE, alignment=TA_CENTER)),
     Paragraph('Migrate to Microsoft List', s_td), Paragraph(str(total), s_td_c), Paragraph('IT / Project Manager', s_td)],
]
t_sum = Table(sum_rows, colWidths=[1.8*cm, 6.5*cm, 2.5*cm, 5*cm])
t_sum.setStyle(TableStyle([
    ('BACKGROUND',  (0,0), (-1,0),  C_DARK),
    ('BACKGROUND',  (0,1), (0,3),   C_RED),
    ('BACKGROUND',  (0,4), (0,5),   C_ORANGE),
    ('BACKGROUND',  (0,6), (0,-1),  C_GRAY),
    ('GRID',        (0,0), (-1,-1), 0.5, colors.HexColor('#e0e0e0')),
    ('TOPPADDING',  (0,0), (-1,-1), 5),
    ('BOTTOMPADDING',(0,0),(-1,-1), 5),
    ('LEFTPADDING', (0,0), (-1,-1), 6),
    ('RIGHTPADDING',(0,0), (-1,-1), 6),
    ('ROWBACKGROUNDS', (0,1),(-1,-1), [C_WHITE, colors.HexColor('#f9f9f9')]),
    ('VALIGN',      (0,0), (-1,-1), 'MIDDLE'),
]))
story.append(t_sum)

# ─ Footer note
story.append(Spacer(1, 18))
story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#dddddd')))
story.append(Spacer(1, 6))
footer_p = Paragraph(
    'Astra Software License Management Report  ·  BDMIL AI Use Case  ·  April 2, 2026  '
    '·  Generated by GitHub Copilot  ·  Source: Astra Shoppping List with SWS.xlsx',
    ParagraphStyle('footer', fontName='Helvetica', fontSize=7.5, textColor=C_GRAY,
                   alignment=TA_CENTER, leading=10)
)
story.append(footer_p)

# ── Build PDF ──────────────────────────────────────────────────────────────
out_path = os.path.join(BASE, 'Astra_SWS_Management_Report.pdf')
doc = SimpleDocTemplate(
    out_path,
    pagesize=A4,
    leftMargin=2*cm,
    rightMargin=2*cm,
    topMargin=2*cm,
    bottomMargin=2.5*cm,
    title='Astra Software License Management Report',
    author='GitHub Copilot – BDMIL Use Case',
    subject='Software License Procurement Status – April 2026',
)
doc.build(story, onFirstPage=on_first_page, onLaterPages=on_later_pages)
print(f'Written: {out_path}  ({os.path.getsize(out_path)//1024} KB)')
