#!/usr/bin/env python3
"""
Project Trinity - Generate 1-Page PDF Status Report

This script creates a professional 1-page PDF status report for Project Trinity
with key metrics, phase status, risks, and immediate actions.

Requirements:
    - reportlab: pip install reportlab
    
Usage:
    python3 generate_trinity_pdf_report.py [output_file]
    
Example:
    python3 generate_trinity_pdf_report.py Trinity_Status_Report_20260328.pdf
"""

import sys
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor, white, black
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib import colors

# Define colors for the report
COLOR_PRIMARY = HexColor("#667eea")
COLOR_SUCCESS = HexColor("#10b981")
COLOR_WARNING = HexColor("#f59e0b")
COLOR_DANGER = HexColor("#ef4444")
COLOR_LIGHT = HexColor("#f8f9fa")
COLOR_BORDER = HexColor("#e5e7eb")
COLOR_TEXT = HexColor("#333333")

def create_pdf_report(output_file="Trinity_Status_Report.pdf"):
    """
    Create a 1-page PDF status report for Project Trinity.
    
    Args:
        output_file (str): Path to save the PDF report
    """
    
    # Create PDF document
    doc = SimpleDocTemplate(
        output_file,
        pagesize=letter,
        rightMargin=0.5*inch,
        leftMargin=0.5*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch
    )
    
    # Container for report elements
    story = []
    
    # Define styles
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        textColor=COLOR_PRIMARY,
        spaceAfter=3,
        alignment=TA_LEFT,
        fontName='Helvetica-Bold'
    )
    
    heading2_style = ParagraphStyle(
        'CustomHeading2',
        parent=styles['Heading2'],
        fontSize=10,
        textColor=COLOR_PRIMARY,
        spaceAfter=6,
        spaceBefore=6,
        alignment=TA_LEFT,
        fontName='Helvetica-Bold'
    )
    
    body_style = ParagraphStyle(
        'CustomBody',
        parent=styles['Normal'],
        fontSize=8.5,
        leading=10,
        textColor=COLOR_TEXT
    )
    
    # Report date
    report_date = datetime.now().strftime("%B %d, %Y")
    
    # Header
    story.append(Paragraph("PROJECT TRINITY — EXECUTIVE STATUS REPORT", title_style))
    story.append(Paragraph("IT Carve-Out: Bosch → Keenfinity (JCI)", body_style))
    story.append(Spacer(1, 0.08*inch))
    
    # Key metrics table - 6 columns across
    metrics_data = [
        [
            "Report Date",
            "Status",
            "Duration",
            "Go-Live",
            "Scope",
            "Budget"
        ],
        [
            f"{report_date}",
            "⚠️ PRE-LAUNCH",
            "24 months",
            "Nov 2027",
            "8K emp | 180 sites",
            "€5.1M"
        ]
    ]
    
    metrics_table = Table(metrics_data, colWidths=[1.2*inch, 1.0*inch, 1.0*inch, 1.0*inch, 1.2*inch, 0.9*inch])
    metrics_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), COLOR_PRIMARY),
        ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTSIZE', (0, 1), (-1, 1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, COLOR_BORDER),
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), [white, COLOR_LIGHT]),
    ]))
    story.append(metrics_table)
    story.append(Spacer(1, 0.08*inch))
    
    # Phase Status Section
    story.append(Paragraph("📅 PHASE STATUS & TIMELINE", heading2_style))
    
    phase_data = [
        ["Phase", "Timeline", "Status", "Key Deliverable"],
        ["Phase 1: INITIALIZATION", "Jul–Sep 2026", "Ready", "Governance, IT inventory (180 sites)"],
        ["Phase 2: CONCEPT", "Oct 2026–Mar 2027", "Planned", "IT architecture, migration strategy"],
        ["Phase 3: DEVELOPMENT", "Apr–Jul 2027", "Planned", "3 regional DCs, Merger Zone online"],
        ["Phase 4: IMPLEMENTATION", "Aug–Oct 2027", "Planned", "Wave migrations, WAN cutover"],
        ["🎯 Phase 5: GO-LIVE", "Nov 1, 2027", "CRITICAL", "Day 1 cutover, independent ops"],
        ["Phase 6: STABILIZATION", "Dec 2027–Feb 2028", "Planned", "Hypercare (3 mo), TSA exit"],
    ]
    
    phase_table = Table(phase_data, colWidths=[1.6*inch, 1.4*inch, 0.8*inch, 2.4*inch])
    phase_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), COLOR_PRIMARY),
        ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (1, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 5), (0, 5), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7.5),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.5, COLOR_BORDER),
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), [COLOR_PRIMARY, white, white, white, white, HexColor("#fff3cd"), white]),
        ('TEXTCOLOR', (0, 5), (-1, 5), COLOR_DANGER),
    ]))
    story.append(phase_table)
    story.append(Spacer(1, 0.07*inch))
    
    # Top 5 Risks and Decisions Required (2 columns)
    story.append(Paragraph("🔴 KEY RISKS & IMMEDIATE ACTIONS", heading2_style))
    
    # Create two-column layout for risks and decisions
    risks_text = """
    <b>Top Risks (Rating):</b><br/>
    • R1: Schedule Compression (25 HIGH)<br/>
    • R2: Scope Creep (20 MEDIUM)<br/>
    • R3: Data Separation Errors (25 CRITICAL)<br/>
    • R4: Country-Specific Delays (20 MEDIUM)<br/>
    • R5: Resource Availability (15 MEDIUM)
    """
    
    decisions_text = """
    <b>Decisions Required (by Apr 15):</b><br/>
    ✓ Approve Project Charter & Governance<br/>
    ✓ Authorize €5.1M Labour Budget + HW/SW<br/>
    ✓ Greenlight KPMG Engagement (Apr 10)<br/>
    ✓ Confirm Phase 1 Kickoff (Jul 1, 2026)
    """
    
    risk_decision_table = Table([
        [Paragraph(risks_text, body_style), Paragraph(decisions_text, body_style)]
    ], colWidths=[3.2*inch, 3.2*inch])
    
    risk_decision_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, COLOR_BORDER),
        ('BACKGROUND', (0, 0), (0, -1), COLOR_LIGHT),
        ('BACKGROUND', (1, 0), (1, -1), COLOR_LIGHT),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
    ]))
    story.append(risk_decision_table)
    story.append(Spacer(1, 0.06*inch))
    
    # Budget Snapshot
    story.append(Paragraph("💰 BUDGET SNAPSHOT (Labour Only)", heading2_style))
    
    budget_data = [
        ["Category", "Cost", "Share"],
        ["Governance/PMO", "€633.6K", "12.4%"],
        ["Bosch IT Team", "€1.95M", "38.2%"],
        ["JCI IT Team", "€1.2M", "23.5%"],
        ["KPMG Consulting", "€1.3M", "25.5%"],
        ["TOTAL", "€5.1M", "100%"],
    ]
    
    budget_table = Table(budget_data, colWidths=[2.5*inch, 1.5*inch, 1.2*inch])
    budget_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), COLOR_PRIMARY),
        ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('BACKGROUND', (0, 4), (-1, 4), COLOR_WARNING),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 4), (-1, 4), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.5, COLOR_BORDER),
        ('ROWBACKGROUNDS', (0, 0), (-1, -1), [COLOR_PRIMARY, white, white, white, white, COLOR_WARNING]),
    ]))
    story.append(budget_table)
    story.append(Spacer(1, 0.06*inch))
    
    # Critical Milestones
    story.append(Paragraph("📊 CRITICAL MILESTONES", heading2_style))
    
    milestones_data = [
        ["Milestone", "Date", "Status", "Days to Gate"],
        ["Charter & Budget Approval", "Apr 15, 2026", "⏳ Due", "18"],
        ["⭐ QG1 — Initialization Sign-Off", "Sep 30, 2026", "⚠️ Critical", "186"],
        ["⭐ QG2 — Concept Phase Sign-Off", "Mar 31, 2027", "⚠️ Critical", "368"],
        ["🎯 DAY 1 GO-LIVE", "Nov 1, 2027", "🚨 CRITICAL", "583"],
        ["Project Closure", "Feb 28, 2028", "⏳ Future", "671"],
    ]
    
    milestones_table = Table(milestones_data, colWidths=[2.2*inch, 1.3*inch, 1.0*inch, 1.0*inch])
    milestones_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), COLOR_PRIMARY),
        ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('BACKGROUND', (0, 3), (-1, 3), HexColor("#fff3cd")),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 3), (0, 3), 'Helvetica-Bold'),
        ('FONTNAME', (0, 4), (0, 4), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.5, COLOR_BORDER),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    story.append(milestones_table)
    story.append(Spacer(1, 0.06*inch))
    
    # Footer
    footer_text = f"""
    <b>Distribution:</b> Bosch Executive Leadership, JCI Board, Steering Committee, IT Leadership &nbsp;&nbsp;&nbsp;
    <b>Confidentiality:</b> Internal Use Only &nbsp;&nbsp;&nbsp;
    <b>Next Review:</b> April 15, 2026
    """
    footer_style = ParagraphStyle(
        'CustomFooter',
        parent=styles['Normal'],
        fontSize=7,
        leading=9,
        textColor=HexColor("#999999"),
        alignment=TA_CENTER
    )
    story.append(Spacer(1, 0.04*inch))
    story.append(Paragraph(footer_text, footer_style))
    
    # Build PDF
    try:
        doc.build(story)
        print(f"✓ PDF report created successfully: {output_file}")
        return True
    except Exception as e:
        print(f"✗ Error creating PDF: {str(e)}")
        return False


def main():
    """Main entry point."""
    
    # Get output filename from command line or use default
    if len(sys.argv) > 1:
        output_file = sys.argv[1]
    else:
        output_file = "Trinity_Status_Report.pdf"
    
    # Ensure .pdf extension
    if not output_file.endswith('.pdf'):
        output_file += '.pdf'
    
    print(f"Generating Project Trinity Status Report...")
    print(f"Output: {output_file}")
    
    success = create_pdf_report(output_file)
    
    if success:
        print(f"\n📄 Report saved to: {output_file}")
        sys.exit(0)
    else:
        print(f"\n❌ Failed to create report")
        sys.exit(1)


if __name__ == "__main__":
    main()
