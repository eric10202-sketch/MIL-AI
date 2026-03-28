# Project Trinity PDF Report Generator

## Overview
Python script that generates a professional 1-page PDF status report for Project Trinity (IT Carve-Out: Bosch → Keenfinity).

## Files
- **`generate_trinity_pdf_report.py`** — Main script to generate the PDF report
- **`Trinity_Status_Report_20260328.pdf`** — Sample generated report

## Requirements
- Python 3.7+
- reportlab library

## Installation

### 1. Set up Python environment (if not already done):
```bash
cd /Users/erikho/Desktop/Carveout
python3 -m venv .venv
source .venv/bin/activate
```

### 2. Install dependencies:
```bash
pip install reportlab
```

## Usage

### Quick Start (default filename):
```bash
python3 generate_trinity_pdf_report.py
```
This creates `Trinity_Status_Report.pdf`

### With custom filename:
```bash
python3 generate_trinity_pdf_report.py My_Custom_Report.pdf
```

### Example with timestamp:
```bash
python3 generate_trinity_pdf_report.py Trinity_Status_Report_$(date +%Y%m%d).pdf
```

## Report Contents

The generated 1-page PDF includes:

1. **Header** — Project name, date, status
2. **Key Metrics** — Report date, status, duration, go-live, scope, budget
3. **Phase Status & Timeline** — 6 project phases with key deliverables
4. **Key Risks & Immediate Actions** — Top 5 risks and decisions required
5. **Budget Snapshot** — Labour cost breakdown (€5.1M total)
6. **Critical Milestones** — Key dates and gates (QG1, QG2, Day 1 Go-Live)
7. **Footer** — Distribution, confidentiality, next review date

## Customization

To modify report content, edit the following sections in `generate_trinity_pdf_report.py`:

- **metrics_data** — Update key metrics table (line ~120)
- **phase_data** — Update phase timeline (line ~135)
- **risks_text / decisions_text** — Update risks and decisions (line ~175)
- **budget_data** — Update budget breakdown (line ~200)
- **milestones_data** — Update critical milestones (line ~225)

## Output Format

- **Page Size:** Letter (8.5" × 11")
- **Margins:** 0.5" on all sides
- **Color Scheme:** Professional (purple primary, green/amber/red for status)
- **File Format:** PDF (embeddable fonts, high quality)

## Automation

To generate reports automatically on a schedule:

### macOS/Linux (Cron job):
```bash
# Add to crontab (e.g., every Monday at 9am)
0 9 * * 1 cd /Users/erikho/Desktop/Carveout && /Users/erikho/Desktop/Carveout/.venv/bin/python3 generate_trinity_pdf_report.py Trinity_Status_Report_$(date +\%Y\%m\%d).pdf
```

### PowerShell (Windows):
```powershell
# Schedule using Task Scheduler to run:
python generate_trinity_pdf_report.py Trinity_Status_Report_%date:~-4,4%%date:~-10,2%%date:~-7,2%.pdf
```

## Troubleshooting

### ModuleNotFoundError: reportlab
```bash
# Ensure you're using the venv Python:
/Users/erikho/Desktop/Carveout/.venv/bin/python generate_trinity_pdf_report.py
```

### File Permission Error
```bash
# Make script executable:
chmod +x generate_trinity_pdf_report.py
```

### PDF Output is blank or incomplete
- Check Python version (3.7+)
- Verify reportlab is installed: `pip list | grep reportlab`
- Run with verbose output and check for errors

## Sample Output
```
$ python3 generate_trinity_pdf_report.py Trinity_Status_Report_20260328.pdf
Generating Project Trinity Status Report...
Output: Trinity_Status_Report_20260328.pdf
✓ PDF report created successfully: Trinity_Status_Report_20260328.pdf

📄 Report saved to: Trinity_Status_Report_20260328.pdf
```

## Related Documents

- **Trinity_Executive_Status_Report.md** — Detailed Markdown version
- **Trinity_Status_Dashboard.html** — Interactive HTML dashboard
- **Trinity_Steering_Committee_Brief.md** — 1-page brief for steering
- **Trinity_Stakeholder_Communications_Plan.md** — Full communications strategy

## Support

For questions or modifications:
1. Edit the script directly (Python beginner-friendly)
2. Refer to [ReportLab Documentation](https://www.reportlab.com/docs/reportlab-userguide.pdf)
3. Copy and modify for other projects

---

**Last Updated:** March 28, 2026  
**Script Version:** 1.0  
**Status:** Production-ready
