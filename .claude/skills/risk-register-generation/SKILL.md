---
name: risk-register-generation
description: "Use when generating carve-out risk registers with probability x impact scoring, risk categories, mitigation plans, ownership, and top-risk prioritization."
---

# Risk Register Generation

## Purpose
Generate a project-specific risk register using carve-out risk patterns and templates.

## Canonical Sources
- generate_risk_register.py
- 20240718_FRAME_IT Risk Assessment_completed.csv
- Risk_analysis_template.xlsx

## Core Method
- Risk rating = Probability x Impact (1-5 scale each).
- Use category taxonomy: ScR, SR, RR, BtR, QR, BR, LR, CR.
- Include mitigation, owner, target date, and status per risk.

## Rules
- Keep risks specific to current engagement parties and scope.
- Do not copy historical project facts from reference files.
- Prioritize risks by rating and delivery impact.
- Ensure top risks align with schedule critical path and milestones.
