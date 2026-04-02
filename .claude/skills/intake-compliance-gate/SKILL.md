---
name: intake-compliance-gate
description: "Use when validating carve-out intake, confirming mandatory buyer/seller fields, blocking generation for missing inputs, and applying budget TBC exception handling."
---

# Intake and Compliance Gate

## Purpose
Apply non-negotiable engagement validation before any deliverable generation.

## Required Inputs
- Project name
- Seller
- Buyer
- Business being carved out
- Carve-out model (Stand Alone / Integration / Combination)
- PMO / methodology lead
- Number of worldwide sites
- Number of IT users

## Rules
- Never generate deliverables if mandatory inputs are missing.
- Ask all missing mandatory fields in one response.
- Never substitute benchmark values for sites/users.
- If budget is unknown and user explicitly confirms this, use: TBC - to be approved at QG1.
- Sponsor Customer = Buyer.
- Sponsor Contractor = Seller.

## Reference vs Active Engagement
- Reference project files are methodology-only.
- Never copy reference-specific parties, dates, or scope into active engagement output.

## Output Contract
Before downstream work, return either:
1. Validation passed with all required inputs listed.
2. Validation failed with a complete missing-fields list.
