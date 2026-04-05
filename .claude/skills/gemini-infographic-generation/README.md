# Gemini Infographic Generation Skill

**AI image generation for 1-pager marketing infographics using Google's Gemini models.**

This skill provides a Creative Director orchestration pattern that interprets your intent, selects domain expertise, and generates publication-ready infographics.

---

## Quick Start

### 1. Setup (one-time)

```bash
cd .claude/skills/gemini-infographic-generation/
python setup.py
```

Provides your Google API key (free at https://aistudio.google.com/apikey).

### 2. Generate an Infographic

```bash
# Interactive mode (guided Q&A)
python generate_infographic.py --interactive

# Direct mode
python generate_infographic.py --topic "Q1 Sales Performance" --style corporate

# Batch: 3 variations
python generate_infographic.py --topic "2026 Results" --batch 3 --output-dir ./variations/
```

### 3. Track Costs

```bash
python cost_tracker.py summary
python cost_tracker.py today
```

---

## Key Features

✅ **Creative Director Pipeline** – Understands intent, selects domain, applies prompt engineering
✅ **5-Component Prompts** – Subject, Action, Context, Composition, Style
✅ **Domain Modes** – Infographic, Editorial, UI, Product, Abstract
✅ **Bosch Branding** – Pre-configured color palette and typography  
✅ **Batch Variations** – Generate N versions with rotated components  
✅ **Cost Tracking** – Monitor API usage and spending  

---

## Commands

| Command | Purpose |
|---------|---------|
| `generate_infographic.py --interactive` | Guided Q&A mode |
| `generate_infographic.py --topic "..."` | Generate single infographic |
| `generate_infographic.py --batch 3 --output-dir DIR` | Generate 3 variations |
| `cost_tracker.py summary` | View total cost summary |
| `cost_tracker.py today` | View today's usage |

---

## Parameters

```bash
--topic TEXT               # Required: what to visualize
--domain DOMAIN           # infographic, editorial, ui, product, abstract
--style STYLE             # corporate, modern, minimal, bold
--color-palette PALETTE   # bosch-blue, neutral, vibrant
--aspect-ratio RATIO      # 16:9, 1:1, 4:3, 9:16, 21:9, 8.5:11
--audience AUDIENCE       # internal, executive, external, customer
--batch N                 # Generate N variations
--output FILE             # Output file path
--interactive             # Interactive guided mode
```

---

## Use Cases

- **1-Pager Executive Summaries** – Board-level infographics with key metrics
- **Marketing Assets** – Campaign visuals, project highlights
- **Data Visualization** – Q1 Results, team structures, process flows
- **Communication Materials** – Internal announcements, customer updates

---

## Cost

**~$0.075–0.15 per image** (Gemini Flash 3.1)

Free tier: ~20 requests/day = ~$1.50/day

Track usage: `python cost_tracker.py summary`

---

## Documentation

See **SKILL.md** for detailed curriculum, API reference, troubleshooting, and workflow examples.

---

## Acknowledgments

Adapts the **banana-claude** Creative Director pattern from [@AgriciDaniel](https://github.com/AgriciDaniel/banana-claude).

