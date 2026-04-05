# Gemini Infographic Generation Skill

**Purpose:** Generate 1-pager marketing and communication infographics using Google's Gemini image generation models with Creative Director orchestration.

**Scope:** Marketing assets, communication materials, 1-page visual narratives outside the core carve-out project domain.

---

## Overview

This skill adapts the **banana-claude Creative Director pattern** to generate publication-ready infographics on demand. Unlike simple API wrappers, Claude analyzes your intent, selects the appropriate visual domain (Editorial, Infographic, Data Visualization), applies structured prompt engineering, and orchestrates Gemini for optimal results.

**Use Cases:**
- Executive 1-pagers for company announcements
- Marketing campaign visuals
- Communication assets for internal/external stakeholders
- Data visualization infographics

---

## Prerequisites

1. **Google AI API Key**
   - Obtain free at [aistudio.google.com/apikey](https://aistudio.google.com/apikey)
   - Set environment variable: `GOOGLE_API_KEY`
   - Free tier: ~5–15 requests/min; ~20–500 requests/day

2. **Python 3.9+**

3. **Required Packages**
   ```
   google-generativeai>=0.7.0
   pillow>=10.0.0
   ```

---

## Quick Start

### 1. Setup

```bash
# Install dependencies
pip install google-generativeai pillow

# Set your API key (Windows PowerShell)
$env:GOOGLE_API_KEY = "your-key-here"

# Or add to .env for persistence
echo "GOOGLE_API_KEY=your-key-here" > .env
```

### 2. Generate an Infographic

```bash
python generate_infographic.py \
  --topic "Q1 Sales Growth" \
  --style "editorial" \
  --output infographic_q1.png
```

### 3. Interactive Mode

```bash
python generate_infographic.py --interactive
```

---

## Domain Modes for Infographics

| Domain | Best For | Example Prompt |
|--------|----------|-----------------|
| **Infographic** | Data, statistics, processes | "Q1 revenue growth breakdown by region" |
| **Editorial** | Marketing, lifestyle, storytelling | "Our company transformation journey 2025–2026" |
| **UI/Web** | Clean diagrams, icons, illustrations | "Organizational restructure org chart" |
| **Product** | Product-centric visuals | "Our new SaaS platform feature highlights" |
| **Abstract** | Pattern-based, conceptual | "Innovation ecosystem infographic" |

---

## The Creative Director Pipeline

### Intent Analysis
Claude interprets your request:
- Topic (what data/story to convey)
- Audience (internal/external, technical/non-technical)
- Format (1-pager, poster, dashboard card)
- Visual style (corporate, modern, minimal, bold)

### Domain Selection
Based on intent, Claude selects the best visual mode:
```
topic + audience + format → {Infographic, Editorial, UI, Product, Abstract}
```

### 5-Component Prompt Formula

Instead of "our sales grew 20%", Claude constructs:

> A clean corporate infographic using a 16:9 format showing Q1 revenue growth of 20% with large typography for "Q1 2026: +20% Revenue Growth". Left side: bar chart with regional breakdown (EMEA 35%, APAC 42%, Americas 23%) in Bosch blue (#003B6E), green (#66BB6A) for growth. Right side: key metrics in modern sans-serif with icons. White background, minimal spacing, ready for presentation slides.

**5 Components:**
1. **Subject** – The core data/story
2. **Action** – What's being shown/compared
3. **Location/Context** – Where it appears, what audience
4. **Composition** – Layout, aspect ratio, information hierarchy
5. **Style** – Color palette, typography, brand alignment

### Prompt Adaptation
Claude translates patterns from a curated database to Gemini's natural language format, accounting for:
- Brand guidelines (Bosch blue `#003B6E`, accent `#0066CC`)
- Corporate standards
- Platform requirements (slide ratio, web size, print dimensions)

### Batch Variations
Generate N variations with rotated components:
```
--batch 3
```

Produces 3 variations with different:
- Regional layout orders
- Color emphasis
- Icon styles
- Chart type options

---

## Commands

### Direct Generation

```bash
python generate_infographic.py \
  --topic "Carve-out project timeline" \
  --style "infographic" \
  --domain "infographic" \
  --output timeline.png
```

### Parameters

| Param | Type | Default | Notes |
|-------|------|---------|-------|
| `--topic` | str | – | **Required** – The data/story to visualize |
| `--style` | str | `corporate` | Visual style: `corporate`, `modern`, `minimal`, `bold` |
| `--domain` | str | `infographic` | Domain mode: `infographic`, `editorial`, `ui`, `product`, `abstract` |
| `--aspect-ratio` | str | `16:9` | `16:9`, `1:1`, `4:3`, `9:16`, `21:9` |
| `--color-palette` | str | `bosch-blue` | `bosch-blue`, `neutral`, `vibrant`, `monochrome` |
| `--audience` | str | `external` | `internal`, `executive`, `external`, `customer` |
| `--output` | str | `infographic.png` | Output file path |
| `--batch` | int | 1 | Generate N variations |
| `--interactive` | flag | False | Guided Q&A mode |

### Interactive Mode

```bash
python generate_infographic.py --interactive
```

Claude will guide you through:
1. Topic & data points
2. Intended audience
3. Brand/style preferences
4. Format & platform
5. Generate + review variations

---

## Cost Tracking

Gemini Flash 3.1 (default):
- **Cost:** ~$0.075–0.15 per image (generation only)
- **Resolution:** up to 4K (4096×4096)
- **Aspect Ratios:** 14 standard + custom

Track cumulative usage:

```bash
python cost_tracker.py --summary
```

---

## Brand Presets

Define reusable brand/style presets for consistency:

### Bosch Corporate Preset

```yaml
name: bosch-corporate
colors:
  primary: "#003B6E"      # Bosch Blue
  accent: "#0066CC"       # Accent Blue
  success: "#66BB6A"      # Green
  warning: "#FFA726"      # Amber
  error: "#E20015"        # Bosch Red (sparingly)
typography:
  primary: "Arial, sans-serif"
  weight: "500–600"
style: "corporate, clean, minimal"
tone: "professional, trustworthy"
```

### Marketing Preset

```yaml
name: marketing-bold
colors:
  primary: "#0066CC"
  accent: "#E20015"
  tertiary: "#66BB6A"
typography:
  primary: "Montserrat, sans-serif"
  weight: "700–900"
style: "modern, bold, dynamic"
tone: "energetic, innovative"
```

---

## Workflow: 1-Pager Infographic

**Typical flow for a marketing 1-pager:**

1. **Define intent**
   ```
   Topic: "2026 CarveOut Carveout Results"
   Audience: Board of Directors
   Format: Full-page PDF export (8.5"×11")
   ```

2. **Generate with brand preset**
   ```bash
   python generate_infographic.py \
     --topic "2026 CarveOut Results: 20% Cost Savings, 3-Month Timeline" \
     --preset bosch-corporate \
     --aspect-ratio 8.5:11 \
     --output Board_Summary_2026.png
   ```

3. **Review variations** (3 versions)
   ```bash
   python generate_infographic.py \
     --topic "2026 CarveOut Results..." \
     --preset bosch-corporate \
     --batch 3 \
     --output-dir ./variations/
   ```

4. **Select + export**
   - Pick best version from 3 variations
   - Export to PDF for distribution
   - Embed in presentations/emails

---

## Safety & Compliance

- **Banned Keywords:** Avoid political/religious content, trademarked names (unless licensed), harmful imagery
- **Data Privacy:** Never include PII, confidential financial data, or proprietary information in generated assets
- **Gemini Auto-Filtering:** Platform blocks NSFW, violence, hate speech; Claude re-phrases rejected prompts transparently

---

## Post-Processing (Optional)

Enhance generated images:

```bash
python enhance_infographic.py \
  --input infographic.png \
  --crop 16:9 \
  --watermark "© 2026 Bosch" \
  --output infographic_final.png
```

---

## API Reference

### `generate_infographic()`

```python
from generate_infographic import create_infographic

image = create_infographic(
    topic="Q1 Performance Dashboard",
    domain="infographic",
    style="corporate",
    color_palette="bosch-blue",
    aspect_ratio="16:9",
    audience="executive"
)

image.save("output.png")
```

### `batch_generate()`

```python
images = batch_generate(
    topic="Sales forecast 2026",
    variations=3,
    preset="bosch-corporate"
)

for i, img in enumerate(images):
    img.save(f"variation_{i+1}.png")
```

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| `API key not found` | Set `GOOGLE_API_KEY` environment variable |
| `Rate limit exceeded` | Free tier: 5–15 RPM max. Wait 60 seconds, retry |
| `Image quality poor` | Increase `--aspect-ratio` detail or use `--batch` to try alternatives |
| `Prompt rejected` | Reword topic; avoid banned keywords (politics, trademarks, harm) |
| `Timeout` | API max response time ~60s; simplify topic or split into multiple infographics |

---

## References

- **banana-claude** (source): https://github.com/AgriciDaniel/banana-claude
- **Google Gemini API Docs**: https://ai.google.dev
- **Prompt Engineering Guide**: https://platform.openai.com/docs/guides/prompt-engineering
- **Bosch Brand Colors**: See `references/bosch-brand.md` in workspace

---

## Version

**v1.0** – Initial integration of banana-claude Creative Director pipeline for infographic generation.

