#!/usr/bin/env python3
"""
Gemini Infographic Generator
AI image generation skill for infographics using Google's Gemini models.
Implements the Creative Director orchestration pattern.
"""

import os
import sys
import json
import argparse
from pathlib import Path
from typing import Optional, List, Dict, Tuple
from PIL import Image
import io

try:
    from google import genai
    from google.genai import types as genai_types
    HAS_NEW_GENAI = True
except Exception:
    HAS_NEW_GENAI = False

# Configuration
GEMINI_MODELS = [
    "gemini-3.1-flash-image-preview",
    "gemini-2.5-flash-image",
    "imagen-4.0-fast-generate-001",
]
DEFAULT_ASPECT_RATIO = "16:9"
DEFAULT_STYLE = "corporate"
DEFAULT_DOMAIN = "infographic"

# Bosch Brand Colors
BOSCH_PALETTE = {
    "bosch-blue": {
        "primary": "#003B6E",
        "accent": "#0066CC",
        "success": "#66BB6A",
        "warning": "#FFA726",
        "error": "#E20015",
    },
    "neutral": {
        "primary": "#333333",
        "accent": "#666666",
        "success": "#4CAF50",
        "warning": "#FF9800",
        "error": "#F44336",
    },
    "vibrant": {
        "primary": "#FF6B6B",
        "accent": "#4ECDC4",
        "success": "#45B7D1",
        "warning": "#FFA07A",
        "error": "#FF6B6B",
    },
}

# Domain modes
DOMAIN_MODES = {
    "infographic": "Data-driven visualizations with charts, statistics, and key metrics",
    "editorial": "Marketing, lifestyle, storytelling visuals with narrative arc",
    "ui": "Clean diagrams, icons, illustrations, organizational charts",
    "product": "Product-centric visuals, feature highlights, packshots",
    "abstract": "Pattern-based, conceptual, generative art, textures",
}

# Aspect ratio mapping
ASPECT_RATIOS = {
    "16:9": (1920, 1080),
    "1:1": (1024, 1024),
    "4:3": (1400, 1050),
    "9:16": (1080, 1920),
    "21:9": (2400, 1024),
    "8.5:11": (1024, 1331),  # Letter aspect
}


def construct_creative_brief(
    topic: str,
    domain: str,
    style: str,
    color_palette: str,
    aspect_ratio: str,
    audience: str,
) -> str:
    """
    Construct a detailed 5-component prompt using the Creative Director pipeline.

    Components:
    1. Subject – Core data/story
    2. Action – What's being shown
    3. Location/Context – Where it appears
    4. Composition – Layout, hierarchy
    5. Style – Colors, typography, brand
    """
    
    colors = BOSCH_PALETTE.get(color_palette, BOSCH_PALETTE["bosch-blue"])
    
    # Aspect ratio guidance
    ratio_guidance = {
        "16:9": "landscape, wide format for presentations",
        "1:1": "square, balanced composition for social media",
        "8.5:11": "portrait, full-page 1-pager format",
        "21:9": "ultra-wide, cinematic banner",
    }
    ratio_desc = ratio_guidance.get(aspect_ratio, "standard proportion")
    
    # Audience-specific context
    audience_context = {
        "executive": "high-level overview, key metrics, strategic focus",
        "internal": "team-focused insights, detailed data, operational focus",
        "external": "marketing narrative, customer benefits, polished finish",
        "customer": "product-focused, use-case emphasis, professional quality",
    }
    audience_desc = audience_context.get(audience, "general audience")
    
    # Domain description
    domain_desc = DOMAIN_MODES.get(domain, DOMAIN_MODES["infographic"])
    
    # Construct the 5-component brief
    brief = f"""You are a Creative Director for professional infographic design.

OBJECTIVE: Generate a publication-ready infographic with the following specifications:

1. SUBJECT (Core Data/Story)
   Topic: {topic}
   Domain: {domain_desc}

2. ACTION (What's Being Conveyed)
   - Present data clearly with hierarchy and emphasis
   - Highlight key insights visually
   - Use visual metaphors for abstract concepts

3. LOCATION/CONTEXT (Where This Will Appear)
   Audience: {audience_desc}
   Format: {ratio_desc}
   Platform: Professional document, presentation slide, or marketing material

4. COMPOSITION (Layout & Information Hierarchy)
   - Use clear geometric grid (3–5 column layout)
   - Top: Headline with large typography
   - Middle: Main data visualization (charts, icons, diagrams)
   - Bottom: Supporting metrics or call-to-action
   - Whitespace for readability

5. STYLE (Visual Identity)
   - Primary color: {colors['primary']} (Bosch Blue or equivalent)
   - Accent color: {colors['accent']}
   - Secondary colors: {colors['success']} (growth), {colors['warning']} (caution), {colors['error']} (critical)
   - Typography: Clean sans-serif (Arial, Helvetica, or sans-serif equivalent)
   - Font weight: 500–700 for hierarchy
   - Visual style: {style} (corporate, professional, modern)
   - Background: White or light gray for contrast
   - No watermarks or branding text

TECHNICAL REQUIREMENTS:
- Aspect ratio: {aspect_ratio}
- Resolution: High-quality, print-ready (minimum 150 DPI equivalent)
- Format: PNG with transparent background where appropriate
- Ensure all text is legible and sharp

Generate a single, cohesive infographic that follows these 5 components exactly.
Focus on visual clarity, professional appearance, and data-driven storytelling."""
    
    return brief


def generate_infographic(
    topic: str,
    domain: str = DEFAULT_DOMAIN,
    style: str = DEFAULT_STYLE,
    color_palette: str = "bosch-blue",
    aspect_ratio: str = DEFAULT_ASPECT_RATIO,
    audience: str = "external",
) -> Image.Image:
    """Generate an infographic using Gemini API."""
    
    api_key = os.getenv("GOOGLE_API_KEY")
    if not api_key:
        raise ValueError(
            "GOOGLE_API_KEY environment variable not set. "
            "Get a free key at https://aistudio.google.com/apikey"
        )
    
    # Construct the creative brief
    brief = construct_creative_brief(
        topic=topic,
        domain=domain,
        style=style,
        color_palette=color_palette,
        aspect_ratio=aspect_ratio,
        audience=audience,
    )
    
    if not HAS_NEW_GENAI:
        raise RuntimeError(
            "google-genai package is required. Install with: py -m pip install google-genai"
        )

    client = genai.Client(api_key=api_key)
    last_error = None

    for model_name in GEMINI_MODELS:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=brief,
                config=genai_types.GenerateContentConfig(
                    response_modalities=["IMAGE", "TEXT"],
                ),
            )

            for candidate in getattr(response, "candidates", []) or []:
                for part in getattr(candidate.content, "parts", []) or []:
                    inline_data = getattr(part, "inline_data", None)
                    if inline_data and getattr(inline_data, "data", None):
                        return Image.open(io.BytesIO(inline_data.data))
        except Exception as exc:
            last_error = exc
            continue

    raise RuntimeError(
        f"Gemini failed to return image bytes across models {GEMINI_MODELS}. Last error: {last_error}"
    )


def batch_generate(
    topic: str,
    variations: int = 3,
    domain: str = DEFAULT_DOMAIN,
    style: str = DEFAULT_STYLE,
    color_palette: str = "bosch-blue",
    aspect_ratio: str = DEFAULT_ASPECT_RATIO,
    audience: str = "external",
) -> List[Image.Image]:
    """Generate multiple variations of an infographic."""
    
    images = []
    for i in range(variations):
        print(f"Generating variation {i+1}/{variations}...", file=sys.stderr)
        try:
            img = generate_infographic(
                topic=topic,
                domain=domain,
                style=style,
                color_palette=color_palette,
                aspect_ratio=aspect_ratio,
                audience=audience,
            )
            images.append(img)
        except Exception as e:
            print(f"Error generating variation {i+1}: {e}", file=sys.stderr)
            continue
    
    return images


def interactive_mode() -> None:
    """Interactive guided Q&A for infographic generation."""
    
    print("\n=== Gemini Infographic Generator (Interactive Mode) ===\n")
    
    # Topic
    topic = input("What is the topic/data for this infographic? ").strip()
    if not topic:
        print("Topic required. Exiting.")
        return
    
    # Domain
    print("\nWhat type of infographic?")
    for i, (key, desc) in enumerate(DOMAIN_MODES.items(), 1):
        print(f"  {i}. {key.capitalize()} – {desc}")
    domain_choice = input("Select (1–5) [default: 1]: ").strip() or "1"
    domain = list(DOMAIN_MODES.keys())[int(domain_choice) - 1]
    
    # Audience
    audiences = ["internal", "executive", "external", "customer"]
    print("\nTarget audience?")
    for i, aud in enumerate(audiences, 1):
        print(f"  {i}. {aud.capitalize()}")
    aud_choice = input("Select (1–4) [default: 3]: ").strip() or "3"
    audience = audiences[int(aud_choice) - 1]
    
    # Style
    styles = ["corporate", "modern", "minimal", "bold"]
    print("\nVisual style?")
    for i, st in enumerate(styles, 1):
        print(f"  {i}. {st.capitalize()}")
    style_choice = input("Select (1–4) [default: 1]: ").strip() or "1"
    style = styles[int(style_choice) - 1]
    
    # Color palette
    print("\nColor palette?")
    for i, pal in enumerate(BOSCH_PALETTE.keys(), 1):
        print(f"  {i}. {pal.replace('-', ' ').title()}")
    color_choice = input("Select (1–3) [default: 1]: ").strip() or "1"
    color_palette = list(BOSCH_PALETTE.keys())[int(color_choice) - 1]
    
    # Aspect ratio
    print("\nAspect ratio?")
    for i, ratio in enumerate(ASPECT_RATIOS.keys(), 1):
        print(f"  {i}. {ratio}")
    ratio_choice = input("Select (1–5) [default: 1]: ").strip() or "1"
    aspect_ratio = list(ASPECT_RATIOS.keys())[int(ratio_choice) - 1]
    
    # Number of variations
    batch_count = input("\nHow many variations to generate? (1–5) [default: 1]: ").strip() or "1"
    batch_count = int(batch_count)
    
    # Output directory
    output_dir = input("Output directory? [default: ./infographics/]: ").strip() or "./infographics/"
    Path(output_dir).mkdir(parents=True, exist_ok=True)
    
    # Generate
    print(f"\nGenerating {batch_count} variation(s)...")
    images = batch_generate(
        topic=topic,
        variations=batch_count,
        domain=domain,
        style=style,
        color_palette=color_palette,
        aspect_ratio=aspect_ratio,
        audience=audience,
    )
    
    # Save
    for i, img in enumerate(images, 1):
        filename = f"infographic_var{i}.png"
        filepath = os.path.join(output_dir, filename)
        img.save(filepath)
        print(f"Saved: {filepath}")
    
    print(f"\nGenerated {len(images)} infographic(s).")


def main():
    parser = argparse.ArgumentParser(
        description="Generate 1-pager infographics using Google's Gemini API with Creative Director orchestration."
    )
    
    parser.add_argument("--topic", type=str, help="Topic/data for the infographic")
    parser.add_argument("--domain", type=str, default=DEFAULT_DOMAIN, 
                        choices=DOMAIN_MODES.keys(),
                        help="Visual domain mode")
    parser.add_argument("--style", type=str, default=DEFAULT_STYLE,
                        choices=["corporate", "modern", "minimal", "bold"],
                        help="Visual style")
    parser.add_argument("--color-palette", type=str, default="bosch-blue",
                        choices=BOSCH_PALETTE.keys(),
                        help="Color palette")
    parser.add_argument("--aspect-ratio", type=str, default=DEFAULT_ASPECT_RATIO,
                        choices=ASPECT_RATIOS.keys(),
                        help="Aspect ratio")
    parser.add_argument("--audience", type=str, default="external",
                        choices=["internal", "executive", "external", "customer"],
                        help="Target audience")
    parser.add_argument("--batch", type=int, default=1, help="Number of variations to generate")
    parser.add_argument("--output", type=str, default="infographic.png", help="Output file path")
    parser.add_argument("--output-dir", type=str, help="Output directory for batch mode")
    parser.add_argument("--interactive", action="store_true", help="Interactive guided mode")
    
    args = parser.parse_args()
    
    # Interactive mode
    if args.interactive:
        interactive_mode()
        return
    
    # Direct generation mode
    if not args.topic:
        parser.print_help()
        print("\nError: --topic is required. Use --interactive for guided mode.")
        sys.exit(1)
    
    # Generate
    if args.batch > 1:
        # Batch mode
        output_dir = args.output_dir or "./infographics/"
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        
        print(f"Generating {args.batch} variation(s)...")
        images = batch_generate(
            topic=args.topic,
            variations=args.batch,
            domain=args.domain,
            style=args.style,
            color_palette=args.color_palette,
            aspect_ratio=args.aspect_ratio,
            audience=args.audience,
        )
        
        for i, img in enumerate(images, 1):
            filename = f"variation_{i}.png"
            filepath = os.path.join(output_dir, filename)
            img.save(filepath)
            print(f"Saved: {filepath}")
    else:
        # Single image mode
        print(f"Generating infographic: {args.topic}")
        img = generate_infographic(
            topic=args.topic,
            domain=args.domain,
            style=args.style,
            color_palette=args.color_palette,
            aspect_ratio=args.aspect_ratio,
            audience=args.audience,
        )
        img.save(args.output)
        print(f"Saved: {args.output}")


if __name__ == "__main__":
    main()
