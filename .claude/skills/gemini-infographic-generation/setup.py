#!/usr/bin/env python3
"""
Setup script for Gemini Infographic Generation skill.
Verifies dependencies and API key configuration.
"""

import os
import sys
import subprocess
from pathlib import Path


def check_python_version() -> bool:
    """Verify Python 3.9+."""
    if sys.version_info < (3, 9):
        print(f"❌ Python 3.9+ required (current: {sys.version.split()[0]})")
        return False
    print(f"✓ Python {sys.version.split()[0]}")
    return True


def check_packages() -> bool:
    """Check required packages."""
    required = {
        "google.generativeai": "google-generativeai>=0.7.0",
        "PIL": "pillow>=10.0.0",
    }
    
    all_ok = True
    for module, package in required.items():
        try:
            __import__(module)
            print(f"✓ {package}")
        except ImportError:
            print(f"❌ {package} not installed")
            all_ok = False
    
    if not all_ok:
        print("\nInstalling missing packages...")
        subprocess.check_call([
            sys.executable, "-m", "pip", "install",
            "google-generativeai>=0.7.0", "pillow>=10.0.0"
        ])
        print("✓ Packages installed")
    
    return True


def check_api_key() -> bool:
    """Check/configure Google AI API key."""
    key = os.getenv("GOOGLE_API_KEY")
    
    if key:
        print(f"✓ GOOGLE_API_KEY set (starts with: {key[:8]}...)")
        return True
    
    print("\n⚠️  GOOGLE_API_KEY not set")
    print("Get a free key at https://aistudio.google.com/apikey")
    
    key_input = input("\nEnter your Google AI API key (or leave blank to skip): ").strip()
    
    if key_input:
        os.environ["GOOGLE_API_KEY"] = key_input
        
        # Try to persist to .env
        env_file = Path(".env")
        if env_file.exists():
            with open(env_file, "a") as f:
                f.write(f"\nGOOGLE_API_KEY={key_input}\n")
            print("✓ API key saved to .env")
        else:
            print("Create a .env file to persist your API key:")
            print(f"  echo GOOGLE_API_KEY={key_input[:8]}...{key_input[-4:]} > .env")
        
        return True
    
    print("⚠️  Skipped. You'll need to set GOOGLE_API_KEY before generating images.")
    return True


def main():
    print("=== Gemini Infographic Generation — Setup ===\n")
    
    checks = [
        ("Python version", check_python_version),
        ("Required packages", check_packages),
        ("Google AI API key", check_api_key),
    ]
    
    all_ok = True
    for name, check_fn in checks:
        print(f"\nChecking {name}...")
        try:
            if not check_fn():
                all_ok = False
        except Exception as e:
            print(f"❌ Error: {e}")
            all_ok = False
    
    print("\n" + "="*50)
    if all_ok:
        print("✓ Setup complete! Ready to generate infographics.")
        print("\nQuick start:")
        print("  python generate_infographic.py --interactive")
        print("  python generate_infographic.py --topic 'Q1 Sales Growth'")
    else:
        print("❌ Setup incomplete. Fix errors above and try again.")
        sys.exit(1)


if __name__ == "__main__":
    main()
