#!/usr/bin/env python3
"""
Cost Tracker for Gemini Infographic Generation
Logs and summarizes API usage and costs.
"""

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

# Cost per image (Gemini Flash 3.1)
COST_PER_IMAGE = 0.075  # USD, approximate

class CostTracker:
    def __init__(self, log_file: str = ".gemini_usage.json"):
        self.log_file = log_file
        self.data = self.load_log()
    
    def load_log(self) -> Dict:
        """Load existing usage log."""
        if os.path.exists(self.log_file):
            with open(self.log_file, "r") as f:
                return json.load(f)
        return {"sessions": [], "total_images": 0, "total_cost": 0.0}
    
    def save_log(self) -> None:
        """Save usage log."""
        with open(self.log_file, "w") as f:
            json.dump(self.data, f, indent=2)
    
    def log_generation(self, topic: str, count: int = 1, notes: str = "") -> None:
        """Log an image generation session."""
        session = {
            "timestamp": datetime.now().isoformat(),
            "topic": topic,
            "images_generated": count,
            "cost_usd": round(count * COST_PER_IMAGE, 4),
            "notes": notes,
        }
        
        self.data["sessions"].append(session)
        self.data["total_images"] += count
        self.data["total_cost"] = round(self.data["total_cost"] + (count * COST_PER_IMAGE), 2)
        
        self.save_log()
    
    def get_today_summary(self) -> Dict:
        """Get today's usage."""
        today = datetime.now().date().isoformat()
        today_sessions = [s for s in self.data["sessions"] if s["timestamp"].startswith(today)]
        
        return {
            "date": today,
            "sessions": len(today_sessions),
            "images": sum(s["images_generated"] for s in today_sessions),
            "cost_usd": round(sum(s["cost_usd"] for s in today_sessions), 4),
        }
    
    def get_summary(self) -> Dict:
        """Get overall summary."""
        return {
            "total_sessions": len(self.data["sessions"]),
            "total_images": self.data["total_images"],
            "total_cost_usd": self.data["total_cost"],
            "cost_per_image_usd": COST_PER_IMAGE,
            "last_session": self.data["sessions"][-1] if self.data["sessions"] else None,
        }
    
    def print_summary(self) -> None:
        """Print formatted summary."""
        summary = self.get_summary()
        today = self.get_today_summary()
        
        print("\n=== Gemini Infographic Generation — Cost Summary ===\n")
        print(f"Total images generated: {summary['total_images']}")
        print(f"Total cost: ${summary['total_cost_usd']:.2f} USD")
        print(f"Cost per image: ${COST_PER_IMAGE:.3f} USD")
        print(f"Total sessions: {summary['total_sessions']}")
        
        if summary['last_session']:
            print(f"\nLast session: {summary['last_session']['timestamp']}")
            print(f"  Topic: {summary['last_session']['topic']}")
            print(f"  Images: {summary['last_session']['images_generated']}")
            print(f"  Cost: ${summary['last_session']['cost_usd']:.4f}")
        
        print(f"\nToday's usage ({today['date']}):")
        print(f"  Sessions: {today['sessions']}")
        print(f"  Images: {today['images']}")
        print(f"  Cost: ${today['cost_usd']:.2f}")
        
        # Free tier estimate
        print(f"\nFree tier estimate (20 req/day):")
        print(f"  Daily budget: ~$1.50 USD")
        print(f"  Today usage: ${today['cost_usd']:.2f} / $1.50")
        if today['cost_usd'] > 1.50:
            print(f"  ⚠️  WARNING: Daily quota likely exceeded")
        print()


def main():
    import argparse
    
    parser = argparse.ArgumentParser(description="Track Gemini API usage and costs")
    
    subparsers = parser.add_subparsers(dest="command", help="Command")
    
    subparsers.add_parser("summary", help="Show cost summary")
    subparsers.add_parser("today", help="Show today's usage")
    
    log = subparsers.add_parser("log", help="Log a generation")
    log.add_argument("--topic", required=True, help="Topic generated")
    log.add_argument("--count", type=int, default=1, help="Number of images")
    log.add_argument("--notes", default="", help="Optional notes")
    
    args = parser.parse_args()
    
    tracker = CostTracker()
    
    if args.command == "summary":
        tracker.print_summary()
    elif args.command == "today":
        today = tracker.get_today_summary()
        print(f"\nToday's usage ({today['date']}):")
        print(f"  Sessions: {today['sessions']}")
        print(f"  Images: {today['images']}")
        print(f"  Cost: ${today['cost_usd']:.2f} USD")
        print()
    elif args.command == "log":
        tracker.log_generation(args.topic, args.count, args.notes)
        print(f"Logged: {args.count} image(s) for '{args.topic}'")
        print(f"Cost: ${args.count * COST_PER_IMAGE:.4f}")
    else:
        tracker.print_summary()


if __name__ == "__main__":
    main()
