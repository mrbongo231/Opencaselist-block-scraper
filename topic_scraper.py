#!/usr/bin/env python3
"""
Topic Scraper — Date-Window Method

Automatically finds and downloads all OpenCaselist prep for a specific
PF topic by filtering rounds based on their upload timestamp (created_at).

Since NSDA PF topics rotate on a strict monthly schedule, we can determine
which topic a file belongs to purely from when it was uploaded — no guessing
tournament names, no keyword scanning.

Usage:
    python3 topic_scraper.py              # Interactive menu
    python3 topic_scraper.py --topic 5    # Skip menu, pick March 2026
    python3 topic_scraper.py --topic 5 --cache-only   # Only use cached rounds
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

# ───────────────────────────────────────────────────────────────
#  NSDA 2026-2027 PF TOPIC CALENDAR
# ───────────────────────────────────────────────────────────────

TOPICS = [
    {
        "number": 1,
        "label": "Sep/Oct 2026",
        "resolution": "The European Union should establish a nuclear sharing agreement with France to create an independent deterrent capability.",
        "start": "2026-09-01",
        "end": "2026-10-31",  # Ends at month boundary — Nov topic starts Nov 1
    },
    {
        "number": 2,
        "label": "Nov/Dec 2026",
        "resolution": "The United States federal government should require technology companies to provide lawful access to encrypted communications.",
        "start": "2026-11-01",
        "end": "2026-12-31",  # Ends at month boundary — Jan topic starts Jan 1
    },
    {
        "number": 3,
        "label": "Jan 2027",
        "resolution": "The benefits of the African Continental Free Trade Area outweigh the harms.",
        "start": "2027-01-01",
        "end": "2027-01-31",  # Ends at month boundary — Feb topic starts Feb 1
    },
    {
        "number": 4,
        "label": "Feb 2027",
        "resolution": "The Federal Trade Commission should establish a federal regulatory framework for sports betting.",
        "start": "2027-02-01",
        "end": "2027-02-28",  # Ends at month boundary — Mar topic starts Mar 1
    },
    {
        "number": 5,
        "label": "Mar 2027",
        "resolution": "The United States federal government should ban corporate acquisition of single-family residences.",
        "start": "2027-03-08",  # Pushed ~1 week to dodge late Feb topic uploads
        "end": "2027-03-31",  # Ends at month boundary — Apr topic starts Apr 1
    },
    {
        "number": 6,
        "label": "Apr 2027",
        "resolution": "The United States should eliminate the President's authority to deploy military forces abroad without Congressional approval.",
        "start": "2027-04-01",
        "end": "2027-05-07",  # Small grace for late uploads; last topic of year
    },
    {
        "number": 7,
        "label": "All Season (Everything)",
        "resolution": "All topics and rounds for the entire 2026-2027 season.",
        "start": "2026-08-01",
        "end": "2027-07-31",
    },
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scrape all OpenCaselist prep for a specific PF topic using date-window filtering."
    )
    parser.add_argument(
        "--topic",
        type=int,
        choices=[t["number"] for t in TOPICS],
        help=f"Topic number (1-{len(TOPICS)}) to skip the interactive menu.",
    )
    parser.add_argument(
        "--cache-only",
        action="store_true",
        help="Only scan cached round data — do not hit the API for new rounds.",
    )
    parser.add_argument(
        "--skip-nav",
        action="store_true",
        help="Skip running google_docs_nav_at_headings.py on the output.",
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="List all tournaments in the date window so you can review/exclude before downloading.",
    )
    parser.add_argument(
        "--exclude-tournaments",
        nargs="*",
        default=[],
        help="Tournament name substrings to exclude (case-insensitive). E.g. --exclude-tournaments 'Stanford' 'FFL'",
    )
    return parser.parse_args()


def display_topic_menu() -> dict:
    print("\n" + "=" * 70)
    print("  NSDA 2026-2027 PF Topic Selector")
    print("=" * 70)

    for t in TOPICS:
        print(f"\n  [{t['number']}] {t['label']}")
        # Wrap resolution text for readability
        res = t["resolution"]
        print(f"      Resolved: {res}")
        print(f"      Date window: {t['start']} → {t['end']}")

    print("\n" + "-" * 70)
    while True:
        try:
            choice = int(input("\nSelect a topic (1-6): ").strip())
            for t in TOPICS:
                if t["number"] == choice:
                    return t
            print("[!] Invalid choice. Enter a number 1-6.")
        except (ValueError, EOFError):
            print("[!] Invalid input. Enter a number 1-6.")


def _scan_cache_rounds(topic: dict) -> list[dict]:
    """Scan all cached round JSON files and return rounds in the topic's date window."""
    import json

    start_dt = datetime.strptime(topic["start"], "%Y-%m-%d")
    end_dt = datetime.strptime(topic["end"], "%Y-%m-%d")
    cache_dir = Path("caselist_output/cache")
    matching = []

    for cache_file in cache_dir.glob("rounds_*.json"):
        try:
            rounds = json.loads(cache_file.read_text())
        except Exception:
            continue
        for r in rounds:
            raw_ts = r.get("created_at") or r.get("updated_at") or ""
            if not raw_ts:
                continue
            try:
                dt = datetime.strptime(raw_ts[:19], "%Y-%m-%d %H:%M:%S")
            except Exception:
                continue
            if start_dt <= dt <= end_dt:
                matching.append(r)

    return matching


def preview_tournaments(topic: dict) -> set[str]:
    """
    List all tournaments found in the date window.
    Returns a set of tournament name substrings the user wants to EXCLUDE.
    """
    from collections import Counter

    print(f"\n[→] Scanning cache for tournaments in {topic['start']} → {topic['end']}...\n")
    rounds = _scan_cache_rounds(topic)

    tourney_counts = Counter()
    for r in rounds:
        tourn = r.get("tournament", "(unknown)")
        tourney_counts[tourn] += 1

    if not tourney_counts:
        print("[!] No rounds found in this date window.")
        return set()

    # Display numbered list
    sorted_tourneys = tourney_counts.most_common()
    print(f"  Found {sum(tourney_counts.values())} rounds across {len(sorted_tourneys)} tournaments:\n")
    print(f"  {'#':>4}  {'Rounds':>6}  Tournament")
    print(f"  {'─'*4}  {'─'*6}  {'─'*50}")
    for i, (name, count) in enumerate(sorted_tourneys, 1):
        print(f"  {i:4d}  {count:6d}  {name}")

    print(f"\n" + "─" * 70)
    print("  Enter tournament numbers to EXCLUDE (comma-separated), or press Enter to keep all.")
    print("  Example: 1,5,12")
    print("─" * 70)

    try:
        raw = input("\n  Exclude: ").strip()
    except (EOFError, KeyboardInterrupt):
        raw = ""

    if not raw:
        print("\n[✓] Keeping all tournaments.")
        return set()

    exclude_names = set()
    for token in raw.split(","):
        token = token.strip()
        if token.isdigit():
            idx = int(token) - 1
            if 0 <= idx < len(sorted_tourneys):
                name = sorted_tourneys[idx][0]
                exclude_names.add(name)
                print(f"  [✗] Excluding: {name}")

    print(f"\n[✓] Excluded {len(exclude_names)} tournament(s).")
    return exclude_names


def run_scraper(topic: dict, cache_only: bool, excluded_tournaments: set[str] | None = None) -> Path | None:
    """
    Hook into Caselistscrapper using date-window filtering.
    Returns the path to the compiled DOCX, or None on failure.
    """
    excluded_tournaments = excluded_tournaments or set()

    # Set token before importing Caselistscrapper
    if not os.environ.get("CASELIST_TOKEN"):
        os.environ["CASELIST_TOKEN"] = "7197e95921e0982ac01651ae3045ff26"

    import Caselistscrapper

    start_dt = datetime.strptime(topic["start"], "%Y-%m-%d")
    end_dt = datetime.strptime(topic["end"], "%Y-%m-%d")

    # Build lowercase exclusion set for substring matching (CLI --exclude-tournaments)
    exclude_lower = {e.lower() for e in excluded_tournaments}
    # Build exact-match exclusion set (from --preview selection)
    exclude_exact = {e for e in excluded_tournaments}

    def date_window_match(rnd: dict) -> bool:
        """Return True if this round is in the date window AND not excluded."""
        raw_ts = rnd.get("created_at") or rnd.get("updated_at") or ""
        if not raw_ts:
            return False

        parsed = Caselistscrapper._parse_api_timestamp(raw_ts)
        if parsed is None:
            return False

        if not (start_dt <= parsed <= end_dt):
            return False

        # Check tournament exclusions
        if excluded_tournaments:
            tourn = (rnd.get("tournament") or "").strip()
            # Exact match (from preview)
            if tourn in exclude_exact:
                return False
            # Substring match (from CLI)
            tourn_low = tourn.lower()
            for excl in exclude_lower:
                if excl in tourn_low:
                    return False

        return True

    # Monkey-patch the topic matching function
    Caselistscrapper._matches_topic = date_window_match

    # Force "topic" mode so it scans all schools
    Caselistscrapper.prompt_for_target_mode = lambda: ("topic", {})
    Caselistscrapper.prompt_optional_topic_filter = lambda: []
    Caselistscrapper.TOPIC_KEYWORDS = ["DateWindowFilter"]

    # Set descriptive output name
    safe_label = topic["label"].replace("/", "-").replace(" ", "_")
    Caselistscrapper.OUTPUT_NAME = f"topic_{safe_label}"

    print(f"\n{'=' * 70}")
    print(f"  Topic:       {topic['label']}")
    print(f"  Resolution:  {topic['resolution']}")
    print(f"  Date window: {topic['start']} → {topic['end']}")
    if excluded_tournaments:
        print(f"  Excluding:   {len(excluded_tournaments)} tournament(s)")
    print(f"{'=' * 70}\n")

    if cache_only:
        print("[→] Cache-only mode: scanning local round data only (no API calls for rounds).\n")
        from collections import defaultdict
        import json

        cache_dir = Path("caselist_output/cache")

        def resolve_from_cache():
            results = []
            teams_dict = defaultdict(list)

            for cache_file in cache_dir.glob("rounds_*.json"):
                try:
                    rounds = json.loads(cache_file.read_text())
                except Exception:
                    continue

                for r in rounds:
                    if not date_window_match(r):
                        continue
                    path = r.get("opensource")
                    if not path:
                        continue
                    parts = [p for p in path.split("/") if p]
                    if len(parts) >= 3:
                        school = parts[1]
                        team = parts[2]
                        teams_dict[(school, team)].append(r)

            for (school, team), rounds in teams_dict.items():
                results.append((school, team, rounds))

            print(f"[✓] Found {len(results)} teams with {topic['label']} evidence in cache.")
            return results

        Caselistscrapper.resolve_targets = resolve_from_cache

    print(f"[→] Starting scan for {topic['label']} topic...\n")
    Caselistscrapper.main()

    output_path = Path("caselist_output") / f"topic_{safe_label}.docx"
    if output_path.exists():
        return output_path
    return None


def run_nav_headings(docx_path: Path) -> Path:
    """Run google_docs_nav_at_headings.py on the compiled output."""
    nav_script = Path("google_docs_nav_at_headings.py")
    if not nav_script.exists():
        print(f"[!] Navigation script not found: {nav_script}")
        return docx_path

    output_nav = docx_path.with_name(docx_path.stem + "_nav" + docx_path.suffix)

    print(f"\n[→] Applying Google Docs navigation headings...")
    result = subprocess.run(
        [
            sys.executable,
            str(nav_script),
            str(docx_path),
            "--include-contentions",
            "--remove-file-headers",
            "--at-contention-only",
            "--strip-non-target-outline",
            "-o",
            str(output_nav),
        ],
        capture_output=True,
        text=True,
    )

    if result.returncode == 0:
        print(result.stdout)
        print(f"[✓] Navigation-ready file: {output_nav}")
        return output_nav
    else:
        print(f"[!] Navigation script error: {result.stderr}")
        return docx_path


def main():
    args = parse_args()

    # Select topic
    if args.topic:
        topic = next(t for t in TOPICS if t["number"] == args.topic)
        print(f"\n[→] Auto-selected topic #{args.topic}: {topic['label']}")
    else:
        topic = display_topic_menu()

    print(f"\n[✓] Selected: {topic['label']}")
    print(f"    {topic['resolution']}\n")

    # Collect tournament exclusions
    excluded = set(args.exclude_tournaments) if args.exclude_tournaments else set()

    # Preview mode: show tournaments, let user pick exclusions interactively
    if args.preview:
        preview_excluded = preview_tournaments(topic)
        excluded |= preview_excluded

    # Run scraper
    output_path = run_scraper(topic, cache_only=args.cache_only, excluded_tournaments=excluded)

    if output_path is None:
        print("\n[!] No output file was generated. Check for errors above.")
        return

    # Apply navigation headings
    if not args.skip_nav:
        final_path = run_nav_headings(output_path)
    else:
        final_path = output_path

    print(f"\n{'=' * 70}")
    print(f"  Done!")
    print(f"  Topic:  {topic['label']} — {topic['resolution'][:60]}...")
    print(f"  Output: {final_path}")
    print(f"{'=' * 70}\n")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[!] Cancelled by user (Ctrl+C).")
