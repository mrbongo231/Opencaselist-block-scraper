#!/usr/bin/env python3
import os
import sys
import json
from pathlib import Path

# Provide a fallback token if not set
if not os.environ.get("CASELIST_TOKEN"):
    os.environ["CASELIST_TOKEN"] = "7197e95921e0982ac01651ae3045ff26"

import Caselistscrapper

def get_all_known_tournaments():
    """Scans local cache to discover active tournaments."""
    tournaments = set()
    cache_dir = Path("caselist_output/cache")
    if cache_dir.exists():
        for cache_file in cache_dir.glob("rounds_*.json"):
            try:
                rounds = json.loads(cache_file.read_text())
                for r in rounds:
                    t = r.get("tournament")
                    if t:
                        tournaments.add(t.strip())
            except Exception:
                pass
    return sorted(list(tournaments), key=lambda x: str(x).lower())

def run_interactive_menu():
    print("\n" + "="*60)
    print("  Topic/Tournament Interactive Search Menu")
    print("="*60)
    
    tournaments = get_all_known_tournaments()
    if not tournaments:
        print("[!] No tournaments found. Have you fetched any rounds yet?")
        sys.exit(1)
        
    print(f"\n[i] Discovered {len(tournaments)} unique tournaments in the database.")
    
    selected_tournaments = set()
    
    while True:
        print(f"\n[Status] You have selected {len(selected_tournaments)} tournaments.")
        if selected_tournaments:
            print("Selected: " + ", ".join(list(selected_tournaments)[:5]) + ("..." if len(selected_tournaments)>5 else ""))
            
        print("\nEnter a keyword to search for your desired tournaments (e.g., 'TFA', 'TOC', 'Harvard', 'March')")
        print("Or type 'done' to begin downloading evidence for your selections.")
        
        q = input("\nSearch keyword: ").strip().lower()
        if q == 'done':
            break
        if not q:
            continue
            
        matches = [t for t in tournaments if q in t.lower()]
        if not matches:
            print(f"[!] No tournaments found matching '{q}'. Try something else.")
            continue
            
        print(f"\nFound {len(matches)} matching tournaments:")
        for idx, m in enumerate(matches, 1):
            print(f"  {idx}. {m}")
            
        ans = input(f"\nAdd all {len(matches)} tournaments to your selection? (y/n/cancel): ").strip().lower()
        if ans == 'y':
            selected_tournaments.update(matches)
            print(f"[✓] Added {len(matches)} tournaments to your selection.")
            
    if not selected_tournaments:
        print("[!] No valid tournaments selected. Exiting.")
        sys.exit(0)
        
    print("\n[✓] You finalized the following tournaments:")
    for t in selected_tournaments:
        print(f"  - {t}")
        
    return list(selected_tournaments)

if __name__ == "__main__":
    selected_tournaments = run_interactive_menu()
    
    # Monkey-patch the topic matching function to ONLY allow selected tournaments
    def strict_tournament_match(rnd):
        t = rnd.get("tournament")
        if not t:
            return False
        return t.strip() in selected_tournaments
        
    Caselistscrapper._matches_topic = strict_tournament_match
    
    # Bypass original prompt and force the scraper into "topic" mode
    Caselistscrapper.prompt_for_target_mode = lambda: ("topic", {})
    Caselistscrapper.prompt_optional_topic_filter = lambda: []
    
    # Provide a dummy keyword so the "topic" mode doesn't abort early
    Caselistscrapper.TOPIC_KEYWORDS = ["Bypass_Keyword_Check"]
    
    print("\n[→] Starting comprehensive caselist scan for your selected tournaments...\n")
    Caselistscrapper.main()
