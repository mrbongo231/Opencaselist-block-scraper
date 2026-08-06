import os
import sys
from datetime import datetime

# Ensure the token is set
os.environ["CASELIST_TOKEN"] = "7197e95921e0982ac01651ae3045ff26"

import Caselistscrapper

def march_is_recent(rnd, cutoff):
    created_dt = Caselistscrapper._parse_api_timestamp(rnd.get("created_at"))
    if not created_dt:
        return False
    # Only return true if it's March (and optionally we could restrict year to 2026 or 2025)
    return created_dt.month == 3

# Monkey patch
Caselistscrapper._is_recent = march_is_recent

# Mock interactive prompts
Caselistscrapper.prompt_for_target_mode = lambda: ("recent", {"DAYS_RECENT": 365})
Caselistscrapper.prompt_optional_topic_filter = lambda: []

if __name__ == "__main__":
    print("Running scraper for March topic...")
    Caselistscrapper.main()
