"""
Comprehensive verification of extracted output (CLI wrapper).

Calls the core verification module.
"""

import json
import sys
from pathlib import Path
from datetime import datetime

from core.verifier import run_verification

PROJECT_ROOT = Path(__file__).parent.parent
CONFIG_FILE = PROJECT_ROOT / "config.json"

with open(CONFIG_FILE, "r") as f:
    config = json.load(f)

today = datetime.now()
date_str = today.strftime("%d%m%y")
output_filename = f"SalesCalc_{date_str}.xlsx"
OUTPUT = PROJECT_ROOT / "output" / output_filename

if __name__ == "__main__":
    try:
        # For CLI, always use standard years [2026, 2027, 2028]
        result = run_verification(
            output_path=OUTPUT,
            selected_years=config["target_years"],
            canonical_regions=config["canonical_regions"],
        )
        sys.exit(result["exit_code"])
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
