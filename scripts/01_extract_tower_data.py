"""
Sales Calculation - Tower Component Cost Extractor (CLI wrapper).

CLI entry point that calls the core extraction module.
"""

import json
import sys
from pathlib import Path
from datetime import datetime

from core.extractor import run_extraction

PROJECT_ROOT = Path(__file__).parent.parent
CONFIG_FILE = PROJECT_ROOT / "config.json"

with open(CONFIG_FILE, "r") as f:
    config = json.load(f)

INPUT_FILE = config["input_file_path"]
today = datetime.now()
date_str = today.strftime("%d%m%y")
output_filename = f"SalesCalc_{date_str}.xlsx"
OUTPUT_FILE = PROJECT_ROOT / "output" / output_filename

if __name__ == "__main__":
    try:
        result = run_extraction(
            input_path=INPUT_FILE,
            output_path=OUTPUT_FILE,
            target_years=config["target_years"],
            sheet_ts=config["source_sheets"]["ts_sc"],
            sheet_tcs=config["source_sheets"]["tcs_mb"],
            canonical_regions=config["canonical_regions"],
        )
        sys.exit(0)
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
