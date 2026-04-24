"""
Analyze rows without costs (CLI wrapper).

Calls the core gap analyzer module.
"""

import sys
from pathlib import Path
from datetime import datetime

from core.gap_analyzer import load_output, get_gap_rows, print_summary

PROJECT_ROOT = Path(__file__).parent.parent

today = datetime.now()
date_str = today.strftime("%d%m%y")
output_filename = f"SalesCalc_{date_str}.xlsx"
OUTPUT = PROJECT_ROOT / "output" / output_filename

if __name__ == "__main__":
    try:
        df = load_output(OUTPUT)
        gap_df = get_gap_rows(df)
        print_summary(gap_df, len(df))
        sys.exit(0)
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
