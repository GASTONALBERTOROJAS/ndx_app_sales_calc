"""Verify extracted output values against source Excel."""

import json
import pandas as pd
from pathlib import Path
from datetime import datetime

PROJECT_ROOT = Path(__file__).parent.parent
CONFIG_FILE = PROJECT_ROOT / "config.json"

with open(CONFIG_FILE, "r") as f:
    config = json.load(f)

today = datetime.now()
date_str = today.strftime("%d%m%y")
output_filename = f"SalesCalc_{date_str}.xlsx"
OUTPUT = PROJECT_ROOT / "output" / output_filename

df = pd.read_excel(OUTPUT)

tower = "Tower N117/3000 Controlled IEC2a TS76 TiT 50Hz NCV"

# (component, year, region, expected_eur, expected_usd)
checks = [
    ("Tower Shell", 2026, "Europe", 255993.398875, None),
    ("Tower Shell", 2026, "Asia", 45753.257325, 131977.1933106),
    ("Tower Shell", 2027, "Europe", 263120.80772125, None),
    ("Tower Shell", 2028, "Europe", 276276.8481073125, None),
    ("Tower Shell", 2026, "Germany", 290195.32807609043, None),
    ("Tower Internals", 2026, "Europe", 151000, None),
    ("Tower Bolts Set", 2026, "Europe", 3637.636325503345, None),
    ("Steel Tower Quality Inspectors", 2026, "Europe", 1950, None),
    ("Steel Tower Quality Inspectors", 2026, "US", None, 5519.999999999999),
    ("Option Coating c4/c5", 2026, "Europe", 5000, None),
]

print("=== VERIFICATION: Source vs Output ===")
print(f"Output file: {OUTPUT}")
print(f"Tower: {tower}\n")

all_ok = True
for comp, year, region, expected_eur, expected_usd in checks:
    mask = (
        (df["Key"] == tower)
        & (df["Component"] == comp)
        & (df["Year_Production"] == year)
        & (df["Region"] == region)
    )
    rows = df[mask]
    if len(rows) == 0:
        print(f"FAIL  {comp} / {year} / {region}: NOT FOUND in output")
        all_ok = False
        continue

    r = rows.iloc[0]
    eur = r["Cost_EUR"]
    usd = r["Cost_USD"]

    eur_ok = True
    usd_ok = True

    if expected_eur is not None:
        eur_ok = abs(float(eur) - expected_eur) < 0.01 if pd.notna(eur) else False

    if expected_usd is not None:
        usd_ok = abs(float(usd) - expected_usd) < 0.01 if pd.notna(usd) else False

    status = "OK" if (eur_ok and usd_ok) else "FAIL"
    if status == "FAIL":
        all_ok = False

    print(f"{status}  {comp} / {year} / {region}")
    print(f"       EUR: source={expected_eur} -> output={eur}")
    if expected_usd is not None:
        print(f"       USD: source={expected_usd} -> output={usd}")

result = "ALL CHECKS PASSED" if all_ok else "SOME CHECKS FAILED"
print(f"\n{result}")
