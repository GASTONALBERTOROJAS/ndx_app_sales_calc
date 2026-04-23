"""Analyze rows without Cost_EUR or Cost_USD."""

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

# Filas sin ningún costo
missing = df[(df["Cost_EUR"].isna()) & (df["Cost_USD"].isna())].copy()

print(f"\n{'='*70}")
print(f"ANALISIS: Filas sin Cost_EUR ni Cost_USD ({len(missing)} filas)")
print(f"{'='*70}\n")

if len(missing) == 0:
    print("✓ Sin filas problemáticas.\n")
else:
    print(f"Distribución por componente:\n")
    comp_counts = missing.groupby("Component").size().sort_values(ascending=False)
    for comp, count in comp_counts.items():
        print(f"  {comp:<50} {count:>6,} filas")

    print(f"\nDistribución por región:\n")
    reg_counts = missing.groupby("Region").size().sort_values(ascending=False)
    for reg, count in reg_counts.items():
        print(f"  {reg:<20} {count:>6,} filas")

    print(f"\nDistribución por año:\n")
    yr_counts = missing.groupby("Year_Production").size().sort_values(ascending=False)
    for yr, count in yr_counts.items():
        print(f"  {int(yr):<20} {count:>6,} filas")

    print(f"\nPrimeros 10 ejemplos:\n")
    sample = missing[["Component", "Key", "Year_Production", "Region"]].head(10)
    for idx, row in sample.iterrows():
        print(f"  {row['Component']:<45} | {row['Year_Production']} | {row['Region']}")

print(f"\n{'='*70}\n")
