"""Comprehensive verification of extracted output.

Validates:
  1. Schema — columns exist, dtypes correct
  2. Row count — exactly 133,588 rows
  3. Null integrity — no nulls in Key, Brand, Component_Category, Component, Region, Year_Production
  4. Domain values — Years in {2026,2027,2028}, Regions in canonical 11, Component_Category always "Tower"
  5. Component coverage — exactly 14 unique components
  6. Duplicates — no duplicate rows on merge keys
  7. Cost sanity — all non-null costs > 0
  8. Spot checks — 10 hardcoded value assertions (tolerance 0.01)
  9. Gap summary — count rows with both costs null (expected: ~10,410, INFO only)
"""

import json
import sys
import warnings
import pandas as pd
from pathlib import Path
from datetime import datetime

warnings.filterwarnings("ignore", category=FutureWarning)

PROJECT_ROOT = Path(__file__).parent.parent
CONFIG_FILE = PROJECT_ROOT / "config.json"

with open(CONFIG_FILE, "r") as f:
    config = json.load(f)

today = datetime.now()
date_str = today.strftime("%d%m%y")
output_filename = f"SalesCalc_{date_str}.xlsx"
OUTPUT = PROJECT_ROOT / "output" / output_filename

df = pd.read_excel(OUTPUT)

# Convert Cost columns to numeric (must do before dtype check)
df = df.copy()
df["Cost_EUR"] = pd.to_numeric(df["Cost_EUR"], errors="coerce")
df["Cost_USD"] = pd.to_numeric(df["Cost_USD"], errors="coerce")

# Validation results: list of (section, status, message)
results = []

# --- 1. Schema validation ---
expected_cols = ["Component_Category", "Component", "Key", "Brand", "Year_Production", "Region", "Cost_EUR", "Cost_USD"]
cols_ok = all(col in df.columns for col in expected_cols)
if cols_ok:
    results.append(("SCHEMA", "OK", f"8 columnas | {len(df):,} filas"))
else:
    missing = [c for c in expected_cols if c not in df.columns]
    results.append(("SCHEMA", "FAIL", f"Columnas faltantes: {missing}"))

# --- 2. Dtypes ---
dtype_ok = True
dtype_msgs = []
if df["Year_Production"].dtype not in ["int64", "Int64"]:
    dtype_ok = False
    dtype_msgs.append(f"Year_Production={df['Year_Production'].dtype} (expected int64)")
if df["Cost_EUR"].dtype != "float64":
    dtype_ok = False
    dtype_msgs.append(f"Cost_EUR={df['Cost_EUR'].dtype} (expected float64)")
if df["Cost_USD"].dtype != "float64":
    dtype_ok = False
    dtype_msgs.append(f"Cost_USD={df['Cost_USD'].dtype} (expected float64)")

if dtype_ok:
    results.append(("DTYPES", "OK", "Year_Production=int64 | Cost_EUR=float64 | Cost_USD=float64"))
else:
    results.append(("DTYPES", "FAIL", " | ".join(dtype_msgs)))

# --- 3. Row count ---
expected_rows = 133588
row_count_ok = len(df) == expected_rows
if row_count_ok:
    results.append(("ROW_COUNT", "OK", f"{len(df):,} filas"))
else:
    results.append(("ROW_COUNT", "FAIL", f"Esperado={expected_rows:,} | Encontrado={len(df):,}"))

# --- 4. Null integrity (Key, Brand, Component_Category, Component, Region, Year_Production) ---
critical_cols = ["Key", "Brand", "Component_Category", "Component", "Region", "Year_Production"]
null_counts = {col: df[col].isna().sum() for col in critical_cols}
nulls_ok = all(count == 0 for count in null_counts.values())
if nulls_ok:
    results.append(("NULLS", "OK", "Sin nulos en columnas clave"))
else:
    null_details = [f"{col}={count}" for col, count in null_counts.items() if count > 0]
    results.append(("NULLS", "FAIL", " | ".join(null_details)))

# --- 5. Year domain ---
canonical_years = set(config["target_years"])
found_years = set(df["Year_Production"].dropna().unique())
years_ok = found_years == canonical_years
if years_ok:
    results.append(("YEARS", "OK", f"Años validos: {sorted(canonical_years)}"))
else:
    extra = found_years - canonical_years
    missing = canonical_years - found_years
    diff = []
    if extra:
        diff.append(f"Extra={extra}")
    if missing:
        diff.append(f"Faltantes={missing}")
    results.append(("YEARS", "FAIL", " | ".join(diff)))

# --- 6. Region domain ---
canonical_regions = set(config["canonical_regions"])
found_regions = set(df["Region"].dropna().unique())
regions_ok = found_regions == canonical_regions
if regions_ok:
    results.append(("REGIONS", "OK", f"{len(found_regions)}/11 regiones canonicas encontradas"))
else:
    extra = found_regions - canonical_regions
    missing = canonical_regions - found_regions
    diff = []
    if extra:
        diff.append(f"Extra={extra}")
    if missing:
        diff.append(f"Faltantes={missing}")
    results.append(("REGIONS", "FAIL", " | ".join(diff)))

# --- 7. Component_Category always "Tower" ---
comp_cat_ok = (df["Component_Category"] == "Tower").all()
if comp_cat_ok:
    results.append(("COMPONENT_CATEGORY", "OK", 'Siempre "Tower"'))
else:
    bad_count = (df["Component_Category"] != "Tower").sum()
    results.append(("COMPONENT_CATEGORY", "FAIL", f"{bad_count} filas con valor != Tower"))

# --- 8. Component coverage (14 expected) ---
expected_components = {
    "Tower Shell",
    "Tower Internals",
    "Anchor cage",
    "Tower Bolts Set",
    "Steel Tower Quality Inspectors",
    "Option Coating c4/c5",
    "Option Anchor Cage: German requirements for NAT AC",
    "Option hybrid tower: MB Monthly Cost Indexation",
    "Option hybrid tower: white coating of concrete part",
    "Option hybrid tower: no red stripe on concrete part",
    "Foundations",
    "Concrete Tower Logistics",
    "Concrete Tower Keystones + Internals",
    "Concrete Tower C&I",
}
found_components = set(df["Component"].unique())
components_ok = found_components == expected_components
if components_ok:
    results.append(("COMPONENTS", "OK", f"14/14 componentes esperados encontrados"))
else:
    extra = found_components - expected_components
    missing = expected_components - found_components
    diff = []
    if extra:
        diff.append(f"Extra={len(extra)}")
    if missing:
        diff.append(f"Faltantes={len(missing)}")
    results.append(("COMPONENTS", "FAIL", " | ".join(diff)))

# --- 9. No duplicates on merge keys ---
merge_keys = ["Component", "Key", "Brand", "Year_Production", "Region"]
dups = df.duplicated(subset=merge_keys, keep=False)
dup_count = dups.sum()
if dup_count == 0:
    results.append(("DUPLICATES", "OK", "Sin filas duplicadas"))
else:
    results.append(("DUPLICATES", "WARN", f"{dup_count} filas duplicadas (revisar datos fuente)"))

# --- 10. Cost sanity (all non-null > 0) ---
eur_non_null = df["Cost_EUR"].dropna()
usd_non_null = df["Cost_USD"].dropna()
eur_bad = (eur_non_null <= 0).sum() if len(eur_non_null) > 0 else 0
usd_bad = (usd_non_null <= 0).sum() if len(usd_non_null) > 0 else 0
cost_issues = eur_bad + usd_bad
if cost_issues == 0:
    results.append(("COST_SANITY", "OK", "Todos los costos > 0"))
else:
    problems = []
    if eur_bad > 0:
        problems.append(f"EUR<=0: {eur_bad}")
    if usd_bad > 0:
        problems.append(f"USD<=0: {usd_bad}")
    results.append(("COST_SANITY", "WARN", " | ".join(problems)))

# --- 11. Spot checks (10 hardcoded assertions) ---
tower = "Tower N117/3000 Controlled IEC2a TS76 TiT 50Hz NCV"
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

spot_checks_passed = 0
spot_checks_failed = 0
spot_check_details = []

for comp, year, region, expected_eur, expected_usd in checks:
    mask = (
        (df["Key"] == tower)
        & (df["Component"] == comp)
        & (df["Year_Production"] == year)
        & (df["Region"] == region)
    )
    rows = df[mask]
    if len(rows) == 0:
        spot_checks_failed += 1
        spot_check_details.append(f"✗ {comp}/{year}/{region}: NOT FOUND")
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

    if eur_ok and usd_ok:
        spot_checks_passed += 1
    else:
        spot_checks_failed += 1
        problem = []
        if not eur_ok:
            problem.append(f"EUR mismatch")
        if not usd_ok:
            problem.append(f"USD mismatch")
        spot_check_details.append(f"✗ {comp}/{year}/{region}: {', '.join(problem)}")

if spot_checks_failed == 0:
    results.append(("SPOT_CHECKS", "OK", f"{spot_checks_passed}/10 checks pasados"))
else:
    results.append(("SPOT_CHECKS", "FAIL", f"{spot_checks_passed} pasados, {spot_checks_failed} fallidos"))

# --- 12. Gap summary (INFO only, expected ~10,410) ---
gaps = df[(df["Cost_EUR"].isna()) & (df["Cost_USD"].isna())]
gap_count = len(gaps)
results.append(("GAPS", "INFO", f"{gap_count:,} filas sin costo (EUR y USD null)"))

# --- Print results ---
print("\n" + "=" * 60)
print("VERIFICATION REPORT")
print("=" * 60 + "\n")

# Collect pass/fail/warn counts
pass_count = sum(1 for _, status, _ in results if status == "OK")
fail_count = sum(1 for _, status, _ in results if status == "FAIL")
warn_count = sum(1 for _, status, _ in results if status == "WARN")

# Print each result
for section, status, msg in results:
    status_str = f"[{status}]".ljust(7)
    section_str = f"[{section}]".ljust(18)
    print(f"{section_str} {status_str} {msg}")

# Print spot check details if any failed
if spot_check_details:
    print()
    for detail in spot_check_details:
        print(f"  {detail}")

# Final result line
print("\n" + "-" * 60)
if fail_count == 0:
    status_summary = f"{pass_count}/11 checks OK"
    if warn_count > 0:
        status_summary += f" | {warn_count} advertencia(s)"
    final = f"RESULTADO FINAL: PASS  —  {status_summary} | {spot_checks_passed}/10 spot checks"
    print(final)
    exit_code = 0
else:
    final = f"RESULTADO FINAL: FAIL  —  {fail_count} error(es) detectado(s)"
    print(final)
    exit_code = 1

print("=" * 60 + "\n")

sys.exit(exit_code)
