"""Verify extracted output values against source Excel."""

import pandas as pd

OUTPUT = r"C:\Users\Gaston Alberto\Desktop\WORKSPACE\ai_developer\projects\tower_sales_extraction\tower_components_extracted.xlsx"

df = pd.read_excel(OUTPUT)

tower = "Tower N117/3000 Controlled IEC2a TS76 TiT 50Hz NCV"

# (component, year, region, expected_c1, expected_c2)
checks = [
    ("tower shell", 2026, "Europe", 255993.398875, None),
    ("tower shell", 2026, "Asia", 45753.257325, 131977.1933106),
    ("tower shell", 2027, "Europe", 263120.80772125, None),
    ("tower shell", 2028, "Europe", 276276.8481073125, None),
    ("tower shell", 2026, "Germany", 290195.32807609043, None),
    ("Tower internals", 2026, "Europe", 151000, None),
    ("Anchor cage", 2026, "Europe", 22339.23818, None),
    ("Tower bolts set", 2026, "Europe", 3637.636325503345, None),
    ("Steel Tower Quality Inspectors", 2026, "Europe", 1950, None),
    ("Steel Tower Quality Inspectors", 2026, "US", None, 5519.999999999999),
    ("Option Coating c4/c5", 2026, "Europe", 5000, None),
    ("Option hybrid tower: MB Monthly Cost Indexation", 2026, "Europe", "n.a.", None),
]

print("=== VERIFICATION: Source vs Output ===")
print(f"Tower: {tower}\n")

all_ok = True
for comp, year, region, expected_c1, expected_c2 in checks:
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
    c1 = r["Cost_Currency1"]
    c2 = r["Cost_Currency2"]

    c1_ok = True
    c2_ok = True

    if expected_c1 is not None:
        if isinstance(expected_c1, str):
            c1_ok = str(c1) == expected_c1
        else:
            c1_ok = abs(float(c1) - expected_c1) < 0.01 if pd.notna(c1) else False

    if expected_c2 is not None:
        c2_ok = abs(float(c2) - expected_c2) < 0.01 if pd.notna(c2) else False

    status = "OK" if (c1_ok and c2_ok) else "FAIL"
    if status == "FAIL":
        all_ok = False

    print(f"{status}  {comp} / {year} / {region}")
    print(f"       C1: source={expected_c1} -> output={c1}")
    if expected_c2 is not None:
        print(f"       C2: source={expected_c2} -> output={c2}")

result = "ALL CHECKS PASSED" if all_ok else "SOME CHECKS FAILED"
print(f"\n{result}")
