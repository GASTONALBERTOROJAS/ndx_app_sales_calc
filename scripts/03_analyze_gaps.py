"""Analyze rows without Cost_EUR or Cost_USD.

Interactive tool to examine gap rows by component.
"""

import json
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

# Convert Cost columns to numeric
df = df.copy()
df["Cost_EUR"] = pd.to_numeric(df["Cost_EUR"], errors="coerce")
df["Cost_USD"] = pd.to_numeric(df["Cost_USD"], errors="coerce")

# Filas sin ningún costo
missing = df[(df["Cost_EUR"].isna()) & (df["Cost_USD"].isna())].copy()

print(f"\n{'='*80}")
print(f"ANALISIS: Filas sin Cost_EUR ni Cost_USD ({len(missing)} filas)")
print(f"{'='*80}\n")

if len(missing) == 0:
    print("✓ Sin filas problemáticas.\n")
    print(f"{'='*80}\n")
else:
    print(f"Distribución por componente:\n")
    comp_counts = missing.groupby("Component").size().sort_values(ascending=False)
    comp_list = comp_counts.index.tolist()
    for idx, (comp, count) in enumerate(comp_counts.items(), 1):
        print(f"  {idx:2d}. {comp:<50} {count:>6,} filas")

    print(f"\nDistribución por región:\n")
    reg_counts = missing.groupby("Region").size().sort_values(ascending=False)
    for reg, count in reg_counts.items():
        print(f"  {reg:<20} {count:>6,} filas")

    print(f"\nDistribución por año:\n")
    yr_counts = missing.groupby("Year_Production").size().sort_values(ascending=False)
    for yr, count in yr_counts.items():
        print(f"  {int(yr):<20} {count:>6,} filas")

    print(f"\n{'='*80}\n")

    # Interactive mode: allow user to select component
    while True:
        print("MODO INTERACTIVO: Revisar detalles de filas sin costo\n")
        choice = input("Ingrese número de componente (1-{}) o 'q' para salir: ".format(len(comp_list))).strip()

        if choice.lower() == "q":
            print("\nSaliendo...\n")
            break

        try:
            idx = int(choice) - 1
            if 0 <= idx < len(comp_list):
                selected_comp = comp_list[idx]
                comp_missing = missing[missing["Component"] == selected_comp].copy()

                print(f"\n{'='*80}")
                print(f"Componente: {selected_comp}")
                print(f"Total filas sin costo: {len(comp_missing)}")
                print(f"{'='*80}\n")

                # Display in table format
                display_cols = ["Component", "Key", "Brand", "Year_Production", "Region", "Cost_EUR", "Cost_USD"]
                comp_missing_display = comp_missing[display_cols].sort_values(
                    by=["Year_Production", "Region", "Key"]
                ).reset_index(drop=True)

                # Pagination: show max 50 rows per chunk
                rows_per_page = 50
                total_rows = len(comp_missing_display)
                pages = (total_rows + rows_per_page - 1) // rows_per_page

                pd.set_option("display.max_rows", None)
                pd.set_option("display.max_columns", None)
                pd.set_option("display.width", None)
                pd.set_option("display.max_colwidth", None)

                for page in range(pages):
                    start_idx = page * rows_per_page
                    end_idx = min(start_idx + rows_per_page, total_rows)
                    chunk = comp_missing_display.iloc[start_idx:end_idx]

                    print(f"[Página {page + 1}/{pages}]")
                    print(chunk.to_string(index=False))

                    if page < pages - 1:
                        input("\nPresione ENTER para ver más filas...")
                        print()

                print(f"\n{'='*80}\n")

            else:
                print("❌ Opción inválida. Intente de nuevo.\n")
        except ValueError:
            print("❌ Entrada inválida. Ingrese un número o 'q'.\n")

print(f"{'='*80}\n")
