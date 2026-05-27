"""
Sales Calculation - Tower Component Cost Extractor (module version)

Extracts and normalizes tower component cost data from Excel spreadsheets.
Supports progress callbacks for UI integration.
"""

import re
import json
import logging
import warnings
import openpyxl
import pandas as pd
from pathlib import Path
from datetime import datetime

from .paths import ensure_output_dir

warnings.filterwarnings("ignore", category=FutureWarning, message=".*ChainedAssignmentError.*")

# --- Constants ---------------------------------------------------------------

DEFAULT_CANONICAL_REGIONS = [
    "Europe", "Germany", "Turkey", "Turkey_DOM",
    "Asia", "China", "Poland", "Italy", "Greece", "CAN", "US"
]

DEFAULT_SHEET_TS = "TS SC v26.2"
DEFAULT_SHEET_TCS = "TCS SC v26.2"

REGION_MAP = {
    "Europe": "Europe",
    "Germany": "Germany",
    "Turkey": "Turkey",
    "Turkey_DOM": "Turkey_DOM",
    "Asia": "Asia",
    "Asia EUR": "Asia",
    "Asia USD": "Asia",
    "China": "China",
    "China EUR": "China",
    "China USD": "China",
    "Poland": "Poland",
    "Italy": "Italy",
    "Greece": "Greece",
    "CAN": "CAN",
    "US": "US",
    "Arcosa US": "US",
    "Arcosa(US)": "US",
    "CS Wind US": "US",
    "CS Wind(US)": "US",
}

COMPONENTS_WITH_YEAR = [
    "tower shell",
    "Tower internals",
    "Anchor cage",
    "Tower bolts set",
]

COMPONENTS_REGION_ONLY = [
    "Steel Tower Quality Inspectors",
]

COMPONENTS_GLOBAL = [
    "Option Coating c4/c5",
    "Option Anchor Cage: German requirements for NAT AC",
    "Option hybrid tower: MB Monthly Cost Indexation",
    "Option hybrid tower: white coating of concrete part",
    "Option hybrid tower: no red stripe on concrete part",
]

COMPONENT_NAME_MAP = {
    "tower shell": "Tower Shell",
    "Tower internals": "Tower Internals",
    "Tower bolts set": "Tower Bolts Set",
}

TCS_MB_COMPONENTS = [
    (5, "Tower Shell"),
    (6, "Tower Internals"),
    (7, "Foundations"),
    (8, "Concrete Tower Keystones + Internals"),
    (9, "Concrete Tower Logistics"),
    (10, "Concrete Tower C&I"),
    (11, "Tower Bolts Set"),
    (13, "Option Coating c4/c5"),
    (14, "Steel Tower Quality Inspectors"),
    (15, "Option hybrid tower: white coating of concrete part"),
    (16, "Option hybrid tower: no red stripe on concrete part"),
]

# --- Logging -----------------------------------------------------------------

logging.basicConfig(
    level=logging.WARNING,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

# --- Helpers -----------------------------------------------------------------

def parse_row3_value(cell_value: str):
    """
    Parse a concatenated row-3 string like:
      'tower shell2026Cost_Currency1Europe'
    Returns (component, year, currency_type, region) or None.

    currency_type is 1 or 2 (from Cost_Currency1 / Cost_Currency2).
    year and region may be None for global/region-only components.
    """
    if not cell_value or not isinstance(cell_value, str):
        return None

    # Try to extract currency type
    m_curr = re.search(r"Cost_Currency([12])", cell_value)
    if not m_curr:
        return None
    currency_type = int(m_curr.group(1))

    # Part before Cost_Currency contains component + optional year
    before_cost = cell_value[: m_curr.start()]
    # Part after Cost_CurrencyX is the region
    after_cost = cell_value[m_curr.end():]
    region = after_cost.strip() if after_cost.strip() else None

    # Try to extract year from the end of before_cost
    m_year = re.search(r"(2025|2026|2027|2028)$", before_cost)
    if m_year:
        year = m_year.group(1)
        component = before_cost[: m_year.start()]
    else:
        year = None
        component = before_cost

    return component.strip(), year, currency_type, region


def extract_region_from_row7(cell_value: str):
    """
    Row 7 values like 'tower shell\\nEurope' — extract the region part.
    """
    if not cell_value or not isinstance(cell_value, str):
        return None
    parts = cell_value.split("\n")
    if len(parts) >= 2:
        return parts[-1].strip()
    return None


def normalize_region(raw_region: str) -> str | None:
    """
    Map a raw region string to one of the 11 canonical regions.
    Returns None if the region is not recognized.
    """
    if not raw_region:
        return None
    
    # Try case-insensitive lookup first
    raw_clean = raw_region.strip().lower()
    for key, val in REGION_MAP.items():
        if key.lower() == raw_clean:
            return val
            
    mapped = REGION_MAP.get(raw_region)
    if mapped is None:
        log.warning("Unknown region '%s' — skipping.", raw_region)
    return mapped


def _print_summary(df: pd.DataFrame, output_path: Path, log_callback=None) -> None:
    """Print extraction summary: years, regions, components, data quality."""
    def _log(msg: str):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    sep = "=" * 60
    total = len(df)
    ncols = len(df.columns)

    years = sorted(df["Year_Production"].dropna().unique())
    regions = sorted(df["Region"].dropna().unique())

    comp_stats = (
        df.groupby("Component")
        .agg(
            Filas=("Component", "count"),
            Con_C1=("Cost_Currency1", lambda x: x.notna().sum()),
            Con_C2=("Cost_Currency2", lambda x: x.notna().sum()),
        )
        .reset_index()
        .sort_values("Filas", ascending=False)
    )

    con_c1 = df["Cost_Currency1"].notna().sum()
    con_c2 = df["Cost_Currency2"].notna().sum()
    con_ambas = (df["Cost_Currency1"].notna() & df["Cost_Currency2"].notna()).sum()

    _log(f"\n{sep}")
    _log("EXTRACCION COMPLETADA")
    _log(sep)

    _log(f"\n  Archivo guardado:")
    _log(f"    {output_path}")
    _log(f"    {total:,} filas | {ncols} columnas")

    _log(f"\n  Years:")
    for y in years:
        _log(f"    - {int(y)}")

    _log(f"\n  Regiones ({len(regions)}):")
    for r in regions:
        _log(f"    - {r}")

    _log(f"\n  Componentes encontrados ({len(comp_stats)} unicos):\n")
    hdr = f"    {'Componente':<50} {'Filas':>7}"
    _log(hdr)
    _log("    " + "-" * (len(hdr) - 4))
    for _, row in comp_stats.iterrows():
        _log(f"    {row['Component']:<50} {int(row['Filas']):>7,}")

    _log(f"\n  Calidad de datos:")
    _log(f"    Filas con Cost_Currency1: {con_c1:>7,}  ({con_c1/total*100:.1f}%)")
    _log(f"    Filas con Cost_Currency2: {con_c2:>7,}  ({con_c2/total*100:.1f}%)")
    _log(f"    Filas con ambas:          {con_ambas:>7,}")

    _log(f"\n{sep}\n")


# --- Main extraction ---------------------------------------------------------

def run_extraction(
    input_path: Path | str,
    output_path: Path | str,
    target_years: list[int],
    sheet_ts: str = DEFAULT_SHEET_TS,
    sheet_tcs: str = DEFAULT_SHEET_TCS,
    canonical_regions: list[str] = None,
    progress_callback=None,
    log_callback=None,
) -> dict:
    """
    Extract and normalize tower component cost data.

    Args:
        input_path: Path to source Excel file
        output_path: Path to output Excel file
        target_years: List of years to process [2026, 2027, 2028, ...]
        sheet_ts: Name of TS SC sheet (default: "TS SC v26.1")
        sheet_tcs: Name of TCS_MB SC sheet (default: "TCS_MB SC v26.1")
        canonical_regions: List of canonical region names
        progress_callback: fn(pct: int, msg: str) for UI progress updates
        log_callback: fn(msg: str) for logging to UI

    Returns:
        dict with keys: total_rows, years, regions, components, output_path
    """
    if canonical_regions is None:
        canonical_regions = DEFAULT_CANONICAL_REGIONS

    input_path = Path(input_path)
    output_path = Path(output_path)
    ensure_output_dir(output_path)

    # Convert target_years to strings for comparison
    TARGET_YEARS = [str(y) for y in target_years]

    def _progress(pct: int, msg: str):
        log.info(msg)
        if progress_callback:
            progress_callback(pct, msg)
        if log_callback:
            log_callback(f"[{pct}%] {msg}")

    _progress(5, f"Abriendo workbook: {input_path}")
    wb = openpyxl.load_workbook(str(input_path), data_only=True, read_only=True)
    ws = wb[sheet_ts]
    _progress(10, f"Hoja '{sheet_ts}' cargada")

    # ---- Step 1: Build column map from rows 1-3 and 7 ----------------------
    _progress(12, "Leyendo encabezados...")
    row_data = {1: {}, 2: {}, 3: {}, 7: {}}
    for row in ws.iter_rows(min_row=1, max_row=7):
        for c in row:
            try:
                col = c.column
                r = c.row
                if r in row_data and c.value is not None:
                    row_data[r][col] = c.value
            except AttributeError:
                continue

    _progress(15, "Mapeando columnas...")

    # Parse row 3 to identify target columns
    all_target_components = (
        [c.lower() for c in COMPONENTS_WITH_YEAR]
        + [c.lower() for c in COMPONENTS_REGION_ONLY]
        + [c.lower() for c in COMPONENTS_GLOBAL]
    )

    col_map = {}
    for col, val in row_data[3].items():
        parsed = parse_row3_value(str(val))
        if parsed is None:
            continue
        component, year, currency_type, region = parsed

        if component.lower() not in all_target_components:
            continue

        if component.lower() in [c.lower() for c in COMPONENTS_WITH_YEAR]:
            if year not in TARGET_YEARS:
                continue

        region_r7 = extract_region_from_row7(row_data[7].get(col))
        raw_region = region_r7 or region
        final_region = normalize_region(raw_region)

        if raw_region and final_region is None:
            continue

        col_map[col] = {
            "component": component,
            "year": year,
            "currency_type": currency_type,
            "region": final_region,
        }

    _progress(30, f"Identificadas {len(col_map)} columnas objetivo")

    # ---- Step 2: Read data rows (row 8+) ------------------------------------
    _progress(35, "Leyendo datos de TS SC...")
    records = []
    row_count = 0

    for row in ws.iter_rows(min_row=8, max_row=1000):
        row_vals = {}
        for c in row:
            try:
                row_vals[c.column] = c.value
            except AttributeError:
                continue

        key = row_vals.get(4)
        if not key:
            continue

        brand = row_vals.get(5)
        row_count += 1

        for col, meta in col_map.items():
            value = row_vals.get(col)

            if meta["year"] is None:
                years_to_emit = TARGET_YEARS
            else:
                years_to_emit = [meta["year"]]

            if meta["region"] is None:
                regions_to_emit = canonical_regions
            else:
                regions_to_emit = [meta["region"]]

            for yr in years_to_emit:
                for rgn in regions_to_emit:
                    records.append({
                        "Component_Category": "Tower",
                        "Component": meta["component"],
                        "Key": key,
                        "Brand": brand,
                        "Year_Production": yr,
                        "Region": rgn,
                        "currency_type": meta["currency_type"],
                        "value": value,
                    })

    _progress(55, f"Extraídos {len(records)} registros de {row_count} torres (TS SC)")
    wb.close()

    # ---- Step 2b: Extract from TCS SC (dynamic header parser for v26.2+) ----------------
    _progress(60, f"Abriendo workbook para {sheet_tcs}...")
    wb2 = openpyxl.load_workbook(str(input_path), data_only=True, read_only=True)
    ws2 = wb2[sheet_tcs]

    tcs_records = []
    tcs_row_count = 0

    _progress(65, "Analizando estructura dinámica de columnas en TCS SC...")
    
    # Read rows 1 to 5 to understand headers
    header_rows = {1: {}, 2: {}, 3: {}, 4: {}, 5: {}}
    for row in ws2.iter_rows(min_row=1, max_row=5):
        for c in row:
            try:
                if c.value is not None:
                    header_rows[c.row][c.column] = c.value
            except AttributeError:
                continue

    # 1. Identify key columns (Key, Brand) in Row 5 (or fallbacks)
    key_col = None
    brand_col = None
    for col, val in header_rows[5].items():
        val_str = str(val).strip().lower()
        if val_str == "key":
            key_col = col
        elif val_str == "brand":
            brand_col = col

    if not key_col:
        # Fallback to column B (2) if not found explicitly
        key_col = 2
    if not brand_col:
        # Fallback to column C (3) if not found explicitly
        brand_col = 3

    # 2. Parse the Region + Year block columns from Row 3 (or Row 4) and Row 5
    # Standard components we expect to extract
    COMPONENT_PATTERNS = {
        "tower shell": "Tower Shell",
        "internals": "Tower Internals",
        "foundations": "Foundations",
        "keystones": "Concrete Tower Keystones + Internals",
        "logistics": "Concrete Tower Logistics",
        "c&i": "Concrete Tower C&I",
        "ac": "Anchor cage",
        "bolts": "Tower Bolts Set"
    }

    tcs_col_map = {}

    # Identify blocks from Row 3
    # A block header is like "EUROPE 2027", "GERMANY 2027", "TURKEY  2027", "POLAND 2027", etc.
    block_starts = [] # list of (col_idx, region, year)
    for col, val in header_rows[3].items():
        if not val or not isinstance(val, str):
            continue
        val_clean = val.strip()
        m = re.match(r"^([A-Za-z\s]+?)\s*(\d{4})$", val_clean)
        if m:
            raw_region = m.group(1).strip()
            norm_region = normalize_region(raw_region)
            year = int(m.group(2))
            if norm_region:
                block_starts.append((col, norm_region, year))

    # Sort blocks by column index
    block_starts.sort(key=lambda x: x[0])

    # Assign columns to blocks and map components
    for i, (start_col, region, year) in enumerate(block_starts):
        end_col = block_starts[i+1][0] if i + 1 < len(block_starts) else start_col + 20
        
        for col in range(start_col, end_col):
            cell_val = header_rows[5].get(col)
            if not cell_val:
                continue
            cell_str = str(cell_val).strip()
            
            for pattern, comp_name in COMPONENT_PATTERNS.items():
                if pattern in cell_str.lower():
                    tcs_col_map[col] = {
                        "type": "block",
                        "component": comp_name,
                        "region": region,
                        "year": year
                    }
                    break

    # 3. Parse options/standalone columns from Row 2
    for col, val in header_rows[2].items():
        if not val or not isinstance(val, str):
            continue
        val_clean = val.strip()
        
        if "option coating" in val_clean.lower():
            tcs_col_map[col] = {
                "type": "global",
                "component": "Option Coating c4/c5"
            }
        elif "white coating of concrete" in val_clean.lower():
            tcs_col_map[col] = {
                "type": "global",
                "component": "Option hybrid tower: white coating of concrete part"
            }
        elif "no red stripe on concrete" in val_clean.lower():
            tcs_col_map[col] = {
                "type": "global",
                "component": "Option hybrid tower: no red stripe on concrete part"
            }
        elif "steel tower quality inspectors" in val_clean.lower():
            r3_val = header_rows[3].get(col)
            region_name = None
            if r3_val:
                r3_str = str(r3_val)
                m_rgn = re.search(r"Region\s+([A-Za-z\s]+)", r3_str, re.IGNORECASE)
                if m_rgn:
                    region_name = normalize_region(m_rgn.group(1).strip())
                else:
                    region_name = normalize_region(r3_str.strip())
            
            if region_name:
                tcs_col_map[col] = {
                    "type": "region_option",
                    "component": "Steel Tower Quality Inspectors",
                    "region": region_name
                }

    log.info("Mapeadas %d columnas en TCS SC", len(tcs_col_map))
    _progress(70, f"Mapeadas {len(tcs_col_map)} columnas en TCS SC. Procesando filas...")

    # 4. Read data rows (Row 6 onwards)
    TARGET_YEARS = [int(y) for y in target_years]

    for row in ws2.iter_rows(min_row=6, max_row=1000):
        row_vals = {}
        for c in row:
            try:
                row_vals[c.column] = c.value
            except AttributeError:
                continue

        key = row_vals.get(key_col)
        if not key:
            continue

        brand = row_vals.get(brand_col) or "Nx"
        tcs_row_count += 1

        for col, meta in tcs_col_map.items():
            value = row_vals.get(col)
            if value is not None:
                if isinstance(value, str) and value.strip().lower() in ['n.a.', 'na', '']:
                    value = None

            if value is None:
                continue

            if meta["type"] == "block":
                if meta["year"] in TARGET_YEARS:
                    tcs_records.append({
                        "Component_Category": "Tower",
                        "Component": meta["component"],
                        "Key": key,
                        "Brand": brand,
                        "Year_Production": meta["year"],
                        "Region": meta["region"],
                        "currency_type": 1,
                        "value": value,
                    })

            elif meta["type"] == "global":
                for yr in TARGET_YEARS:
                    for rgn in canonical_regions:
                        tcs_records.append({
                            "Component_Category": "Tower",
                            "Component": meta["component"],
                            "Key": key,
                            "Brand": brand,
                            "Year_Production": yr,
                            "Region": rgn,
                            "currency_type": 1,
                            "value": value,
                        })

            elif meta["type"] == "region_option":
                for yr in TARGET_YEARS:
                    tcs_records.append({
                        "Component_Category": "Tower",
                        "Component": meta["component"],
                        "Key": key,
                        "Brand": brand,
                        "Year_Production": yr,
                        "Region": meta["region"],
                        "currency_type": 1,
                        "value": value,
                    })

    _progress(75, f"Extraídos {len(tcs_records)} registros de {tcs_row_count} torres (TCS SC)")
    wb2.close()

    records.extend(tcs_records)

    # ---- Step 3: Pivot Currency1 / Currency2 into separate columns ----------
    _progress(80, "Pivotando monedas...")
    df = pd.DataFrame(records)

    df.loc[:, "Component"] = df["Component"].replace(COMPONENT_NAME_MAP)
    df.loc[:, "Year_Production"] = pd.to_numeric(df["Year_Production"], errors="coerce").astype("Int64")

    df_c1 = df[df["currency_type"] == 1].rename(columns={"value": "Cost_Currency1"})
    df_c2 = df[df["currency_type"] == 2].rename(columns={"value": "Cost_Currency2"})

    merge_keys = [
        "Component_Category", "Component", "Key", "Brand",
        "Year_Production", "Region",
    ]

    df_c1 = df_c1.drop(columns=["currency_type"])
    df_c2 = df_c2[merge_keys + ["Cost_Currency2"]]

    df_final = pd.merge(df_c1, df_c2, on=merge_keys, how="outer")

    output_cols = [
        "Component_Category", "Component", "Key", "Brand",
        "Year_Production", "Region", "Cost_Currency1", "Cost_Currency2",
    ]
    df_final = df_final[output_cols].sort_values(
        by=["Component", "Year_Production", "Region", "Key"]
    ).reset_index(drop=True)

    _progress(95, f"Escribiendo {len(df_final):,} filas a {output_path}...")
    df_final.to_excel(str(output_path), index=False, sheet_name="Tower_Extracted")

    _progress(100, "Completado")

    _print_summary(df_final, output_path, log_callback)

    return {
        "total_rows": len(df_final),
        "years": sorted(df_final["Year_Production"].dropna().unique()),
        "regions": sorted(df_final["Region"].dropna().unique()),
        "components": sorted(df_final["Component"].unique()),
        "output_path": str(output_path),
    }
