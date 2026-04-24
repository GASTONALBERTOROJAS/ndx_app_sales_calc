"""
Analyze rows without costs (module version - no input() interaction).
"""

import warnings
import pandas as pd
from pathlib import Path

warnings.filterwarnings("ignore", category=FutureWarning)


def load_output(output_path: Path | str) -> pd.DataFrame:
    """Load output Excel file and return DataFrame."""
    output_path = Path(output_path)
    return pd.read_excel(str(output_path))


def get_gap_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Filter rows where both Cost_Currency1 and Cost_Currency2 are null."""
    return df[(df["Cost_Currency1"].isna()) & (df["Cost_Currency2"].isna())].copy()


def get_component_summary(gap_df: pd.DataFrame) -> list[tuple[str, int]]:
    """
    Get summary of gap counts by component.

    Returns:
        list of (component_name, count) tuples, sorted by count descending
    """
    if len(gap_df) == 0:
        return []
    counts = gap_df.groupby("Component").size().sort_values(ascending=False)
    return list(zip(counts.index, counts.values))


def get_region_summary(gap_df: pd.DataFrame) -> list[tuple[str, int]]:
    """
    Get summary of gap counts by region.

    Returns:
        list of (region_name, count) tuples, sorted by count descending
    """
    if len(gap_df) == 0:
        return []
    counts = gap_df.groupby("Region").size().sort_values(ascending=False)
    return list(zip(counts.index, counts.values))


def get_year_summary(gap_df: pd.DataFrame) -> list[tuple[int, int]]:
    """
    Get summary of gap counts by year.

    Returns:
        list of (year, count) tuples, sorted by year ascending
    """
    if len(gap_df) == 0:
        return []
    counts = gap_df.groupby("Year_Production").size().sort_values(ascending=False)
    return list(zip(counts.index, counts.values))


def get_component_rows(
    gap_df: pd.DataFrame,
    component: str,
    max_rows: int = 500,
) -> pd.DataFrame:
    """
    Get rows for a specific component (limited to max_rows).

    Args:
        gap_df: DataFrame of gap rows (output of get_gap_rows)
        component: Component name to filter
        max_rows: Maximum number of rows to return

    Returns:
        DataFrame with selected columns, sorted, limited to max_rows
    """
    comp_rows = gap_df[gap_df["Component"] == component].copy()

    display_cols = [
        "Component", "Key", "Brand", "Year_Production", "Region",
        "Cost_Currency1", "Cost_Currency2"
    ]
    comp_rows = comp_rows[display_cols].sort_values(
        by=["Year_Production", "Region", "Key"]
    ).reset_index(drop=True)

    return comp_rows.head(max_rows)


def print_summary(gap_df: pd.DataFrame, total_rows: int, log_callback=None) -> None:
    """
    Print gap analysis summary (for CLI use).

    Args:
        gap_df: DataFrame of gap rows
        total_rows: Total rows in the output file
        log_callback: Optional fn(msg: str) for logging
    """
    def _log(msg: str):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    sep = "=" * 80
    gap_count = len(gap_df)

    _log(f"\n{sep}")
    _log(f"ANALISIS: Filas sin Cost_Currency1 ni Cost_Currency2 ({gap_count} filas)")
    _log(f"{sep}\n")

    if gap_count == 0:
        _log("✓ Sin filas problemáticas.\n")
        _log(f"{sep}\n")
        return

    comp_summary = get_component_summary(gap_df)
    region_summary = get_region_summary(gap_df)
    year_summary = get_year_summary(gap_df)

    _log("Distribución por componente:\n")
    for idx, (comp, count) in enumerate(comp_summary, 1):
        _log(f"  {idx:2d}. {comp:<50} {count:>6,} filas")

    _log("\nDistribución por región:\n")
    for reg, count in region_summary:
        _log(f"  {reg:<20} {count:>6,} filas")

    _log("\nDistribución por año:\n")
    for yr, count in year_summary:
        _log(f"  {int(yr):<20} {count:>6,} filas")

    _log(f"\n{sep}\n")
    _log(f"Total de filas sin costo: {gap_count}/{total_rows} ({gap_count/total_rows*100:.1f}%)\n")
