"""Path utilities compatible with PyInstaller."""

import sys
from pathlib import Path


def get_exe_dir() -> Path:
    """
    Get directory containing exe (PyInstaller) or project root (dev).

    In PyInstaller onedir mode: directory containing NDX_Sales_Calculation.exe
    In dev: two levels above core/paths.py (i.e., project root)
    """
    if getattr(sys, 'frozen', False):
        # Running as PyInstaller exe
        return Path(sys.executable).parent
    # Running in development: __file__ is src/core/paths.py
    return Path(__file__).parent.parent.parent

def get_bundled_resource(relative_path: str) -> Path:
    """
    Get path to bundled resources (e.g., inside PyInstaller bundle).

    In PyInstaller: extracts to sys._MEIPASS
    In dev: relative to project root
    """
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS) / relative_path
    return Path(__file__).parent.parent.parent / relative_path


def ensure_output_dir(output_path: Path) -> None:
    """Create parent directory of output file if needed."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
