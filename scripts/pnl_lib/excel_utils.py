"""Excel helpers shared by SKU and project-level extractors."""

from __future__ import annotations

import os
import re

import pandas as pd
from openpyxl.utils import coordinate_to_tuple

ALLOWED_EXTENSIONS = (".xlsx", ".xls", ".xlsm")


def list_excel_files(directory: str) -> list[str]:
    """Return absolute paths to Excel files under directory (recursive)."""
    directory = os.path.abspath(directory)
    paths: list[str] = []
    for root, _dirs, files in os.walk(directory):
        for name in files:
            if name.endswith(ALLOWED_EXTENSIONS) and not name.startswith("~"):
                paths.append(os.path.join(root, name))
    return sorted(paths)


def basename_only(filepath: str) -> str:
    return os.path.basename(filepath)


def extract_project_name(filepath: str) -> str | None:
    filename = basename_only(filepath)
    matches = re.findall(r"-\s*([^-\n\r]+?)\s*-", filename)
    return matches[0].strip() if matches else None


def extract_year(filepath: str) -> str | None:
    components = re.split(r"[_\-\s\.\\/]", filepath)
    for comp in components:
        match = re.search(r"20\d{2}", comp)
        if match:
            return match.group(0)
    return None


def extract_gate_number(filepath: str) -> str | None:
    filename = basename_only(filepath)
    match = re.match(r"^\s*([^-\n\r]+?)\s*-", filename)
    if not match:
        return None
    g_match = re.search(r"G\d+", match.group(1).strip())
    return g_match.group(0) if g_match else None


def extract_excel_table_to_df(
    filepath: str,
    sheet_name: str | int,
    start_cell: str,
    end_cell: str,
    header_row: bool = False,
    engine: str = "openpyxl",
) -> pd.DataFrame:
    df_raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None, engine=engine)
    start_row, start_col = coordinate_to_tuple(start_cell)
    end_row, end_col = coordinate_to_tuple(end_cell)
    df_range = df_raw.iloc[start_row - 1 : end_row, start_col - 1 : end_col]
    if header_row:
        df_range.columns = df_range.iloc[0]
        df_range = df_range.iloc[1:].reset_index(drop=True)
    else:
        df_range = df_range.reset_index(drop=True)
    return df_range


def resolve_sheet_name(filepath: str, preferred: str, engine: str = "openpyxl") -> str:
    xf = pd.ExcelFile(filepath, engine=engine)
    names = list(xf.sheet_names)
    if preferred in names:
        return preferred
    lower_map = {str(n).strip().lower(): n for n in names}
    key = preferred.strip().lower()
    if key in lower_map:
        return lower_map[key]
    raise ValueError(f"Sheet {preferred!r} not found. Available: {names}")
