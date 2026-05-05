#!/usr/bin/env python3
"""
CLI: walk a directory of P&L Excel files and extract Project Summary rows
into one DataFrame (same logic as PandL_mastersku_extract.ipynb).

Parameters (sheet name, cell ranges, feature blocks, file extensions) live in
config/pnl_extract.json by default. Use `scripts/extract_igate_key_metrics.py` for iGate Key Metrics.

Usage:
  python scripts/extract_pnl_directory.py /path/to/folder
  python scripts/extract_pnl_directory.py /path/to/folder -o out.csv
  python scripts/extract_pnl_directory.py /path/to/folder --config /path/to/custom.json
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from typing import Any

import pandas as pd
from openpyxl.utils import coordinate_to_tuple


def default_config_path() -> str:
    return os.path.abspath(
        os.path.join(os.path.dirname(__file__), "..", "config", "pnl_extract.json")
    )


def load_config(path: str | None) -> dict[str, Any]:
    config_path = os.path.abspath(os.path.expanduser(path or default_config_path()))
    if not os.path.isfile(config_path):
        raise FileNotFoundError(f"Config file not found: {config_path}")
    try:
        with open(config_path, encoding="utf-8") as f:
            cfg: dict[str, Any] = json.load(f)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON in config {config_path}: {e}") from e

    required_top = (
        "sheet_name",
        "list_price_table",
        "clean_list_price",
        "melt",
        "feature_tables",
        "walk",
    )
    missing = [k for k in required_top if k not in cfg]
    if missing:
        raise ValueError(f"Config missing keys: {missing}")

    for key in ("start_cell", "end_cell"):
        if key not in cfg["list_price_table"]:
            raise ValueError(f"list_price_table missing '{key}'")
    for key in ("column_name_substring", "exclude_rows_value_contains"):
        if key not in cfg["clean_list_price"]:
            raise ValueError(f"clean_list_price missing '{key}'")
    for key in ("id_var_column", "var_name", "value_name"):
        if key not in cfg["melt"]:
            raise ValueError(f"melt missing '{key}'")
    if not isinstance(cfg["feature_tables"], list) or not cfg["feature_tables"]:
        raise ValueError("feature_tables must be a non-empty list")
    for i, row in enumerate(cfg["feature_tables"]):
        for key in ("field_name", "start_cell", "end_cell"):
            if key not in row:
                raise ValueError(f"feature_tables[{i}] missing '{key}'")
    if "extensions" not in cfg["walk"]:
        raise ValueError("walk.extensions is required")

    return cfg


def resolve_sheet_name(
    filepath: str,
    preferred: str,
    engine: str,
    aliases: list[str] | None = None,
) -> str:
    """Match configured sheet name to workbook tabs (exact, then aliases, then case-insensitive)."""
    xf = pd.ExcelFile(filepath, engine=engine)
    names = list(xf.sheet_names)
    candidates: list[str] = [preferred] + list(aliases or [])
    for c in candidates:
        if c in names:
            return c
    lower_map: dict[str, str] = {}
    for n in names:
        key = str(n).strip().lower()
        if key not in lower_map:
            lower_map[key] = n
    for c in candidates:
        key = str(c).strip().lower()
        if key in lower_map:
            return lower_map[key]
    raise ValueError(
        f"No sheet matching any of {candidates!r}. Available sheets ({len(names)}): {names}"
    )


def extract_project_name(filepath: str) -> str | None:
    filename = filepath.split("/")[-1] if "/" in filepath else filepath
    matches = re.findall(r"-\s*([^-\n\r]+?)\s*-", filename)
    if matches:
        return matches[0].strip()
    return None


def extract_year(filepath: str) -> str | None:
    components = re.split(r"[_\-\s\.\\/]", filepath)
    for comp in components:
        match = re.search(r"20\d{2}", comp)
        if match:
            return match.group(0)
    return None


def extract_gate_number(filepath: str) -> str | None:
    filename = os.path.basename(filepath)
    match = re.match(r"^\s*([^-\n\r]+?)\s*-", filename)
    if not match:
        return None
    item = match.group(1).strip()
    g_match = re.search(r"G\d+", item)
    return g_match.group(0) if g_match else None


def extract_excel_table_to_df(
    filepath: str,
    sheet_name: str | int,
    start_cell: str,
    end_cell: str,
    header_row: bool,
    engine: str,
) -> pd.DataFrame:
    df_raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None, engine=engine)
    start_row, start_col = coordinate_to_tuple(start_cell)
    end_row, end_col = coordinate_to_tuple(end_cell)
    df_range = df_raw.iloc[start_row - 1 : end_row, start_col - 1 : end_col]
    if header_row:
        raw_headers = list(df_range.iloc[0])
        names: list[str] = []
        for i, c in enumerate(raw_headers):
            if c is None or (isinstance(c, float) and pd.isna(c)):
                names.append(f"_unnamed_{i}")
            else:
                names.append(str(c).strip())
        seen: dict[str, int] = {}
        unique: list[str] = []
        for n in names:
            if n in seen:
                seen[n] += 1
                unique.append(f"{n}__dup{seen[n]}")
            else:
                seen[n] = 0
                unique.append(n)
        df_range.columns = unique
        df_range = df_range.iloc[1:].reset_index(drop=True)
    else:
        df_range = df_range.reset_index(drop=True)
    return df_range


def clean_list_price_df(df: pd.DataFrame, cfg: dict[str, Any]) -> pd.DataFrame:
    c = cfg["clean_list_price"]
    substr = c["column_name_substring"]
    exclude = c["exclude_rows_value_contains"]
    id_target = cfg["melt"]["id_var_column"]

    list_price_col = df.columns[df.columns.astype(str).str.contains(substr, case=False, na=False)]
    matched_list_price: str | None = None
    if not list_price_col.empty:
        matched_list_price = list_price_col[0]
        df = df[~df[matched_list_price].astype(str).str.contains(exclude, case=False, na=False)]

    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~((df == 0) | (df.isna())).all(axis=0)]
    if df.empty or len(df.index) == 0:
        return df
    df = df.drop(df.index[0]).reset_index(drop=True)
    # Melt uses exact id_var_column name; header may be "List Price (AUD)" etc.
    if matched_list_price is not None and matched_list_price in df.columns and matched_list_price != id_target:
        df = df.rename(columns={matched_list_price: id_target})
    return df


def melt_list_price_df(df: pd.DataFrame, cfg: dict[str, Any]) -> pd.DataFrame:
    m = cfg["melt"]
    id_var = m["id_var_column"]
    var_name = m["var_name"]
    value_name = m["value_name"]
    sep = m.get("sku_market_separator", "_")

    df_melted = df.melt(id_vars=[id_var], var_name=var_name, value_name=value_name)
    df_melted["Sku_Market"] = df_melted[id_var].astype(str) + sep + df_melted[var_name].astype(str)
    df_melted["sku_name"] = df_melted["Sku_Market"].str.split(sep).str[0]
    df_final = df_melted[["Sku_Market", "sku_name", var_name, value_name]]
    df_final = df_final[(df_final[value_name].notna()) & (df_final[value_name] != 0)]
    return df_final


def extract_append_pnl_feature(
    df: pd.DataFrame,
    field_name: str,
    start_cell: str,
    end_cell: str,
    excel_path: str,
    sheet_name: str,
    engine: str,
    market_column: str,
) -> pd.DataFrame:
    df_feature = extract_excel_table_to_df(
        excel_path, sheet_name, start_cell, end_cell, header_row=True, engine=engine
    )
    if df_feature.columns.duplicated().any():
        df_feature = df_feature.loc[:, ~df_feature.columns.duplicated(keep="first")]
    df_feature.set_index(df_feature.columns[0], inplace=True)

    def lookup_value(row: pd.Series):
        try:
            return df_feature.loc[row["sku_name"], row[market_column]]
        except KeyError:
            return None

    df[field_name] = df.apply(lookup_value, axis=1)
    return df


def extract_project_summary(
    excel_path: str,
    cfg: dict[str, Any],
    verbose: bool = False,
    sheet_name: str | None = None,
) -> pd.DataFrame:
    sheet = sheet_name if sheet_name is not None else cfg["sheet_name"]
    engine = cfg.get("excel", {}).get("engine", "openpyxl")
    market_col = cfg["melt"]["var_name"]
    aliases = cfg.get("sheet_name_aliases") or []

    resolved_sheet = resolve_sheet_name(excel_path, sheet, engine, aliases=aliases if isinstance(aliases, list) else [])

    lp = cfg["list_price_table"]
    filename = excel_path.split("/")[-1]
    file_year = extract_year(excel_path)
    gate_no = extract_gate_number(excel_path)
    proj_name = extract_project_name(excel_path)

    df = extract_excel_table_to_df(
        excel_path,
        resolved_sheet,
        lp["start_cell"],
        lp["end_cell"],
        bool(lp.get("header_row", True)),
        engine,
    )
    df = clean_list_price_df(df, cfg)
    df = melt_list_price_df(df, cfg)
    if df.empty:
        return df

    df_final = df.copy()
    if verbose:
        print(df_final)

    for row in cfg["feature_tables"]:
        df_final = extract_append_pnl_feature(
            df_final,
            field_name=row["field_name"],
            start_cell=row["start_cell"],
            end_cell=row["end_cell"],
            excel_path=excel_path,
            sheet_name=resolved_sheet,
            engine=engine,
            market_column=market_col,
        )

    df_final.insert(0, "filename", filename)
    df_final.insert(1, "sheet_name", resolved_sheet)
    df_final.insert(2, "file_year", file_year)
    df_final.insert(3, "project_name", proj_name)
    df_final.insert(4, "gate_number", gate_no)

    renames = cfg.get("rename_columns") or {}
    if renames:
        df_final = df_final.rename(columns=renames)

    return df_final


def extract_all_proj_summary_aft2024(
    directory_path: str,
    cfg: dict[str, Any],
    verbose: bool = False,
    sheet_name: str | None = None,
) -> tuple[pd.DataFrame | None, dict[str, str]]:
    directory_path = os.path.abspath(os.path.expanduser(directory_path))
    if not os.path.isdir(directory_path):
        raise NotADirectoryError(f"Not a directory: {directory_path}")

    walk = cfg["walk"]
    extensions = tuple(walk["extensions"])
    ignore_prefix = walk.get("ignore_filename_prefix", "~")

    if verbose:
        print(f"Starting extract_all_proj_summary_aft2024. CWD: {os.getcwd()}")
        print(f"Input directory: {directory_path}")

    master_df: pd.DataFrame | None = None
    not_processed_files: dict[str, str] = {}

    excel_files: list[str] = []
    for root, _dirs, files in os.walk(directory_path):
        for filename in files:
            if filename.endswith(extensions) and not filename.startswith(ignore_prefix):
                filepath = os.path.join(root, filename)
                rel_path = os.path.relpath(filepath, directory_path)
                excel_files.append(rel_path)

    total_files = len(excel_files)
    if verbose:
        print(f"Found {total_files} Excel files to process.")

    for idx, rel in enumerate(excel_files, 1):
        file_path = os.path.join(directory_path, rel)
        if verbose:
            print(f"[{idx}/{total_files}] Processing: {file_path}")

        try:
            df = extract_project_summary(file_path, cfg, verbose=verbose, sheet_name=sheet_name)
            if df is not None and not df.empty:
                master_df = df if master_df is None else pd.concat([master_df, df], ignore_index=True)
            else:
                not_processed_files[rel] = "empty dataframe"
        except Exception as e:  # noqa: BLE001
            if verbose:
                print(f"Error loading {file_path}: {e}")
            err_str = str(e).lower()
            if "no sheet matching" in err_str or (
                "worksheet named" in err_str and "not found" in err_str
            ) or ("sheet" in err_str and "not found" in err_str):
                reason = "no sheet found"
            elif "more than" in err_str and "skus" in err_str:
                reason = "more than 10 skus"
            elif "index" in err_str and "out of bounds" in err_str:
                reason = "check irregular row arrangement - index out of bounds"
            elif isinstance(e, KeyError):
                reason = "missing column (check headers vs config melt.id_var_column / list price)"
            else:
                reason = "others"
            not_processed_files[rel] = f"{reason} | {type(e).__name__}: {e}"

    if verbose:
        print(f"Finished processing {total_files} files.")
        if not_processed_files:
            print("The following files were not processed (due to errors or empty data):")
            for fname, reason in not_processed_files.items():
                print(f" - {fname}: {reason}")
        else:
            print("All files processed successfully.")

    return master_df, not_processed_files


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract P&L Project Summary from all Excel files under a directory.")
    parser.add_argument(
        "directory",
        type=str,
        help="Root folder to walk for Excel files (see config walk.extensions)",
    )
    parser.add_argument(
        "-c",
        "--config",
        type=str,
        default=None,
        help=f"JSON config path (default: {default_config_path()})",
    )
    parser.add_argument(
        "--sheet",
        type=str,
        default=None,
        help="Override sheet name from config for this run",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        default=None,
        help="Optional path to write combined CSV (UTF-8)",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Print progress and per-file errors")
    args = parser.parse_args()

    try:
        cfg = load_config(args.config)
    except (FileNotFoundError, ValueError, json.JSONDecodeError) as e:
        print(str(e), file=sys.stderr)
        return 2

    try:
        master_df, skipped = extract_all_proj_summary_aft2024(
            args.directory,
            cfg,
            verbose=args.verbose,
            sheet_name=args.sheet,
        )
    except NotADirectoryError as e:
        print(str(e), file=sys.stderr)
        return 2

    if master_df is None or master_df.empty:
        print("No rows extracted.", file=sys.stderr)
        if skipped:
            for fname, reason in skipped.items():
                print(f"  skipped: {fname}: {reason}", file=sys.stderr)
        return 1

    if args.output:
        master_df.to_csv(args.output, index=False)
        if args.verbose:
            print(f"Wrote {len(master_df)} rows to {args.output}")

    if not args.output:
        print(master_df.head().to_string())

    if skipped and args.verbose:
        print(f"Skipped {len(skipped)} file(s).")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
