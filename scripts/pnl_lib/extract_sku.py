"""SKU-level extraction from old-format P&L workbooks (Project Summary sheet)."""

from __future__ import annotations

import os
from typing import Any

import pandas as pd

from .excel_utils import (
    basename_only,
    extract_excel_table_to_df,
    extract_gate_number,
    extract_project_name,
    extract_year,
    list_excel_files,
    resolve_sheet_name,
)

DEFAULT_SHEET = "Project Summary"

SKU_FEATURE_FIELDS: list[tuple[str, str, str]] = [
    ("forecast_volume_y1", "BW2", "DC15"),
    ("forecast_volume_y2", "BW17", "DC30"),
    ("forecast_volume_y3", "BW32", "DC45"),
    ("forecast_volume_y4", "BW47", "DC58"),
    ("forecast_volume_y5", "BW62", "DC73"),
    ("GTN_perc excluding Launch yr one-time costs", "C16", "AI27"),
    ("COGS/unit", "C69", "AI80"),
    ("Launch Yr one-time cost (Listing fees, launch COOP)", "C30", "AI41"),
    ("forecast_net_sales_y1", "HD2", "IJ15"),
    ("forecast_net_sales_y2", "HD17", "IJ30"),
    ("forecast_net_sales_y3", "HD32", "IJ45"),
    ("forecast_net_sales_y4", "HD47", "IJ58"),
    ("forecast_net_sales_y5", "HD62", "IJ73"),
    ("launchyr_coop_listingfees_y1", "FV2", "HB15"),
    ("launchyr_coop_listingfees_y2", "FV17", "HB30"),
    ("launchyr_coop_listingfees_y3", "FV32", "HB45"),
    ("other_COGS_y1", "JT2", "KZ13"),
    ("other_COGS_y2", "JT17", "KZ28"),
    ("other_COGS_y3", "JT32", "KZ43"),
    ("other_COGS_y4", "JT47", "KZ58"),
    ("other_COGS_y5", "JT62", "KZ73"),
    ("one_time_costs_y1", "LC2", "MI13"),
    ("one_time_costs_y2", "LC17", "MI28"),
    ("one_time_costs_y3", "LC32", "MI43"),
    ("one_time_costs_y4", "LC47", "MI58"),
    ("one_time_costs_y5", "LC62", "MI73"),
    ("GP_y1", "ML2", "NP13"),
    ("GP_y2", "ML17", "NP28"),
    ("GP_y3", "ML32", "NP43"),
    ("GP_y4", "ML47", "NP58"),
    ("GP_y5", "ML62", "NP73"),
    ("forecast_gross_sales_y1", "DF2", "EK15"),
    ("forecast_gross_sales_y2", "DF17", "EK30"),
    ("forecast_gross_sales_y3", "DF32", "EK45"),
    ("forecast_gross_sales_y4", "DF47", "EK58"),
    ("forecast_gross_sales_y5", "DF62", "EK73"),
]


def clean_list_price_df(df: pd.DataFrame) -> pd.DataFrame:
    list_price_col = df.columns[df.columns.astype(str).str.contains("List Price", case=False, na=False)]
    if not list_price_col.empty:
        col_name = list_price_col[0]
        df = df[~df[col_name].astype(str).str.contains("input", case=False, na=False)]
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~((df == 0) | (df.isna())).all(axis=0)]
    if not df.empty:
        df = df.drop(df.index[0]).reset_index(drop=True)
    return df


def melt_list_price_df(df: pd.DataFrame) -> pd.DataFrame:
    id_col = [c for c in df.columns if "List Price" in str(c)][0]
    df_melted = df.melt(id_vars=[id_col], var_name="Market", value_name="list_price_AUD")
    df_melted["Sku_Market"] = df_melted[id_col].astype(str) + "_" + df_melted["Market"].astype(str)
    df_melted["sku_name"] = df_melted["Sku_Market"].str.split("_").str[0]
    df_final = df_melted[["Sku_Market", "sku_name", "Market", "list_price_AUD"]]
    return df_final[(df_final["list_price_AUD"].notna()) & (df_final["list_price_AUD"] != 0)]


def extract_append_pnl_feature(
    df: pd.DataFrame,
    field_name: str,
    start_cell: str,
    end_cell: str,
    excel_path: str,
    sheet_name: str,
) -> pd.DataFrame:
    df_feature = extract_excel_table_to_df(excel_path, sheet_name, start_cell, end_cell, header_row=True)
    if df_feature.columns.duplicated().any():
        df_feature = df_feature.loc[:, ~df_feature.columns.duplicated(keep="first")]
    df_feature.set_index(df_feature.columns[0], inplace=True)

    def lookup_value(row: pd.Series) -> Any:
        try:
            return df_feature.loc[row["sku_name"], row["Market"]]
        except KeyError:
            return None

    df[field_name] = df.apply(lookup_value, axis=1)
    return df


def extract_sku_summary(excel_path: str, sheet_name: str = DEFAULT_SHEET) -> pd.DataFrame:
    resolved = resolve_sheet_name(excel_path, sheet_name)
    base = extract_excel_table_to_df(excel_path, resolved, "C3", "AI14", header_row=True)
    base = clean_list_price_df(base)
    base = melt_list_price_df(base)
    if base.empty:
        return base

    df_final = base.copy()
    for field_name, start_cell, end_cell in SKU_FEATURE_FIELDS:
        df_final = extract_append_pnl_feature(
            df_final,
            field_name=field_name,
            start_cell=start_cell,
            end_cell=end_cell,
            excel_path=excel_path,
            sheet_name=resolved,
        )

    filename = basename_only(excel_path)
    df_final.insert(0, "filename", filename)
    df_final.insert(1, "sheet_name", resolved)
    df_final.insert(2, "file_year", extract_year(excel_path))
    df_final.insert(3, "project_name", extract_project_name(excel_path))
    df_final.insert(4, "gate_number", extract_gate_number(excel_path))
    if "Sku_Market" in df_final.columns:
        df_final = df_final.rename(columns={"Sku_Market": "sku_market"})
    return df_final


def preproc_sku_extractions(df: pd.DataFrame) -> pd.DataFrame:
    """Add COGS_other_combined_y1..y5 from one_time_costs + other_COGS."""
    out = df.copy()

    def _num_series(frame: pd.DataFrame, col_name: str) -> pd.Series:
        if col_name not in frame.columns:
            return pd.Series(0.0, index=frame.index, dtype=float)
        return pd.to_numeric(frame[col_name], errors="coerce").fillna(0.0)

    for y in range(1, 6):
        out[f"COGS_other_combined_y{y}"] = _num_series(out, f"one_time_costs_y{y}") + _num_series(
            out, f"other_COGS_y{y}"
        )
    return out


def extract_sku_summary_directory(
    directory: str,
    *,
    apply_preproc: bool = True,
    progress: bool = True,
) -> tuple[pd.DataFrame, dict[str, str]]:
    directory = os.path.abspath(directory)
    paths = list_excel_files(directory)
    frames: list[pd.DataFrame] = []
    skipped: dict[str, str] = {}

    iterator = paths
    if progress:
        from tqdm import tqdm

        iterator = tqdm(paths, desc="SKU extract", unit="file")

    for path in iterator:
        rel = os.path.relpath(path, directory)
        try:
            df = extract_sku_summary(path)
            if df.empty:
                skipped[rel] = "empty dataframe"
            else:
                frames.append(df)
        except Exception as exc:  # noqa: BLE001
            err = str(exc).lower()
            if "sheet" in err and "not found" in err:
                reason = "no sheet found"
            elif "index" in err and "out of bounds" in err:
                reason = "check irregular row arrangement - index out of bounds"
            else:
                reason = f"{type(exc).__name__}: {exc}"
            skipped[rel] = reason

    if not frames:
        return pd.DataFrame(), skipped

    master = pd.concat(frames, ignore_index=True)
    if apply_preproc:
        master = preproc_sku_extractions(master)
    return master, skipped
