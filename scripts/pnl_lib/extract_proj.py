"""Project-level (market) extraction from old-format P&L workbooks."""

from __future__ import annotations

import os

import pandas as pd

from .excel_utils import basename_only, extract_excel_table_to_df, list_excel_files, resolve_sheet_name

DEFAULT_SHEET = "Project Summary"

AGGREGATION_COLUMNS = [
    "CANZ",
    "Asia Direct",
    "Asia Total",
    "Europe Total",
    "ME total",
    "AMENA",
    "Total iNova",
]


def active_markets_from_list_price(filepath: str, sheet_name: str) -> list[str]:
    active_df = extract_excel_table_to_df(filepath, sheet_name, "C3", "AI14", header_row=True)
    if active_df.empty:
        return []
    active_df = active_df.drop(active_df.index[0]).reset_index(drop=True)

    cols_to_keep: list[str] = []
    for col in active_df.columns:
        if col in AGGREGATION_COLUMNS:
            continue
        col_values = pd.to_numeric(active_df[col], errors="coerce")
        if col_values.notna().sum() == 0:
            continue
        if col_values.fillna(0).sum() == 0:
            continue
        cols_to_keep.append(col)

    active_df = active_df[cols_to_keep]
    active_markets: list[str] = []
    for col in active_df.columns:
        col_values = pd.to_numeric(active_df[col], errors="coerce")
        if (col_values.fillna(0) != 0).any():
            active_markets.append(str(col))
    return active_markets


def extract_active_market_metric_pivot(
    metric_df: pd.DataFrame,
    active_markets: list[str],
    value_name: str,
) -> pd.DataFrame:
    id_var = metric_df.columns[0]
    metric_active = metric_df.loc[
        :, [col for col in metric_df.columns if col in active_markets or col == id_var]
    ]
    melted = metric_active.melt(id_vars=id_var, var_name="Market", value_name=value_name)
    pivoted = melted.pivot(index="Market", columns=id_var, values=value_name).reset_index()
    pivoted.columns = [
        f"{value_name}_{col}" if col != "Market" else "Market" for col in pivoted.columns
    ]
    return pivoted


def extract_project_metrics(excel_path: str, sheet_name: str = DEFAULT_SHEET) -> pd.DataFrame:
    resolved = resolve_sheet_name(excel_path, sheet_name)
    active_markets = active_markets_from_list_price(excel_path, resolved)
    if not active_markets:
        return pd.DataFrame()

    anp_df = extract_excel_table_to_df(excel_path, resolved, "C134", "AI140", header_row=True)
    pivoted = extract_active_market_metric_pivot(
        anp_df, active_markets=active_markets, value_name="ANP_perc of Net Sales"
    )

    incr_df = extract_excel_table_to_df(excel_path, resolved, "C147", "AI152", header_row=True)
    incr_pivoted = extract_active_market_metric_pivot(
        incr_df, active_markets=active_markets, value_name="Incr Opex (Submissions, FTE etc.)"
    )
    pivoted_all = pivoted.merge(incr_pivoted, on="Market", how="left")

    third_df = extract_excel_table_to_df(excel_path, resolved, "C134", "AI145", header_row=True)
    third_df = third_df.iloc[5:].dropna(how="all").reset_index(drop=True)
    third_pivoted = extract_active_market_metric_pivot(
        third_df, active_markets=active_markets, value_name="CAPEX"
    )
    pivoted_all = pivoted_all.merge(third_pivoted, on="Market", how="left")
    pivoted_all.insert(0, "filename", basename_only(excel_path))
    return pivoted_all


def extract_project_metrics_directory(
    directory: str, *, progress: bool = True
) -> tuple[pd.DataFrame, dict[str, str]]:
    directory = os.path.abspath(directory)
    paths = list_excel_files(directory)
    frames: list[pd.DataFrame] = []
    skipped: dict[str, str] = {}

    iterator = paths
    if progress:
        from tqdm import tqdm

        iterator = tqdm(paths, desc="Project extract", unit="file")

    for path in iterator:
        rel = os.path.relpath(path, directory)
        try:
            df = extract_project_metrics(path)
            if df.empty:
                skipped[rel] = "empty dataframe"
            else:
                frames.append(df)
        except Exception as exc:  # noqa: BLE001
            skipped[rel] = f"{type(exc).__name__}: {exc}"

    if not frames:
        return pd.DataFrame(), skipped
    return pd.concat(frames, ignore_index=True), skipped
