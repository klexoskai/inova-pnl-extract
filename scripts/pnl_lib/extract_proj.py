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

# Project Summary cannibalisation block: D:AI — row 274 = market labels; 276–280 = y1–y5; 282 = GP.
CANNI_COL_START = "D"
CANNI_COL_END = "AI"
CANNI_MARKET_HEADER_ROW = 274
CANNI_YEAR_DATA_ROWS = (276, 277, 278, 279, 280)  # y1 .. y5
CANNI_GP_DATA_ROW = 282


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


def _d_ai_market_row_to_long(
    filepath: str,
    sheet_name: str,
    header_row: int,
    data_row: int,
    value_col: str,
    active_markets: list[str],
) -> pd.DataFrame:
    """
    Read two consecutive D:AI rows: ``header_row`` (market labels), ``data_row`` (values).
    Return long DataFrame [Market, value_col] restricted to ``active_markets``.
    """
    top = extract_excel_table_to_df(
        filepath,
        sheet_name,
        f"{CANNI_COL_START}{header_row}",
        f"{CANNI_COL_END}{header_row}",
        header_row=False,
    )
    bot = extract_excel_table_to_df(
        filepath,
        sheet_name,
        f"{CANNI_COL_START}{data_row}",
        f"{CANNI_COL_END}{data_row}",
        header_row=False,
    )
    if top.empty or bot.empty:
        return pd.DataFrame(columns=["Market", value_col])
    hdr = top.iloc[0].tolist()
    vals = bot.iloc[0].tolist()
    active_set = {str(x).strip() for x in active_markets}
    rows: list[dict[str, object]] = []
    for hi, vi in zip(hdr, vals):
        if hi is None or (isinstance(hi, float) and pd.isna(hi)):
            continue
        mk = str(hi).strip()
        if not mk or mk not in active_set:
            continue
        rows.append({"Market": mk, value_col: vi})
    if not rows:
        return pd.DataFrame(columns=["Market", value_col])
    out = pd.DataFrame(rows).drop_duplicates(subset=["Market"], keep="first")
    return out


def merge_cannibalisation_metrics(
    filepath: str,
    sheet_name: str,
    pivoted_wide: pd.DataFrame,
    active_markets: list[str],
) -> pd.DataFrame:
    """Attach canni_perc_y1–y5 and canni_GP_perc (long→wide per Market) onto ``pivoted_wide``."""
    out = pivoted_wide
    for y_idx, drow in enumerate(CANNI_YEAR_DATA_ROWS, start=1):
        col = f"canni_perc_y{y_idx}"
        piece = _d_ai_market_row_to_long(
            filepath,
            sheet_name,
            CANNI_MARKET_HEADER_ROW,
            drow,
            col,
            active_markets,
        )
        out = out.merge(piece, on="Market", how="left")
    gp = _d_ai_market_row_to_long(
        filepath,
        sheet_name,
        CANNI_MARKET_HEADER_ROW,
        CANNI_GP_DATA_ROW,
        "canni_GP_perc",
        active_markets,
    )
    return out.merge(gp, on="Market", how="left")


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
    pivoted_all = merge_cannibalisation_metrics(
        excel_path, resolved, pivoted_all, active_markets
    )
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
