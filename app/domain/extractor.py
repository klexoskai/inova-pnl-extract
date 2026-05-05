import os
import re
from dataclasses import dataclass
from tempfile import TemporaryDirectory
from typing import Iterable

import pandas as pd
from openpyxl.utils import coordinate_to_tuple


ALLOWED_EXTENSIONS = (".xlsx", ".xls", ".xlsm")


def extract_project_name(filepath: str) -> str | None:
    filename = filepath.split("/")[-1] if "/" in filepath else filepath
    matches = re.findall(r"-\s*([^-\n\r]+?)\s*-", filename)
    if not matches:
        return None
    return matches[0].strip()


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
    filepath: str, sheet_name: str | int, start_cell: str, end_cell: str, header_row: bool = False
) -> pd.DataFrame:
    df_raw = pd.read_excel(filepath, sheet_name=sheet_name, header=None, engine="openpyxl")
    start_row, start_col = coordinate_to_tuple(start_cell)
    end_row, end_col = coordinate_to_tuple(end_cell)
    df_range = df_raw.iloc[start_row - 1 : end_row, start_col - 1 : end_col]
    if header_row:
        df_range.columns = df_range.iloc[0]
        df_range = df_range.iloc[1:].reset_index(drop=True)
    else:
        df_range = df_range.reset_index(drop=True)
    return df_range


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
    df_melted = df.melt(id_vars=["List Price"], var_name="Market", value_name="list_price_AUD")
    df_melted["Sku_Market"] = df_melted["List Price"].astype(str) + "_" + df_melted["Market"].astype(str)
    df_melted["sku_name"] = df_melted["Sku_Market"].str.split("_").str[0]
    df_final = df_melted[["Sku_Market", "sku_name", "Market", "list_price_AUD"]]
    return df_final[(df_final["list_price_AUD"].notna()) & (df_final["list_price_AUD"] != 0)]


def extract_append_pnl_feature(
    df: pd.DataFrame, field_name: str, start_cell: str, end_cell: str, excel_path: str, sheet_name: str = "Project Summary"
) -> pd.DataFrame:
    df_feature = extract_excel_table_to_df(excel_path, sheet_name, start_cell, end_cell, header_row=True)
    df_feature.set_index(df_feature.columns[0], inplace=True)

    def lookup_value(row: pd.Series):
        try:
            return df_feature.loc[row["sku_name"], row["Market"]]
        except KeyError:
            return None

    df[field_name] = df.apply(lookup_value, axis=1)
    return df


def extract_project_summary(excel_path: str, sheet_name: str = "Project Summary") -> pd.DataFrame:
    filename = excel_path.split("/")[-1]
    file_year = extract_year(excel_path)
    gate_no = extract_gate_number(excel_path)
    proj_name = extract_project_name(excel_path)

    base_df = extract_excel_table_to_df(excel_path, sheet_name, "C3", "AI14", header_row=True)
    base_df = clean_list_price_df(base_df)
    base_df = melt_list_price_df(base_df)
    if base_df.empty:
        return base_df

    fields = [
        ("forecast_volume_y1", "BW2", "DC15"),
        ("forecast_volume_y2", "BW17", "DC30"),
        ("forecast_volume_y3", "BW32", "DC45"),
        ("forecast_net_sales_y1", "HD2", "IJ15"),
        ("forecast_net_sales_y2", "HD17", "IJ30"),
        ("forecast_net_sales_y3", "HD32", "IJ45"),
        ("launchyr_coop_listingfees_y1", "FV2", "HB15"),
        ("launchyr_coop_listingfees_y2", "FV17", "HB30"),
        ("launchyr_coop_listingfees_y3", "FV32", "HB45"),
        ("one_time_costs_y1", "LC2", "MI15"),
        ("one_time_costs_y2", "LC17", "MI30"),
        ("one_time_costs_y3", "LC32", "MI45"),
    ]

    df_final = base_df.copy()
    for field_name, start_cell, end_cell in fields:
        df_final = extract_append_pnl_feature(
            df_final, field_name=field_name, start_cell=start_cell, end_cell=end_cell, excel_path=excel_path, sheet_name=sheet_name
        )

    df_final.insert(0, "filename", filename)
    df_final.insert(1, "sheet_name", sheet_name)
    df_final.insert(2, "file_year", file_year)
    df_final.insert(3, "project_name", proj_name)
    df_final.insert(4, "gate_number", gate_no)
    if "Sku_Market" in df_final.columns:
        df_final = df_final.rename(columns={"Sku_Market": "sku_market"})
    return df_final


@dataclass
class ExtractionOutcome:
    data: pd.DataFrame
    skipped_files: dict[str, str]
    processed_files: int


def extract_all_project_summaries(directory_path: str) -> ExtractionOutcome:
    all_frames: list[pd.DataFrame] = []
    not_processed_files: dict[str, str] = {}
    processed_files = 0

    for root, _, files in os.walk(directory_path):
        for filename in files:
            if not filename.endswith(ALLOWED_EXTENSIONS) or filename.startswith("~"):
                continue
            full_path = os.path.join(root, filename)
            rel_path = os.path.relpath(full_path, directory_path)
            try:
                df = extract_project_summary(full_path)
                if df is None or df.empty:
                    not_processed_files[rel_path] = "empty dataframe"
                    continue
                processed_files += 1
                all_frames.append(df)
            except Exception as exc:  # noqa: BLE001
                msg = str(exc).lower()
                if "no sheet found" in msg:
                    reason = "no sheet found"
                elif "more than" in msg and "skus" in msg:
                    reason = "more than 10 skus"
                elif "index" in msg and "out of bounds" in msg:
                    reason = "irregular row arrangement"
                else:
                    reason = "others"
                not_processed_files[rel_path] = reason

    output = pd.concat(all_frames, ignore_index=True) if all_frames else pd.DataFrame()
    return ExtractionOutcome(data=output, skipped_files=not_processed_files, processed_files=processed_files)


def write_bytes_to_tempfiles(files: Iterable[tuple[str, bytes]], base_dir: str) -> None:
    for rel_path, content in files:
        local_path = os.path.join(base_dir, rel_path)
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        with open(local_path, "wb") as f:
            f.write(content)


def extract_from_bytes(files: Iterable[tuple[str, bytes]]) -> ExtractionOutcome:
    with TemporaryDirectory() as tmp_dir:
        write_bytes_to_tempfiles(files, tmp_dir)
        return extract_all_project_summaries(tmp_dir)
