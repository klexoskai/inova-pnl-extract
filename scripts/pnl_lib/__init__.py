"""Shared P&L extraction and new-format template insertion logic."""

from .extract_proj import extract_project_metrics, extract_project_metrics_directory
from .extract_sku import extract_sku_summary, extract_sku_summary_directory, preproc_sku_extractions
from .insert_template import (
    apply_insertions_to_workbook,
    parse_launch_year_q,
    write_cogs_other_combined,
    write_cogs_unit,
    write_sales_ded_ka_launch,
    write_sales_ded_td_gtn,
    write_sp_list_prices,
    write_sku_block_sheet,
    write_vol_auto_forecast_volumes,
)

__all__ = [
    "extract_sku_summary",
    "extract_sku_summary_directory",
    "preproc_sku_extractions",
    "extract_project_metrics",
    "extract_project_metrics_directory",
    "apply_insertions_to_workbook",
    "parse_launch_year_q",
    "write_sku_block_sheet",
    "write_sp_list_prices",
    "write_vol_auto_forecast_volumes",
    "write_sales_ded_td_gtn",
    "write_cogs_unit",
    "write_cogs_other_combined",
    "write_sales_ded_ka_launch",
]
