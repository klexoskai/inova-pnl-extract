"""Write collated error reports for extraction and insertion pipelines."""

from __future__ import annotations

import os
from typing import Any

import pandas as pd

EXTRACTION_ERROR_COLUMNS = ["filename", "file_path", "stage", "reason"]
INSERTION_ERROR_COLUMNS = ["filename", "output_file", "error_type", "reason"]


def _basename(path_or_name: str) -> str:
    return os.path.basename(path_or_name.replace("\\", "/"))


def skipped_to_extraction_rows(
    skipped: dict[str, str],
    stage: str,
) -> list[dict[str, Any]]:
    """Turn {relative_path: reason} into rows for error_extraction.csv."""
    rows: list[dict[str, Any]] = []
    for file_path, reason in skipped.items():
        rows.append(
            {
                "filename": _basename(file_path),
                "file_path": file_path,
                "stage": stage,
                "reason": reason,
            }
        )
    return rows


def write_extraction_errors_csv(
    sku_skipped: dict[str, str],
    proj_skipped: dict[str, str],
    output_path: str,
) -> pd.DataFrame:
    rows = skipped_to_extraction_rows(sku_skipped, "sku")
    rows.extend(skipped_to_extraction_rows(proj_skipped, "project"))
    df = pd.DataFrame(rows, columns=EXTRACTION_ERROR_COLUMNS)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    df.to_csv(output_path, index=False)
    return df


def write_insertion_errors_csv(
    failed: dict[str, str],
    output_path: str,
    *,
    npl_name_fn: Any | None = None,
) -> pd.DataFrame:
    if npl_name_fn is None:
        from .insert_template import npl_output_name

        npl_name_fn = npl_output_name

    rows: list[dict[str, Any]] = []
    for filename, message in failed.items():
        error_type, _, reason = message.partition(": ")
        if not reason:
            reason = message
            error_type = "Error"
        rows.append(
            {
                "filename": filename,
                "output_file": npl_name_fn(filename),
                "error_type": error_type,
                "reason": reason,
            }
        )
    df = pd.DataFrame(rows, columns=INSERTION_ERROR_COLUMNS)
    os.makedirs(os.path.dirname(os.path.abspath(output_path)) or ".", exist_ok=True)
    df.to_csv(output_path, index=False)
    return df
