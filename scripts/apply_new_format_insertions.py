#!/usr/bin/env python3
"""
For each unique ``filename`` in extractions CSV, copy the new-format template and
write extracted values. Output: ``<original_stem>_NPL.xlsm`` in the output folder.

Failures are collated to error_insertion.csv (default under --output-dir).

Example:
  python scripts/apply_new_format_insertions.py \\
    --extractions extractions.csv \\
    --extractions-proj extractions_proj.csv \\
    --template /Users/kai/Work/iNova/inova/pnl-data/new_format/iNova LOCAL P&L Project BLANK TEMPLATE_Master V2.1.xlsm \\
    --output-dir /Users/kai/Work/iNova/inova/pnl-data/new_format
"""

from __future__ import annotations

import argparse
import os
import sys

import pandas as pd
from tqdm import tqdm

_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(_SCRIPTS_DIR)
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

from pnl_lib.errors import write_insertion_errors_csv  # noqa: E402
from pnl_lib.insert_template import copy_template_and_insert, npl_output_name  # noqa: E402
from pnl_lib.paths import default_new_format_output_dir, default_new_format_template_path  # noqa: E402

DEFAULT_TEMPLATE = default_new_format_template_path()
DEFAULT_OUTPUT_DIR = default_new_format_output_dir()


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply extractions into new-format P&L templates.")
    parser.add_argument("--extractions", default=os.path.join(_REPO_ROOT, "extractions.csv"))
    parser.add_argument("--extractions-proj", default=os.path.join(_REPO_ROOT, "extractions_proj.csv"))
    parser.add_argument("--template", default=DEFAULT_TEMPLATE)
    parser.add_argument("--output-dir", default=DEFAULT_OUTPUT_DIR)
    parser.add_argument(
        "--error-insertion",
        default=None,
        help="Collated insertion failures CSV (default: <output-dir>/error_insertion.csv)",
    )
    args = parser.parse_args()

    template = os.path.abspath(args.template)
    output_dir = os.path.abspath(args.output_dir)
    errors_path = os.path.abspath(args.error_insertion or os.path.join(output_dir, "error_insertion.csv"))
    if not os.path.isfile(template):
        print(f"Template not found: {template}", file=sys.stderr)
        return 2

    extractions = pd.read_csv(args.extractions)
    if "filename" not in extractions.columns:
        print("extractions CSV must include a 'filename' column", file=sys.stderr)
        return 2

    if os.path.isfile(args.extractions_proj):
        extractions_proj = pd.read_csv(args.extractions_proj)
    else:
        extractions_proj = pd.DataFrame()
        print(f"No project CSV at {args.extractions_proj}; A&P TABLE writes skipped.")

    filenames = list(extractions["filename"].dropna().unique())
    if not filenames:
        print("No filenames in extractions.", file=sys.stderr)
        return 1

    os.makedirs(output_dir, exist_ok=True)
    failed: dict[str, str] = {}

    for fname in tqdm(filenames, desc="Writing _NPL workbooks", unit="file"):
        sku_rows = extractions[extractions["filename"] == fname].reset_index(drop=True)
        if extractions_proj.empty or "filename" not in extractions_proj.columns:
            proj_rows = pd.DataFrame()
        else:
            proj_rows = extractions_proj[extractions_proj["filename"] == fname].reset_index(drop=True)

        out_name = npl_output_name(fname)
        out_path = os.path.join(output_dir, out_name)
        try:
            copy_template_and_insert(template, out_path, sku_rows, proj_rows)
        except Exception as exc:  # noqa: BLE001
            failed[fname] = f"{type(exc).__name__}: {exc}"

    errors_df = write_insertion_errors_csv(failed, errors_path)
    if errors_df.empty:
        print(f"No insertion errors -> {errors_path} (header only)")
    else:
        print(f"Wrote {len(errors_df)} insertion error row(s) -> {errors_path}")

    print(f"\nDone. Wrote {len(filenames) - len(failed)} workbook(s) to {output_dir}")
    if failed:
        print(f"Failed ({len(failed)}):", file=sys.stderr)
        for fname, reason in failed.items():
            print(f"  {fname}: {reason}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
