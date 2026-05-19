#!/usr/bin/env python3
"""
Extract SKU-level and project-level metrics from old-format P&L Excel files.

Writes combined CSVs (extractions.csv, extractions_proj.csv) with a ``filename``
column so downstream insertion can group rows per source workbook.

Failures are collated to error_extraction.csv (columns: filename, file_path,
stage, reason) with stage ``sku`` or ``project``.

Example:
  python scripts/extract_old_format_pnl.py \\
    --input-dir /Users/kai/Work/iNova/inova/pnl-data/old_format \\
    --output-dir /Users/kai/Work/iNova/inova/inova-pnl-extract
"""

from __future__ import annotations

import argparse
import os
import sys

_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(_SCRIPTS_DIR)
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

from pnl_lib.errors import write_extraction_errors_csv  # noqa: E402
from pnl_lib.extract_proj import extract_project_metrics_directory  # noqa: E402
from pnl_lib.extract_sku import extract_sku_summary_directory  # noqa: E402

DEFAULT_INPUT = "/Users/kai/Work/iNova/inova/pnl-data/old_format"


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract old-format P&L files to CSV.")
    parser.add_argument(
        "--input-dir",
        default=DEFAULT_INPUT,
        help=f"Directory of old-format workbooks (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "--output-dir",
        default=_REPO_ROOT,
        help="Directory for extractions.csv and extractions_proj.csv",
    )
    parser.add_argument(
        "--extractions",
        default="extractions.csv",
        help="SKU-level output filename (under --output-dir)",
    )
    parser.add_argument(
        "--extractions-proj",
        default="extractions_proj.csv",
        help="Project-level output filename (under --output-dir)",
    )
    parser.add_argument(
        "--no-preproc",
        action="store_true",
        help="Skip COGS_other_combined_y* aggregation on SKU extract",
    )
    parser.add_argument(
        "--error-extraction",
        default="error_extraction.csv",
        help="Collated extraction failures (under --output-dir)",
    )
    args = parser.parse_args()

    input_dir = os.path.abspath(args.input_dir)
    if not os.path.isdir(input_dir):
        print(f"Input directory not found: {input_dir}", file=sys.stderr)
        return 2

    os.makedirs(args.output_dir, exist_ok=True)
    sku_out = os.path.join(args.output_dir, args.extractions)
    proj_out = os.path.join(args.output_dir, args.extractions_proj)
    errors_out = os.path.join(args.output_dir, args.error_extraction)

    print(f"SKU extraction from: {input_dir}")
    sku_df, sku_skipped = extract_sku_summary_directory(
        input_dir, apply_preproc=not args.no_preproc
    )
    if sku_df.empty:
        print("No SKU rows extracted.", file=sys.stderr)
    else:
        sku_df.to_csv(sku_out, index=False)
        print(f"Wrote {len(sku_df)} SKU rows -> {sku_out}")

    if sku_skipped:
        print(f"\nSKU skipped ({len(sku_skipped)} files):", file=sys.stderr)
        for name, reason in sku_skipped.items():
            print(f"  {name}: {reason}", file=sys.stderr)

    print(f"\nProject extraction from: {input_dir}")
    proj_df, proj_skipped = extract_project_metrics_directory(input_dir)
    if proj_df.empty:
        print("No project rows extracted.", file=sys.stderr)
    else:
        proj_df.to_csv(proj_out, index=False)
        print(f"Wrote {len(proj_df)} project rows -> {proj_out}")

    if proj_skipped:
        print(f"\nProject skipped ({len(proj_skipped)} files):", file=sys.stderr)
        for name, reason in proj_skipped.items():
            print(f"  {name}: {reason}", file=sys.stderr)

    errors_df = write_extraction_errors_csv(sku_skipped, proj_skipped, errors_out)
    if errors_df.empty:
        print(f"No extraction errors -> {errors_out} (header only)")
    else:
        print(f"Wrote {len(errors_df)} extraction error row(s) -> {errors_out}")

    if sku_df.empty and proj_df.empty:
        return 1
    if not errors_df.empty:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
