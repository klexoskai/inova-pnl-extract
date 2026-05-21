#!/usr/bin/env python3
"""
Pre-flight report for old-format P&L files → new-format insertion.

Runs SKU and project extraction (and optionally an in-memory insertion dry-run)
without writing ``_NPL`` workbooks. Output CSV is for finance review:

  filename, can_process (yes/no), notes

``notes`` combines automated findings (issues block processing; warnings are
informational). Finance can add manual discrepancy notes in the same column after export.

Example:
  python scripts/report_pnl_readiness.py \\
    --input-dir /Users/kai/Work/iNova/inova/pnl-data/old_format \\
    --output /Users/kai/Work/iNova/inova/inova-pnl-extract/pnl_readiness_report.csv
"""

from __future__ import annotations

import argparse
import os
import sys

_SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO_ROOT = os.path.dirname(_SCRIPTS_DIR)
if _SCRIPTS_DIR not in sys.path:
    sys.path.insert(0, _SCRIPTS_DIR)

from pnl_lib.paths import default_new_format_template_path  # noqa: E402
from pnl_lib.preflight import assess_directory, write_readiness_report_csv  # noqa: E402

DEFAULT_INPUT = "/Users/kai/Work/iNova/inova/pnl-data/old_format"
DEFAULT_OUTPUT = os.path.join(_REPO_ROOT, "pnl_readiness_report.csv")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Report which P&L files can be processed for _NPL insertion (no conversion)."
    )
    parser.add_argument(
        "--input-dir",
        default=DEFAULT_INPUT,
        help=f"Directory of old-format workbooks (default: {DEFAULT_INPUT})",
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT,
        help=f"Readiness report CSV path (default: {DEFAULT_OUTPUT})",
    )
    parser.add_argument(
        "--template",
        default=None,
        help="New-format master template .xlsm (default: config/pnl_pipeline.json)",
    )
    parser.add_argument(
        "--no-preproc",
        action="store_true",
        help="Skip COGS_other_combined_y* preproc on SKU extract (matches extract CLI)",
    )
    parser.add_argument(
        "--extract-only",
        action="store_true",
        help="Skip in-memory insertion dry-run (faster; less thorough)",
    )
    parser.add_argument(
        "--no-progress",
        action="store_true",
        help="Disable tqdm progress bar",
    )
    args = parser.parse_args()

    input_dir = os.path.abspath(args.input_dir)
    if not os.path.isdir(input_dir):
        print(f"Input directory not found: {input_dir}", file=sys.stderr)
        return 2

    template_path = os.path.abspath(args.template or default_new_format_template_path())
    output_path = os.path.abspath(args.output)

    print(f"Input:    {input_dir}")
    print(f"Template: {template_path}")
    print(f"Output:   {output_path}")
    if args.extract_only:
        print("Mode:     extract validation only (no insertion dry-run)")
    else:
        print("Mode:     extract + in-memory insertion dry-run")

    results = assess_directory(
        input_dir,
        template_path=template_path,
        apply_preproc=not args.no_preproc,
        run_insertion_dry_run=not args.extract_only,
        progress=not args.no_progress,
    )

    if not results:
        print("No Excel files found.", file=sys.stderr)
        df = write_readiness_report_csv([], output_path)
        print(f"Wrote empty report -> {output_path}")
        return 1

    df = write_readiness_report_csv(results, output_path)
    n_yes = int((df["can_process"] == "yes").sum())
    n_no = int((df["can_process"] == "no").sum())
    print(f"\nWrote {len(df)} row(s) -> {output_path}")
    print(f"  can_process=yes: {n_yes}")
    print(f"  can_process=no:  {n_no}")

    if n_no:
        print("\nBlocked files:", file=sys.stderr)
        blocked = df.loc[df["can_process"] == "no", ["filename", "notes"]]
        for _, row in blocked.iterrows():
            print(f"  {row['filename']}: {row['notes']}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
