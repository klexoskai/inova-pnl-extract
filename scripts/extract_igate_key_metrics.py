#!/usr/bin/env python3
"""
Extract fixed rectangular ranges from sheet 'iGate Key Metrics (005)' (or overrides via config).

Hyperparameters live in config/igate_key_metrics_extract.json by default.

Reuses range slicing from scripts/extract_pnl_directory.py.

Usage:
  python scripts/extract_igate_key_metrics.py /path/to/file.xlsm
  python scripts/extract_igate_key_metrics.py /path/to/folder -o combined.csv
  python scripts/extract_igate_key_metrics.py /path/to/folder -c config/custom_igate.json -v
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any

import pandas as pd

# Allow `import extract_pnl_directory` when executed as `python scripts/extract_igate_key_metrics.py`
_SCRIPTS_DIR = Path(__file__).resolve().parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from extract_pnl_directory import extract_excel_table_to_df, resolve_sheet_name  # noqa: E402


def default_config_path() -> str:
    return str(Path(__file__).resolve().parent.parent / "config" / "igate_key_metrics_extract.json")


def load_config(path: str | None) -> dict[str, Any]:
    config_path = Path(path or default_config_path()).expanduser().resolve()
    if not config_path.is_file():
        raise FileNotFoundError(f"Config file not found: {config_path}")
    try:
        with open(config_path, encoding="utf-8") as f:
            cfg: dict[str, Any] = json.load(f)
    except json.JSONDecodeError as e:
        raise ValueError(f"Invalid JSON in config {config_path}: {e}") from e

    if "sheet_name" not in cfg:
        raise ValueError("Config missing sheet_name")
    if "tables" not in cfg or not isinstance(cfg["tables"], list) or not cfg["tables"]:
        raise ValueError("Config missing non-empty tables list")
    for i, t in enumerate(cfg["tables"]):
        for key in ("id", "start_cell", "end_cell"):
            if key not in t:
                raise ValueError(f"tables[{i}] missing '{key}'")
    if "walk" not in cfg or "extensions" not in cfg["walk"]:
        raise ValueError("walk.extensions is required")

    return cfg


def _collect_excel_paths(target: str, cfg: dict[str, Any]) -> list[str]:
    target = os.path.abspath(os.path.expanduser(target))
    walk = cfg["walk"]
    extensions = tuple(walk["extensions"])
    ignore_prefix = walk.get("ignore_filename_prefix", "~")

    if os.path.isfile(target):
        if not target.endswith(extensions):
            raise ValueError(f"File does not match configured extensions {extensions}: {target}")
        return [target]
    if not os.path.isdir(target):
        raise NotADirectoryError(f"Not a file or directory: {target}")

    paths: list[str] = []
    for root, _dirs, files in os.walk(target):
        for name in files:
            if name.endswith(extensions) and not name.startswith(ignore_prefix):
                paths.append(os.path.join(root, name))
    paths.sort()
    return paths


def _normalize_block(
    df: pd.DataFrame,
    *,
    source_file: str,
    table_id: str,
) -> pd.DataFrame:
    """Rename body columns to col_0..col_n and add source_file, table_id, row_index."""
    out = df.copy()
    out = out.reset_index(drop=True)
    ncols = out.shape[1]
    out.columns = [f"col_{i}" for i in range(ncols)]
    out.insert(0, "row_index", range(len(out)))
    out.insert(0, "table_id", table_id)
    out.insert(0, "source_file", os.path.basename(source_file))
    return out


def extract_from_workbook(excel_path: str, cfg: dict[str, Any]) -> list[pd.DataFrame]:
    engine = cfg.get("excel", {}).get("engine", "openpyxl")
    aliases = cfg.get("sheet_name_aliases") or []
    if not isinstance(aliases, list):
        aliases = []
    resolved_sheet = resolve_sheet_name(excel_path, cfg["sheet_name"], engine, aliases=aliases)
    frames: list[pd.DataFrame] = []

    for spec in cfg["tables"]:
        tbl = extract_excel_table_to_df(
            excel_path,
            resolved_sheet,
            spec["start_cell"],
            spec["end_cell"],
            bool(spec.get("header_row", False)),
            engine,
        )
        frames.append(_normalize_block(tbl, source_file=excel_path, table_id=spec["id"]))
    return frames


def run(target: str, cfg: dict[str, Any], verbose: bool = False) -> tuple[pd.DataFrame, dict[str, str]]:
    skipped: dict[str, str] = {}
    all_parts: list[pd.DataFrame] = []

    try:
        paths = _collect_excel_paths(target, cfg)
    except (NotADirectoryError, ValueError) as e:
        raise e

    if not paths:
        return pd.DataFrame(), {"": "no matching excel files"}

    for idx, path in enumerate(paths, 1):
        if verbose:
            print(f"[{idx}/{len(paths)}] {path}")
        try:
            parts = extract_from_workbook(path, cfg)
        except Exception as e:  # noqa: BLE001
            base = os.path.abspath(os.path.expanduser(target))
            rel = os.path.relpath(path, base) if os.path.isdir(base) else os.path.basename(path)
            skipped[rel] = str(e)
            if verbose:
                print(f"  error: {e}")
            continue
        all_parts.extend(parts)

    if not all_parts:
        return pd.DataFrame(), skipped

    return pd.concat(all_parts, ignore_index=True), skipped


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract configured rectangular tables from iGate Key Metrics sheet."
    )
    parser.add_argument("path", type=str, help="Excel file or directory to scan")
    parser.add_argument(
        "-c",
        "--config",
        type=str,
        default=None,
        help=f"JSON config (default: {default_config_path()})",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        default=None,
        help="Write combined long CSV (UTF-8)",
    )
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    try:
        cfg = load_config(args.config)
    except (FileNotFoundError, ValueError) as e:
        print(str(e), file=sys.stderr)
        return 2

    try:
        master, skipped = run(args.path, cfg, verbose=args.verbose)
    except (NotADirectoryError, ValueError) as e:
        print(str(e), file=sys.stderr)
        return 2

    if master.empty:
        print("No rows extracted.", file=sys.stderr)
        for k, v in skipped.items():
            print(f"  skipped {k}: {v}", file=sys.stderr)
        return 1

    if args.output:
        master.to_csv(args.output, index=False)
        if args.verbose:
            print(f"Wrote {len(master)} rows to {args.output}")
    else:
        print(master.head(20).to_string())

    if skipped and args.verbose:
        print("Skipped:", skipped)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
