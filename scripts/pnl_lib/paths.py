"""Repo paths for the P&L extraction / insertion pipeline."""

from __future__ import annotations

import json
import os

_REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_PIPELINE_CONFIG_PATH = os.path.join(_REPO_ROOT, "config", "pnl_pipeline.json")


def load_pipeline_config() -> dict:
    with open(_PIPELINE_CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def default_new_format_template_path() -> str:
    """Absolute path to the blank new-format master template (.xlsm)."""
    cfg = load_pipeline_config()
    rel_dir = cfg.get("new_format_dir_relative", "../pnl-data/new_format")
    filename = cfg["new_format_template_filename"]
    return os.path.abspath(os.path.join(_REPO_ROOT, rel_dir, filename))


def default_new_format_output_dir() -> str:
    cfg = load_pipeline_config()
    rel_dir = cfg.get("new_format_dir_relative", "../pnl-data/new_format")
    return os.path.abspath(os.path.join(_REPO_ROOT, rel_dir))
