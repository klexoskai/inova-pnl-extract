# P&L extraction and new-format insertion

This document describes what the pipeline **reads** from old-format P&L workbooks and **writes** into the new-format master template (`iNova LOCAL P&L Project BLANK TEMPLATE_Master V2.1.xlsm`).

## Overview

```text
old_format/*.xlsm  ──preflight──►  pnl_readiness_report.csv   (optional; no _NPL)
              │
              ├──extract──►  extractions.csv          ──insert──►  *_NPL.xlsm
              │              extractions_proj.csv              (copy of V2.1 template)
              └── error_extraction.csv (failures)
```

| Stage | Script / entry point | Output |
|--------|----------------------|--------|
| Pre-flight | `scripts/report_pnl_readiness.py` | `pnl_readiness_report.csv` (`filename`, `can_process`, `notes`) |
| Extract (SKU) | `scripts/extract_old_format_pnl.py` | `extractions.csv` |
| Extract (project) | same | `extractions_proj.csv` |
| Insert | `scripts/apply_new_format_insertions.py` or `PandL_mastersku_extract.ipynb` | `<source_stem>_NPL.xlsm` |

Template path is configured in [`config/pnl_pipeline.json`](../config/pnl_pipeline.json) and resolved via `pnl_lib.paths.default_new_format_template_path()`.

Shared library: [`scripts/pnl_lib/`](../scripts/pnl_lib/).

---

## Part 1 — Extractions (old format → CSV)

All extraction reads the **Project Summary** sheet (or the first matching alias) unless noted otherwise.

### 1.1 SKU-level extraction (`extractions.csv`)

**Module:** `pnl_lib.extract_sku`  
**One row per** product SKU × market (only rows with non-zero list price).

#### Step A — List price table (defines SKU × market rows)

| Source | Cell range | Notes |
|--------|------------|--------|
| Project Summary | `C3:AI14` | Header row; melt to long format |

Processing:

1. Drop rows whose list-price column contains `"input"`.
2. Drop all-zero / empty columns.
3. Drop first data row (template header row under headers).
4. Melt: id = list price column, `Market` = column headers, `list_price_AUD` = values.
5. Build `sku_name` from SKU label; `sku_market` = `{sku}_{Market}`.
6. Keep rows where `list_price_AUD` is non-null and non-zero.

#### Step B — Feature tables (joined by `sku_name` + `Market`)

Each field is read from a fixed rectangle on Project Summary and looked up by `(sku_name, Market)`:

| CSV column | Source range (Project Summary) |
|------------|--------------------------------|
| `forecast_volume_y1` | `BW2:DC15` |
| `forecast_volume_y2` | `BW17:DC30` |
| `forecast_volume_y3` | `BW32:DC45` |
| `forecast_volume_y4` | `BW47:DC58` |
| `forecast_volume_y5` | `BW62:DC73` |
| `GTN_perc excluding Launch yr one-time costs` | `C16:AI27` |
| `COGS/unit` | `C69:AI80` |
| `Launch Yr one-time cost (Listing fees, launch COOP)` | `C30:AI41` |
| `launch_year_Q` | `C155:AI165` |
| `forecast_net_sales_y1` … `y5` | `HD2:IJ15`, `HD17:IJ30`, `HD32:IJ45`, `HD47:IJ58`, `HD62:IJ73` |
| `launchyr_coop_listingfees_y1` … `y3` | `FV2:HB15`, `FV17:HB30`, `FV32:HB45` |
| `other_COGS_y1` … `y5` | `JT2:KZ13`, `JT17:KZ28`, … `JT62:KZ73` |
| `one_time_costs_y1` … `y5` | `LC2:MI13`, `LC17:MI28`, … `LC62:MI73` |
| `GP_y1` … `y5` | `ML2:NP13`, … `ML62:NP73` |
| `forecast_gross_sales_y1` … `y5` | `DF2:EK15`, … `DF62:EK73` |

#### Step C — Metadata columns (prepended)

| Column | Source |
|--------|--------|
| `filename` | Workbook basename |
| `sheet_name` | Resolved sheet name |
| `file_year` | Parsed from path/filename |
| `project_name` | Parsed from path/filename |
| `gate_number` | Parsed from path/filename |

#### Step D — Post-processing (`preproc_sku_extractions`)

Derived column used by insertion:

| Column | Formula |
|--------|---------|
| `COGS_other_combined_y1` … `y5` | `one_time_costs_y{n}` + `other_COGS_y{n}` (missing → 0) |

#### Columns extracted but not inserted by default

These are present in `extractions.csv` for analysis or future insertions: net sales, gross sales, GP, launchyr coop/listing fees, etc.

#### `launch_year_Q` (insertion dependency)

Extracted from **Project Summary** `C155:AI165` (per SKU × market). Required for **Actual Launch Year (K)** and **Actual Launch Month (L)** on sku-block sheets (including **VOL_Auto**). Example values: `2026 Q1`, `2025 Q3`. Parsed as:

- **K** = 4-digit year  
- **L** = first month of quarter (Q1→January, Q2→April, Q3→July, Q4→October)

#### Product SKU column for insertion

Use **`sku_name.1`** when pandas suffixes a duplicate `sku_name` column (composite `sku_name` vs product name). Insertion groups by product SKU (`Blackcurrant 16`), not `sku_market` (`Blackcurrant 16_Australia`).

---

### 1.2 Project-level extraction (`extractions_proj.csv`)

**Module:** `pnl_lib.extract_proj`  
**One row per** market with active list-price columns (non-zero in `C3:AI14`).

Aggregation columns are excluded: `CANZ`, `Asia Direct`, `Asia Total`, `Europe Total`, `ME total`, `AMENA`, `Total iNova`.

| CSV column pattern | Source range | Notes |
|--------------------|--------------|--------|
| `Market` | From list-price active columns | |
| `ANP_perc of Net Sales_Year 1` … `Year 5` | `C134:AI140` | Pivoted from row labels in col A |
| `Incr Opex (Submissions, FTE etc.)_Year 1` … | `C147:AI152` | Merged on Market |
| `CAPEX_*` | `C134:AI145` rows from row 6 onward | Pivoted CAPEX block; CSV column is often **`CAPEX_CAPEX $`** (depends on Excel row labels in col A) |
| `canni_perc_y1` … `canni_perc_y5` | **D274:AI274** (market labels) + **D276:AI276** … **D280:AI280** (value rows) | One shared header row for all years |
| `canni_GP_perc` | **D274:AI274** (market labels) + **D282:AI282** (values) | Same market header row as years |

| Column | Source |
|--------|--------|
| `filename` | Workbook basename |

---

### 1.3 CLI — extract

```bash
python scripts/extract_old_format_pnl.py \
  --input-dir /path/to/pnl-data/old_format \
  --output-dir /path/to/inova-pnl-extract
```

Writes `extractions.csv`, `extractions_proj.csv`, and `error_extraction.csv` (columns: `filename`, `file_path`, `stage`, `reason`).

---

## Part 2 — Insertions (CSV → new-format template)

**Module:** `pnl_lib.insert_template`  
**Input:** open workbook (copy of master template); filtered rows where `extractions.filename` matches the source file being built.

Output file name: `<original_stem>_NPL.xlsm` (e.g. `Project.xlsm` → `Project_NPL.xlsm`).

### 2.1 Layout convention — sku-block sheets

Sheets: **SP**, **VOL_Auto**, **SALES_DED_TD**, **COGS**, **COGS_OTHER**, **SALES_DED_KA**, plus **A&P ALLOC** (column **X** only uses this layout).

| Concept | Value |
|---------|--------|
| First SKU block start row | 8 |
| Row step per SKU | 15 (blocks at 8, 23, 38, …) |
| Last row | 143 |
| Max SKU blocks | 10 |
| Market column | **C** |
| Rows within block | One row per market row in CSV for that product SKU |

Before writing, market and value cells in the block range are cleared (sku-block tabs clear **C** plus their value columns; **A&P ALLOC** clears column **X** rows 8–143 only).

### 2.2 Insertion map (implemented)

| Sheet | Source column(s) | Target cell(s) | Notes |
|-------|------------------|----------------|--------|
| **Home Tab** | (fixed) | **C6** | Base currency set to **AUD** (replaces existing cell value) |
| **Home Tab** | Unique `sku_name.1` | **J6** downward | Up to 10 SKUs |
| **SP** | `list_price_AUD` | **AI** | Per market row |
| **VOL_Auto** | `forecast_volume_y1` … `y5` | **AQ:AU** | One column per year |
| **SALES_DED_TD** | `GTN_perc excluding Launch yr one-time costs` | **BA, BN, CA, CN, DA, DN** | Same value repeated on each row |
| **COGS** | `COGS/unit` | **AK** | Per market row at sku-block anchors (e.g. **AK8**, **AK9**) |
| **COGS_OTHER** | `COGS_other_combined_y1` … `y5` | **AK:AO** | Requires `preproc_sku_extractions` |
| **SALES_DED_KA** | `Launch Yr one-time cost (Listing fees, launch COOP)` | **AK** | |
| **A&P ALLOC** | Unique `Market` per product SKU (`sku_name.1`), first-seen order within SKU | **X8**, **X9**, … then **X23**, **X24**, … | Same block geometry as column **C** on other sku sheets; **only column X** touched |
| All sku-block sheets above | `launch_year_Q` | **K** (year), **L** (month name) | Only if column exists in CSV |
| **Home Tab** | Unique `Market` from `extractions_proj` (first appearance in CSV) | **I19**, **I20**, … | Clears **I19:I98** then writes; omits blank/NaN; only when project CSV has ANP rows |
| **A&P TABLE** | `Market`, `ANP_perc of Net Sales_Year 1` … `Year 5` | **B4:G…** | From `extractions_proj`; one row per market |
| **CAPEX** | `CAPEX_CAPEX $` | **AK8**, **AK9**, … | From `extractions_proj`; **consecutive** rows (not sku-block stepped); clears **AK8:AK127** first |
| **CANNIBAL** | `canni_perc_y1` … `y5` | **BA/BB/BC/BD/BE** at sku blocks | Year columns that feed **CANNIBAL_REV** `GA:GE` |
| **CANNIBAL_REV** | `canni_GP_perc` | **AI** (same row as market on **CANNIBAL**) | GP multiplier in `GA8`…`GE8` formulas (`*$AI8`) |

**Not inserted today:** Incr Opex columns (present in `extractions_proj` but not written to template).

### 2.3 Pre-flight report (no _NPL output)

```bash
python scripts/report_pnl_readiness.py \
  --input-dir /path/to/pnl-data/old_format \
  --output pnl_readiness_report.csv
```

Writes **`pnl_readiness_report.csv`**: `filename`, `can_process` (`yes`/`no`), `notes` (automated issues/warnings for finance to extend). Runs SKU + project extraction and, by default, an in-memory insertion dry-run against the master template. Use `--extract-only` for a faster extract-only check.

### 2.4 CLI — insert

```bash
python scripts/apply_new_format_insertions.py \
  --extractions extractions.csv \
  --extractions-proj extractions_proj.csv \
  --output-dir /path/to/pnl-data/new_format
```

For each unique `filename` in `extractions.csv`, copies the master template and writes `<filename_stem>_NPL.xlsm`. Failures → `error_insertion.csv`.

### 2.4 Notebook

`PandL_mastersku_extract.ipynb` — insertion cell loads CSVs, calls `apply_insertions_to_workbook`, saves `_NPL` next to the template. **Filter `extractions` by `filename`** when multiple source files are in one CSV.

---

## Part 3 — Declarative spec (reference)

[`config/pnl_template_insertions.yaml`](../config/pnl_template_insertions.yaml) mirrors the insertion map above for a future YAML-driven runner. The live implementation is in `insert_template.py`.

---

## Part 4 — File reference

| File | Role |
|------|------|
| `scripts/extract_old_format_pnl.py` | Batch extract CLI |
| `scripts/apply_new_format_insertions.py` | Batch insert CLI |
| `scripts/pnl_lib/extract_sku.py` | SKU extraction |
| `scripts/pnl_lib/extract_proj.py` | Project/market extraction |
| `scripts/pnl_lib/insert_template.py` | Template writes |
| `scripts/pnl_lib/paths.py` | Template/output paths |
| `scripts/pnl_lib/excel_utils.py` | Excel read helpers |
| `scripts/pnl_lib/errors.py` | Error CSV helpers |
| `config/pnl_pipeline.json` | Master template filename + directory |
| `config/pnl_extract.json` | Legacy FastAPI/notebook extract config (feature table list) |
