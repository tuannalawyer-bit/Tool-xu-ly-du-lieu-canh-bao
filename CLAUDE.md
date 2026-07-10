# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

A Vietnamese-language desktop tool (Tkinter GUI) for analyzing SAP inventory reports and flagging stock issues (expired, slow-moving, non-moving, "phantom stock"). It runs on a bundled, offline Python (no internet/pip access expected at runtime) and produces Excel + interactive HTML dashboard outputs.

There is no build system, package manager, or test suite — this is a single-operator internal tool, run by double-clicking a `.bat` file or executing the `.py` script directly with Python 3.

## Running the tool

- **Primary entry point:** `tool_do_du_lieu.py` (v1.5.2) — the current, integrated tool with a two-tab Tkinter UI (Tab 1: "Phân Tích Chi Tiết", Tab 2: "Tổng Quan Chuỗi"). This supersedes the older split tools.
- Launch via `Chay Tool Do Du Lieu.bat` (root) or `Tool_PhanTichTonKho_Portable/Chay_Tool.bat` (portable copy) — these locate a bundled `python_embed/python.exe` and an offline `libs/` folder (openpyxl, pyxlsb, et_xmlfile) and fall back to system `python` if the embedded interpreter is absent.
- To run directly: `python tool_do_du_lieu.py` from the repo root (needs `openpyxl` installed, plus `pyxlsb` if `.xlsb` inputs are used).
- `tool_chi_tiet.py` and `tool_tong_quan.py` (root) are the **legacy pre-merge** versions of Tab 1 and Tab 2 respectively — kept for reference only; new work should go into `tool_do_du_lieu.py`.
- `tool_do_du_lieu_v1.5.1_backup.py` is a manual backup snapshot, not imported by anything.
- No automated tests, linter, or build step exists in this repo.

## Project structure

```
tool_do_du_lieu.py                    # Main integrated tool (current)
tool_chi_tiet.py / tool_tong_quan.py  # Legacy standalone predecessors (Tab 1 / Tab 2)
dashboard_chi_tiet_template.html      # HTML/JS template injected with Tab 1 output data
dashboard_tong_quan_template.html     # HTML/JS template injected with Tab 2 output data
development_log.csv                   # Manually maintained changelog of features/fixes
Tool_PhanTichTonKho_Portable/         # Self-contained portable bundle (copy of the .py + templates
                                       # + python_embed/ + libs/ + reference lookup files), built for
                                       # distribution to non-dev users; not gitignored uniformly (see below)
raw mẫu/                              # Sample/reference input files (gitignored)
```

`raw mẫu/`, `Tool_PhanTichTonKho_Portable/python_embed/`, `Tool_PhanTichTonKho_Portable/raw mẫu/`, `*.zip`, and generated `* (done).xlsx` outputs are gitignored. `*.xlsx` / `*.xlsb` are tracked via Git LFS (`.gitattributes`).

## Architecture of `tool_do_du_lieu.py`

The file is organized top-to-bottom in clearly delimited sections (search for `# ====...====` banners):

1. **Config & offline libs setup** — resolves `BASE_DIR` (handles both script and PyInstaller-frozen `sys.frozen` execution), inserts `libs/` folders into `sys.path` before importing `openpyxl`, and locates default reference files (`Master article 7.7.xlsx`, `Store List.xlsx`) via `find_default_file()`, which searches several candidate relative paths including a `../Last GR.SALE/` sibling folder.
2. **Utility helpers** — Excel column letter <-> index conversion, `safe_float`, `excel_date_to_datetime`.
3. **Reference data loaders**:
   - `load_store_details()` — parses `Store List.xlsx` into a store lookup map.
   - `load_product_details()` — parses `Master article 7.7.xlsx` into a product/shelf-life lookup map. This file is large (~4.6 MB), so results are **cached to a pickle** (`.Master article 7.7_cache.pkl`, next to the source file) keyed/validated by `_is_cache_valid()` against the source file's mtime/hash — cache load is <1s vs. 2-4 minutes for a cold parse. Pass `force_reload=True` to bypass.
4. **Row-level business logic** — `compute_row()` is the core classification engine. It works on lettered SAP columns (`K`, `N`..`X`) and derives columns `Y` through `AF`:
   - `Y` = days since last goods receipt (GR), capped/labeled `'>90'`.
   - `Z` = shelf life (from `shelf_life_map`, else `'#N/A'` propagates through).
   - `AA` = projected days of stock at current 15-day sales velocity.
   - `AB` = `AA / Z` (stock coverage ratio vs shelf life).
   - `AC` = note flags: `"không giao dịch >90 ngày"` (no GR and no sale >90d) or `"Không sale 90 ngày"` (no sale but has GR).
   - `AD` = `"Check"` if `AB > 5` and stock value exceeds the reference amount threshold.
   - `AE` = `"Hết hạn"` (expired) if last-GR age exceeds shelf life.
   - `AF` = final classification, priority order: `Hết hạn` (expired) > `Nghi vấn tồn ảo` (phantom stock) > `Non-moving` > `Slow moving` > none.
   - When modifying classification rules, this function is the single source of truth — don't duplicate the logic elsewhere.
5. **File processors** — `process_csv_file()` and `process_excel_file()` read a raw SAP export (`.csv` or `.xlsx`, auto-routed by extension), apply `compute_row()`-equivalent logic per row, and write a `*(done).xlsx` output (27/31 columns) using `openpyxl` in write-only/streaming mode (`WriteOnlyCell`) for memory efficiency on large files. `process_excel_file()` opens the source once via `openpyxl.load_workbook(read_only=True)` and locates the header row by scanning the first rows for a literal `'RSM'` cell — do not reintroduce a separate `zipfile`/`ElementTree` pre-pass to "find the header row"; that was removed because it re-parsed the whole file for no benefit (see `development_log.csv` / git history for the perf fix). `shelf_life_map` (derived from `product_map`) is computed once per batch run and passed into these functions rather than recomputed per file.
6. **Dashboard generation**:
   - `generate_html_dashboard()` (Tab 1) — injects per-row detail data into `dashboard_chi_tiet_template.html`, producing an interactive chart + pivot-table HTML per processed file.
   - `read_and_aggregate_files()` + `generate_overview_html()` (Tab 2) — reads multiple `*(done).xlsx` outputs and aggregates them by `[RSM (Vùng), ASM, Cửa hàng, MCH2 (Ngành hàng), Phân loại kiểm tra]`, dropping per-product detail rows entirely to keep the merged dashboard small (goal: 80MB+ raw -> under 300KB HTML). Chart slicer behavior is hierarchical: default view groups by RSM; drilling into an RSM groups by ASM; drilling into an ASM breaks down by individual Store Code.
7. **`App(tk.Tk)`** — the GUI. `_build_ui()` sets up a `ttk.Notebook` with two tabs; `_build_tab1_ui()` / `_build_tab2_ui()` build each tab's widgets. Each tab keeps its own file list, logging (`log_t1`/`log_t2`), and run handler (`_on_run_t1`/`_on_run_t2`), which spawn a background `Thread` to keep the UI responsive during processing.

## Key domain concepts

- **Milestone date (Y1) / reference amount (AF1)**: user-supplied parameters that anchor the "days since" calculations and the stock-value threshold for flagging.
- **Chuỗi báo cáo (report chain)**: a named retail chain/brand (e.g. Urban, WIN) selected by the user, used as the dashboard title — not derived from the data.
- **`*(done).xlsx`**: the canonical intermediate artifact — Tab 1 output, Tab 2 input. Any change to `compute_row()`'s output columns must stay compatible with what `read_and_aggregate_files()` expects to find.
