# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Master's (Cao học K34) final project for the Decision Support Systems course. The full spec is in `dac_ta_do_an_DSS.md`; re-read it before scoping work. The project has two parts on the Superstore sales dataset (`archive/train.csv`):

- **Part A – Data-Driven DSS (Power BI):** data cleaning, at least 5 KPIs (name + formula + managerial meaning), star schema with diagram, ETL into fact/dimension tables, dashboard, OLAP operations (slice/dice, drill-down, pivot), at least 3 insights and recommendations.
- **Part B – Model-Driven DSS (Excel):** at least 3 decision problems. Each needs decision variables, objective/outcome, formulas and constraints, and must use **What-if Data Table, Scenario Manager, Goal Seek and Solver**, followed by managerial interpretation, assumptions and limitations.

The tools are Power BI and Excel. Python (pandas) is only a helper for cleaning, ETL and generating the CSVs that feed those tools. Don't turn deliverables into notebooks.

Deliverables go in folder `MaNhom_MSHV1_MSHV2_MSHV3`: `MaNhom_BaoCao.pdf`, `MaNhom_Dashboard.pbix`, `MaNhom_Model.xlsx`, `MaNhom_TrinhBay.pptx`.

## Language

Write all report/slide prose, KPI descriptions, chart titles and code comments in **Vietnamese**. Keep code identifiers and column names in English.

## Data pipeline

- `data/raw/train.csv` is the immutable read-only original. Never edit it. `archive/train.csv` is the original download.
- `python -I etl/preprocess.py` produces `data/clean/Superstore_clean.csv` (ISO dates, UTF-8 BOM) and regenerates `docs/Nhat_ky_tien_xu_ly.md`, the report's preprocessing log. Don't hand-edit the log; change the script and rerun.
- `python -I etl/build_star_schema.py` (run after preprocess) splits the clean data into 6 star-schema CSVs in `data/star/` (Fact_Sales + Dim_Customer/Product/Geography/Date/ShipMode), runs 16 integrity checks and regenerates `docs/Kiem_tra_star_schema.md`. The diagram lives in `docs/Star_schema.md`; keep it in sync with table changes. Dim_Date has two relationships: OrderDateKey is active and ShipDateKey is inactive (use `USERELATIONSHIP`).
- KPI definitions, DAX measures and the dashboard layout live in `docs/KPI_va_DAX.md`. `python -I etl/kpi_reference.py` recomputes the same KPIs from `data/star/` into `docs/KPI_doi_chieu.md`, the reference numbers for checking Power BI. When a KPI definition changes, update both files together.
- `etl/preprocess_powerquery.m` mirrors the same steps for Power BI. Keep it in sync when cleaning logic changes.
- Policy: clean only real errors. Don't inject simulated errors into the dataset.

## Part B pipeline (Excel)

- The team builds the submitted `MaNhom_Model.xlsx` **by hand** following PHẦN B of `docs/Huong_dan_PowerBI.md`. That file is the single team guide for both Part A (Power BI) and Part B (Excel); `docs/Huong_dan_tu_lam_Excel.md` is only a pointer. They start from `output/PhanB_Mau_nhap_lieu.xlsx`, made by `python -I model/build_template.py`: the same layout with input data only, and no formulas, names or tool results. The generated `output/MaNhom_Model.xlsx` is only a reference answer for checking, not the submission. If the model changes, update the guide's cell formulas and expected values too.

- Assumptions and real 2018 inputs live in `model/params.py`, the single source for both the workbook and the scipy reference. Change assumptions there, never in the generated xlsx.
- Run order: `python -I model/solve_reference.py` → `python -I model/build_model.py` → `powershell -NoProfile -ExecutionPolicy Bypass -File model/apply_excel_tools.ps1` → `python -I model/verify_excel.py`.
  - `apply_excel_tools.ps1` drives real Excel over COM (Goal Seek, Data Table, Scenario Manager, Solver with reports). It needs Excel and the Solver add-in.
  - `build_model.py` overwrites `output/MaNhom_Model.xlsx`, so always rerun the PS script after it.
  - `verify_excel.py` checks Excel against scipy and regenerates `docs/PhanB_Ket_qua.md`. The narrative is the hand-written `docs/PhanB_Mo_hinh.md`; update its numbers when results change.
- PowerShell 5.1 gotchas in the COM script:
  - Keep the .ps1 ASCII-only; Vietnamese strings come from the JSON.
  - Return Ranges with a leading comma, or PowerShell unrolls them into single cells.
  - Wrap `if` results in `@()`, cast JSON numbers to `[double]` before passing them to COM, and don't use `$xl.Union`.
- Sheet names like `B1_Hoa_von` look like cell refs, so defined-name refs must quote them (`'B1_Hoa_von'!$B$6`). A cell value starting with `=` is written as a formula by openpyxl.

## Dataset gotchas (`archive/train.csv`, 9,800 rows, orders 2015-01-03 → 2018-12-30; latest ship date 2019-01-05)

- Dates are **DD/MM/YYYY**. In pandas use `format="%d/%m/%Y"`; in Power Query use "Using Locale" (e.g. English (United Kingdom)). Otherwise day and month get swapped.
- `Sales` is the **only** numeric measure. There is no Profit, Quantity, Discount or cost. Part B needs these, so any cost/margin/budget parameters must be **simulated assumptions**. Document each one (value, rationale) so the report can state it.
- `Product ID` is not a clean key: 32 IDs map to more than one `Product Name`, and 16 names map to more than one ID. Use a surrogate key (Product ID + Product Name) for `DimProduct`.
- There is 1 fully duplicated row (all columns except `Row ID`). Remove it and mention this in the cleaning section.
- `Postal Code` has 11 missing values (all Burlington, Vermont → filled with `05401`), and 429 values lost their leading zero in the CSV (`6824`). Read it as a string and left-pad to 5 digits.
- 220 `Product Name` values contain NBSP (U+00A0), and some contain double spaces. Power Query's `Text.Clean` does not remove NBSP.
- `Order ID` repeats across order lines. Count orders with distinct `Order ID`, not rows.
- `Country` has a single value (United States), so it's useless as a dimension attribute.
