---
name: model-b-problem
description: Design one Part B (Excel Model-Driven DSS) decision problem for this project — decision variables, objective, formulas, constraints, Excel layout, and how What-if/Scenario/Goal Seek/Solver apply, plus managerial interpretation. Use when defining, building, or writing up a Part B problem.
---

Design or write up a Part B decision problem. Topic or problem idea: $ARGUMENTS (if empty, propose 3–4 candidates that fit the Superstore data and let the user pick).

1. **Ground it in the data.** Compute the real inputs from `data/clean/Superstore_clean.csv` (already cleaned) with pandas: sales by Sub-Category, Region, Segment, Ship Mode or month as relevant. Dates there are already ISO (`yyyy-mm-dd`) and the duplicate row is already removed; read `Postal Code` as str. Show the numbers you derive.
2. **List the simulated assumptions.** The dataset only has Sales, so state each parameter you invent (COGS %, shipping cost per Ship Mode, fixed cost, marketing elasticity, budget cap, capacity…) in a table with value, unit and rationale. Keep them plausible and few.
3. **Model components.** Write them in Vietnamese:
   - Biến quyết định: the cells the manager controls.
   - Biến kết quả/mục tiêu: what to maximize or minimize.
   - Công thức: explicit formulas, written as Excel formulas.
   - Ràng buộc: each with its numeric bound.
4. **Excel sheet layout.** Give a cell-by-cell plan for the sheet: input/parameter block, decision cells (highlighted), calculation block, objective cell. Use named ranges.
5. **Excel tools.** Say what each tool does for this problem. Across all Part B problems, every tool must be covered at least once:
   - What-if Data Table: which input(s) vary, and over what range.
   - Scenario Manager: Low/Base/High scenarios and their values.
   - Goal Seek: which target, solved by changing which cell.
   - Solver: objective, variable cells, constraints and solving method (Simplex LP / GRG Nonlinear). The model must be linear if Simplex is used.
6. **Managerial interpretation.** In Vietnamese: what decision the result supports, the expected benefit, and the limitations and assumptions.

If an `.xlsx` is being produced, follow the layout from step 4. Remind the user that Data Tables, Scenario Manager and Solver must be run inside Excel itself, and that Solver answer/sensitivity reports should be kept for the report.
