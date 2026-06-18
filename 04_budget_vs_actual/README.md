# Module 04 — Budget vs. Actual FP&A Dashboard

## Business Case & Problem Statement
Nexoria Commerce Inc.'s CFO and executive team lacked a consolidated, interactive tool to monitor performance against budget, identify OpEx variances, and track multi-state sales tax liabilities. Preparing monthly slides took the senior analyst **2 full days** of manual spreadsheet consolidation, resulting in lagged decision-making.

## The Automation Solution
We built a premium, star-schema Power BI report that links all data sources to a central reporting layer:
1. **Star-Schema Data Model**: Connects the `fact_finances` table (containing GL actuals and budgets) to dimension tables: `dim_department` and `dim_account`.
2. **Midnight Executive Theme**: Customized using a premium dark theme palette (Midnight Navy, Slate Charcoal) with Heritage Gold accents for an executive, modern presentation.
3. **DAX Financial Measures**:
   * `Total Actuals = SUM(fact_finances[actual])`
   * `Total Budget = SUM(fact_finances[budget])`
   * `Variance ($) = [Total Actuals] - [Total Budget]`
   * `Variance (%) = DIVIDE([Variance ($)], [Total Budget], 0)`
4. **Conditional Formatting**: Variance table dynamically highlights under-budget (green) and over-budget (red) items.

## Performance & Business Impact
* **Reporting Efficiency**: Reduced monthly presentation preparation time from **2 days to under 20 minutes** (run/refresh time).
* **Granular Visibility**: Enabled instant drill-down from department level to specific General Ledger (GL) accounts.
* **Proactive Variance Control**: Under-budget and over-budget highlights allow the finance team to isolate cost overruns within seconds.

## Folder Contents
* [Nexoria_BVA_Dashboard.pdf](./04_budget_vs_actual/Nexoria_BVA_Dashboard.pdf) — Exported PDF version of the 2-page dashboard layout.
* [Nexoria_BVA_Dashboard.pbip](./04_budget_vs_actual/Nexoria_BVA_Dashboard.pbip) — Power BI Project report (includes `.Report` and `.SemanticModel` folders for developer Git versioning).
* [POWERBI_LIVE_LINK.txt](./04_budget_vs_actual/POWERBI_LIVE_LINK.txt) — Placeholder file for the published dashboard link.
* [PowerBI_Instructions.md](./04_budget_vs_actual/PowerBI_Instructions.md) — Setup and build documentation for replicating the report.
* [export_powerbi_data.py](./04_budget_vs_actual/export_powerbi_data.py) — Python script that extracts the clean CSV dimensions and facts from SQLite.
* `dim_account.csv` — Dim table containing GL account codes and groupings.
* `dim_department.csv` — Dim table containing department mapping.
* `fact_finances.csv` — Fact table containing monthly actual and budget amounts.
