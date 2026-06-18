# Module 05 — P&L Automation

## Business Case & Problem Statement
Nexoria Commerce Inc.'s senior financial analyst spent **2 full days each month** manually extracting trial balance data, mapping account lines, and building the monthly Profit & Loss (P&L) statement. This manual process led to periodic formula mismatches, inconsistent reporting formatting, and delayed distribution of the final financial packages to the CFO and executive team.

## The Automation Solution
We built an automated, interactive P&L reporting system:
1. **Automated ETL Pipeline**: [export_gl_data.py](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/05_pl_automation/export_gl_data.py) extracts actual and budget transactions from `nexoria.db` and outputs them as normalized, clean CSVs.
2. **Formula-Driven Report Engine**: [build_pl_report.py](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/05_pl_automation/build_pl_report.py) creates an Excel workbook containing raw data worksheets and a formatted `PL_Report` tab. The report uses Excel `SUMIFS` formulas linked to helper cells to dynamically calculate Revenue, COGS, Gross Profit, OpEx, and Operating Income (EBIT).
3. **Interactive Slicer Dropdown**: Cell `B3` contains an Excel Data Validation dropdown listing all available months (`2023-01` to `2024-12`). Selecting a month instantly recalculates the entire report.
4. **Conditional Formatting & Polish**: Favorable variances are highlighted in light green, while unfavorable variances are highlighted in light red. The final sheet columns are auto-fit using Excel's native layout engine via **Windows COM Automation (`win32com.client`)**.

## Performance & Business Impact
* **Reporting Cycle Time**: Preparation time reduced from **2 days to under 10 seconds** of script execution.
* **Calculation Integrity**: Zero hardcoded values; all math is governed by Excel formulas, guaranteeing internal alignment.
* **Proactive Oversight**: Color-coded variance alerts draw the CFO's attention to cost overruns immediately.

## Folder Contents
* [export_gl_data.py](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/05_pl_automation/export_gl_data.py) — Python script to extract GL data from SQLite.
* [build_pl_report.py](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/05_pl_automation/build_pl_report.py) — Python script to build the workbook, apply validation, and compile formatting via COM.
* [Nexoria_PL_Report.xlsx](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/05_pl_automation/Nexoria_PL_Report.xlsx) — The completed, formatted Excel P&L statement.
* `actuals.csv` / `budget.csv` — Extracted raw transactional data.

## Execution Instructions
To regenerate the report from the database:
```bash
# 1. Extract raw actuals and budget data from SQLite
python 05_pl_automation/export_gl_data.py

# 2. Rebuild the Excel report and run COM formatting
python 05_pl_automation/build_pl_report.py
```
*(Requires Microsoft Excel to be installed on the machine for the COM automation step to succeed.)*
