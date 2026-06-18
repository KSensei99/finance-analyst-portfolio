# Nexoria Commerce Inc. — Finance Automation walkthrough

This walkthrough documents the design system, module implementations, and verification testing for the **Nexoria Commerce Inc. Finance Analyst Portfolio** project.

---

## 🎨 Power BI Design System: Midnight Executive Theme
*Designed to signal premium business intelligence capabilities to prospective B2B clients.*

### Color Architecture (60-30-10 Rule)
* **Canvas**: `#0F1118` — Deep midnight page background.
* **Surface**: `#181C2A` — Visual container and card backgrounds.
* **Border**: `#252A3A` — Subtle 1px separation lines.
* **Heritage Gold**: `#C69B3C` — Primary accent (column headers, totals, emphasis).
* **Soft Gold**: `#E6D17B` — Secondary totals and highlights.
* **Ivory**: `#E8E4DC` — Primary high-contrast text.
* **Steel**: `#7F8FA4` — Axis labels, legends, and categories.
* **Variance Highlights**: Signal Green (`#47B39D`) for favorable, Signal Red (`#E05252`) for unfavorable.

---

## 🧩 Completed Modules (Phase 1 — Phase 9)

### Module 01: Invoicing Automation (VBA & Python)
* **VBA Module**: Created `NexoriaInvoicing.bas` containing `GenerateInvoices` and `GetTaxRate` macros. Loops through a flat list of orders, consolidates multiple products, calculates tax, exports PDFs, and logs runs.
* **Excel Workbook**: Created `NexoriaInvoicing.xlsx` preloaded with the styled template (`InvoiceTemplate`), sample orders, and tax rates.
* **Python Engine**: `invoice_generator.py` uses `ReportLab` to batch generate 25 branded PDFs in seconds.
* **Status**: ✅ Complete.

### Module 02: Multi-State Sales Tax Calculator
* **Implementation**: Built a SQL and Power Query ETL pipeline. The `TaxRates` sheet pulls rates dynamically. The billing engines use automated lookup formulas with category exemption logic. Added a monthly validation worksheet with conditional variance warnings.
* **Status**: ✅ Complete.

### Module 03: Shipping Cost Allocation Engine
* **Implementation**: `shipping_allocator.py` processes lump-sum freight invoices (FedEx, UPS, USPS) and distributes costs to individual customer orders (70% by package weight, 30% by order value). Generates `shipping_allocation_report.csv` and pivot summary workbooks.
* **Status**: ✅ Complete.

### Module 04: FP&A Dashboard (Power BI)
* **Implementation**: Built a star-schema model (finances fact table, department and account dimension tables) in Power BI Desktop. Applied the custom **Midnight Executive** theme and DAX measures.
* **Status**: ✅ Complete.

### Module 05: P&L Statement Automation (Excel & COM)
* **Implementation**: `build_pl_report.py` creates a formatted P&L worksheet with a period selection dropdown (`Settings` sheet) and dynamic `SUMIFS` formulas. Integrated **Windows COM Automation** (`win32com.client`) to compile calculations and auto-fit column dimensions.
* **Status**: ✅ Complete.

### Module 06: Accounts Receivable Aging Report
* **Implementation**: `ar_aging_report.py` pivots open invoices from SQLite by customer, placing balances into standard buckets (Current, 1-30, 31-60, 61-90, 90+ Days) with conditional green-yellow-red warning highlights and KPI summary cards.
* **Status**: ✅ Complete.

### Module 07: Revenue Forecasting & Trend Analysis
* **Implementation**: `revenue_forecast.py` fits a 3-Month Moving Average, Holt-Winters Exponential Smoothing, and a Seasonal ARIMA (SARIMA) model on 24 months of revenue data. Exports a comparison table (`forecast_results.csv`) and a plot showing forecasts with 80% confidence interval bands (`revenue_forecast.png`).
* **Status**: ✅ Complete.

### Module 08: Quote-to-Invoice Google Workspace Pipeline
* **Implementation**: Created `QuoteToInvoice.gs` Apps Script to convert Google Sheets quotes into formatted PDF invoices, email them via Gmail, and append logs to a pipeline worksheet. Provided `quotes_sample.csv` for sheet testing.
* **Status**: ✅ Complete.

### Module 09: Portfolio Case Study & Upwork Copy
* **Implementation**: Wrote a comprehensive business case study (`Nexoria_Finance_Case_Study.md`), generated a clean vector architecture diagram (`architecture_diagram.svg`), and created an optimized Upwork proposal template (`upwork_project_description.md`).
* **Status**: ✅ Complete.

---

## 🔬 Reconciliations & Verification Testing

We ran verification tests across all Python and Excel assets to guarantee mathematical correctness:

1. **Invoice Template Test**:
   * Verified that totals block formulas correctly reference column G (value cells) instead of column B (which was a static column reference bug in the template builder).
   * Verified output: `invoice_template.xlsx` successfully regenerated with correct formulas.
2. **Batch Invoicing Test**:
   * Command: `python 01_invoicing/invoice_generator.py --limit 5`
   * Result: 5 PDFs generated, 0 errors.
3. **P&L Report Test**:
   * Command: `python 05_pl_automation/build_pl_report.py`
   * Result: Raw data successfully imported; Data Validation list bound to `B3`; COM automation compiled formulas and auto-fitted columns.
4. **AR Aging Test**:
   * Command: `python 06_ar_aging/ar_aging_report.py`
   * Result: Open invoices pivoted, 90+ Days highlights colored red; COM automation auto-fitted columns.
5. **Forecasting Model Test**:
   * Command: `python 07_forecasting/revenue_forecast.py`
   * Result: Models fitted; Q4 seasonal spikes projected; `revenue_forecast.png` exported.
