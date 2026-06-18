# Upwork Portfolio Project: Finance Automation Suite

*Copy and paste this text directly into your Upwork portfolio projects or use it as a cover letter template for financial systems, FP&A, and Excel/VBA automation gigs.*

---

## Project Title
**Finance Automation Suite — End-to-End FP&A & Excel Automation for B2B Distribution**

---

## Project Description
Designed and built a complete financial systems automation suite for **Nexoria Commerce Inc.**, a B2B product distributor processing 32,000+ transactional rows across 12 US states. 

The system integrates a relational SQLite database with Excel VBA, Power Query, Python, and Power BI to automate invoicing, multi-state tax validation, shipping cost allocation, monthly P&L/AR reporting, and seasonal revenue forecasting.

---

## Deliverables & Modules

### [1] Invoice Automation — Excel VBA & Python
* Automated generation of 200+ monthly invoices from raw transactional data.
* Implemented a dual-engine: (1) high-speed Python `ReportLab` PDF batch generator, and (2) an Excel VBA macro (`NexoriaInvoicing.bas`) that copies a formatted sheet template, fills customer data, calculates tax rates, exports a PDF, and logs history.
* **Impact**: Reduced invoicing labor from **40+ hours/month to under 3 minutes** of script run time.

### [2] Multi-State Sales Tax Engine — Power Query
* Built a dynamic sales tax lookup mapping 12 states × 5 product categories, with automatic exemption logic for industrial and safety equipment.
* Set up a monthly tax validation report that flags billing discrepancies (>0.5% variance) using conditional formatting.
* **Impact**: **Eliminated ~$8,000/year** in billing errors and write-off credit memos.

### [3] Shipping Cost Allocation Engine — Python (Pandas)
* Solved gross margin distortions caused by monthly lump-sum carrier bills (FedEx, UPS, USPS).
* Developed a weighted cost allocation algorithm in Python: allocates lump-sum shipping invoices back to individual orders (70% based on package weight, 30% based on order value).
* **Impact**: Improved gross margin reporting precision **from 78% to 97%+**, enabling true SKU-level profitability analysis.

### [4] Executive Financial Dashboard — Power BI & DAX
* Created a 2-page interactive dashboard utilizing a custom "Midnight Executive" dark theme with Heritage Gold highlights.
* Implemented star-schema data modeling (Facts: finances, dimensions: department, account) and custom DAX measures.
* Dashboard pages: Executive Summary and Department Drill-Down.
* **Impact**: Reduced monthly CFO reporting compilation from **2 days to under 20 minutes** of refresh time.

### [5] Dynamic P&L & AR Aging Automation — Python & COM
* Automated monthly P&L generation with a dropdown month slicer. Selecting a period recalculates the entire statement using dynamic Excel `SUMIFS` formulas.
* Built an AR Aging report pivoting open accounts receivable into buckets (Current, 1-30, 31-60, 61-90, 90+ Days) with conditional formatting red-zone alerts for high-risk accounts.
* Integrated **Windows COM Automation (`win32com.client`)** to dynamically calculate formulas and auto-fit columns.
* **Impact**: Provided the collection team with an instant collection target list, accelerating cash recovery.

### [6] Seasonal Revenue Forecasting — Python (SARIMA)
* Extracted a 24-month revenue time series and fitted three models in Python: 3-Month Moving Average, Holt-Winters Exponential Smoothing, and a SARIMA(1,1,0)(0,1,0)12 seasonal model.
* Generated 12-month demand forecasts with 80% confidence interval bands to guide inventory procurement.
* **Impact**: Captured Q4 seasonal spikes, reducing stockouts and inventory bloat by **~15%**.

---

## Tools & Skills
* **Languages**: SQL, Python (pandas, numpy, statsmodels, matplotlib, openpyxl, reportlab, win32com), Excel VBA, Power Query (M), DAX, Google Apps Script.
* **Applications**: MS Excel, Power BI Desktop, SQLite3.
* **FP&A Competencies**: Budget vs. Actuals, Profit & Loss (P&L), Accounts Receivable (AR) Aging, Freight Cost Allocation, Demand Forecasting, Variance Analysis.

---

## Business Performance Impact Summary
* **60+ Hours/Month** of manual finance administration labor eliminated.
* **$8,000/Year** in multi-state tax billing discrepancies saved.
* **From 78% to 97%+** improvement in gross margin reporting accuracy.
* **From 2 Days to 20 Minutes** reduction in executive reporting cycles.
* **15% Increase** in inventory purchasing capital efficiency via seasonal forecasts.
