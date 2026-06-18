# Nexoria Finance Analyst Portfolio — Task Tracker

## Pre-Phase 5 Improvements
- [x] Run Graphify Knowledge Graph updates for skills
- [x] Add VBA Macro and xlsm workbook to `01_invoicing/` alongside Python PDF generator
- [x] Create per-module `README.md` files for:
  - [x] `01_invoicing/`
  - [x] `02_sales_tax/`
  - [x] `03_shipping_allocation/`
  - [x] `04_budget_vs_actual/`
- [x] Polish root `README.md` (remove false complete marks, add client scenario, business narrative, and quantified impact stats)

## Phase 5 — P&L Automation
- [x] Write GL data extractor from SQLite
- [x] Write Python openpyxl + COM automation script to build formatted P&L workbook
- [x] Implement dynamic month selector via COM automation
- [x] Create `05_pl_automation/README.md` with business context

## Phase 6 — AR Aging Report
- [x] Write Python script to calculate aging buckets from SQLite (`ar_open_invoices`)
- [x] Generate formatted Excel workbook with conditional formatting (green-yellow-red)
- [x] Create `06_ar_aging/README.md` with business context

## Phase 7 — Revenue Forecasting
- [x] Write monthly revenue time series extraction from orders
- [x] Implement Moving Average, Exponential Smoothing, and SARIMA forecasting models
- [x] Export forecast charts and model comparison tables
- [x] Create `07_forecasting/README.md` with business context

## Phase 8 — Case Study & Upwork Packaging
- [x] Write full case study document in `case_study/`
- [x] Generate clean architecture diagram (SVG/PNG)
- [x] Write copy-paste-ready Upwork project description file
- [x] Finalize root `README.md` polish

## Phase 9 — Google Apps Script Quote-to-Invoice
- [x] Create `08_quotes_pipeline/` directory
- [x] Write Apps Script `QuoteToInvoice.gs` with quote conversion, PDF generation, email sending, and pipeline logging
- [x] Write instructions in `08_quotes_pipeline/README.md`
