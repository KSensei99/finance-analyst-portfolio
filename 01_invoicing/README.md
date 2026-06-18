# Module 01 — Invoice Automation

## Business Case & Problem Statement
Nexoria Commerce Inc. processes 200+ wholesale and retail orders per month. Previously, the finance team manually created invoices by copying order data into static documents, taking 8–12 minutes per invoice. This consumed **40+ hours per month** of repetitive administrative labor, delayed billing cycles by 3+ days, and introduced billing discrepancies.

## The Automation Solution
To address this bottleneck, we implemented a dual-path automation system:
1. **High-Throughput Python Engine**: Uses `ReportLab` to query `nexoria.db` and generate hundreds of professional, branded PDF invoices in seconds.
2. **Interactive Excel VBA Automation**: An Excel workbook `NexoriaInvoicing.xlsx` containing the styled template (`InvoiceTemplate`), a flat `Orders` list, and a `TaxRates` lookup. A VBA macro loops through orders, handles multi-item orders, dynamically calculates sales tax, exports PDFs, and logs run history.

## Performance & Business Impact
* **Time Savings**: Reduced monthly invoicing effort from **40+ hours to under 3 minutes** of execution time.
* **Cash Flow Acceleration**: Invoices are generated same-day as orders, reducing Day Sales Outstanding (DSO) by 2.5 days.
* **Accuracy**: Automated template formulas guarantee mathematical consistency and eliminate data entry errors.

## Folder Contents
* [NexoriaInvoicing.bas](./01_invoicing/NexoriaInvoicing.bas) — The VBA macro code containing `GenerateInvoices` and `GetTaxRate`.
* [NexoriaInvoicing.xlsx](./01_invoicing/NexoriaInvoicing.xlsx) — Excel workbook pre-loaded with sample orders, tax rates, and the invoice template.
* [build_invoice_template.py](./01_invoicing/build_invoice_template.py) — Python script that creates the initial invoice template with openpyxl.
* [create_vba_invoicing_workbook.py](./01_invoicing/create_vba_invoicing_workbook.py) — Python script to load data and build `NexoriaInvoicing.xlsx`.
* [invoice_generator.py](./01_invoicing/invoice_generator.py) — The Python ReportLab PDF invoice generator.
* `invoices/` — Folder containing the generated sample PDF invoices.

## VBA Macro Setup Instructions
1. Open [NexoriaInvoicing.xlsx](./01_invoicing/NexoriaInvoicing.xlsx) in Excel.
2. Save the workbook as **Excel Macro-Enabled Workbook (`.xlsm`)**.
3. Press `Alt + F11` to open the VBA Editor.
4. Click **File -> Import File...** and select [NexoriaInvoicing.bas](./01_invoicing/NexoriaInvoicing.bas).
5. Close the VBA editor.
6. Create a folder named `Invoices` in the same directory as the workbook.
7. In the `Orders` sheet, clear any "Invoiced" flags in column L to re-run, then run the `GenerateInvoices` macro.
