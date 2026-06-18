# Module 02 — Multi-State Sales Tax Engine

## Business Case & Problem Statement
Nexoria Commerce Inc. sells to customers across 12 US states. Since tax rates vary by state and product category (for example, safety gear and industrial equipment are exempt), manual tax determination was error-prone, resulting in **~$8,000/year** in credit memos and billing adjustments due to over/under-charging sales tax.

## The Automation Solution
We built a structured tax calculator and validation engine:
1. **Dynamic Power Query Lookup**: Power Query connects to the central `tax_rates.csv` table, formats the categories and rates, and loads them as a named table `TaxRates`. This ensures that rate changes update with a single click.
2. **Formula-Driven Exemption Logic**: The invoicing system automatically maps customer state and product category to determine taxability. Exempt categories receive a 0% rate automatically.
3. **Monthly Tax Validation Report**: An Excel report compares actual tax collected against expected values (subtotal × state rate), using conditional formatting to highlight any monthly variance over 0.5% in red for audit review.

## Performance & Business Impact
* **Error Reduction**: **Eliminated ~$8,000/year** in billing errors and subsequent accounting rework.
* **Audit Readiness**: Created a clean validation log that reduces tax audit prep time from 2 days to under 10 minutes.
* **Client Trust**: Stopped invoice variance issues, improving billing integrity.

## Folder Contents
* [sales_tax_queries.sql](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/02_sales_tax/sales_tax_queries.sql) — SQL queries used to aggregate sales and calculate tax liabilities.
* [generate_tax_report.py](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/02_sales_tax/generate_tax_report.py) — Python script that queries SQLite database and outputs the formatted validation report.
* [sales_tax_report.xlsx](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/02_sales_tax/sales_tax_report.xlsx) — Reconciled sales tax report with validation sheets and variance logs.
* [power_query_steps.md](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/02_sales_tax/power_query_steps.md) — Documentation showing how to recreate the Power Query extraction step-by-step.
