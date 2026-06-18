# Module 06 — Accounts Receivable (AR) Aging Report

## Business Case & Problem Statement
Nexoria Commerce Inc. issues invoices with payment terms such as NET30 or NET45. In practice, many wholesale customers delay payments, leading to a build-up of unpaid accounts receivable. Without a systematic tracking mechanism, the collection team was unable to identify overdue accounts, resulting in **$1M+ in outstanding receivables**, cash flow constraints, and increased write-off risk.

## The Automation Solution
We developed a dynamic AR Aging and collection reporting engine:
1. **Aging Bucket Calculator**: [ar_aging_report.py](./06_ar_aging/ar_aging_report.py) connects to the central database and pulls all unpaid invoices, calculating days outstanding relative to the current date and assigning them to standard accounting buckets:
   * **Current** (not yet due)
   * **1–30 Days Overdue**
   * **31–60 Days Overdue**
   * **61–90 Days Overdue**
   * **90+ Days Overdue** (critical collection risk)
2. **Accounts Receivable Summary Sheet**: Generates a consolidated customer-by-customer table of outstanding balances.
3. **Visual Collection Warning System**: Conditional formatting highlights overdue cells based on collection risk:
   * **90+ Days Overdue** is highlighted in **bold red** (critical cash-recovery focus).
   * **61–90 Days Overdue** is highlighted in **bold yellow** (pre-alert warnings).
   * **Current** is highlighted in **light green**.
4. **Executive Snapshot Cards**: Top of the summary sheet features high-level KPI cards for the CFO:
   * **Total AR Outstanding**
   * **Total Overdue Balance** (cumulative outstanding past due date)
   * **Collection Risk %** (percentage of AR that is 90+ days overdue)
5. **Audit Trail Detail Sheet**: A second sheet `Open_Invoices_Detail` lists every outstanding invoice with Order ID, customer, terms, amount, and exact days outstanding.

## Performance & Business Impact
* **Cash Recovery Prioritization**: The credit control team can instantly sort customers by **90+ Days** outstanding to target collections, accelerating cash receipts.
* **Bad Debt Provisioning**: Helps finance estimate write-offs and calculate the allowance for doubtful accounts.
* **Process Speed**: Generates a complete, audited monthly aging packet in **under 5 seconds**.

## Folder Contents
* [ar_aging_report.py](./06_ar_aging/ar_aging_report.py) — Python ETL and reporting script.
* [ar_aging_report.xlsx](./06_ar_aging/ar_aging_report.xlsx) — Output Excel report with conditional warnings and detailed sheets.

## Execution Instructions
To regenerate the AR aging report from the database:
```bash
python 06_ar_aging/ar_aging_report.py
```
*(Requires pywin32 to be installed on Windows to enable Excel COM automation for column sizing and calculation compilation.)*
