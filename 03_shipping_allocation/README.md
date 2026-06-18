# Module 03 — Shipping Cost Allocation Engine

## Business Case & Problem Statement
Nexoria Commerce Inc. incurs significant freight and courier shipping costs (FedEx, UPS, USPS). Carrier invoices arrive monthly as lump-sum totals (e.g., $15,000 to $30,000) per carrier, with no direct attribution to individual customer orders. Because of this, the finance team lumped all shipping expenses into a general overhead account, leaving gross margin reports **inaccurate by up to 19%** and making SKU-level profitability analysis impossible.

## The Automation Solution
We built a weighted cost allocation model in Python using `pandas`:
1. **Invoice Matching**: The script matches monthly carrier invoices to the list of orders shipped by that carrier during that month.
2. **Weighted Allocation Formula**: To accurately reflect the drivers of shipping costs, the engine distributes the carrier's lump-sum invoice back to individual orders using a weighted formula:
   * **70% based on shipment weight** (weight is the primary freight cost driver)
   * **30% based on order value** (captures insurance and value-based handling)
3. **Financial Output**: Generates `shipping_allocation_report.csv` (which feeds the Power BI dashboard) and a formatted pivot workbook `shipping_allocation_summary.xlsx` for finance reviews.

## Performance & Business Impact
* **Gross Margin Precision**: Improved gross margin reporting accuracy **from 78% to 97%+**.
* **SKU Profitability**: Finance can now identify negative-margin items and optimize pricing strategies.
* **Audit Trail**: Reconciles allocated costs back to carrier invoices with 100% mathematical coverage.

## Folder Contents
* [shipping_allocator.py](./03_shipping_allocation/shipping_allocator.py) — The Python Pandas script that performs the weighted cost allocation.
* [shipping_allocation_report.csv](./03_shipping_allocation/shipping_allocation_report.csv) — The allocated per-order shipping cost dataset.
* [shipping_allocation_summary.xlsx](./03_shipping_allocation/shipping_allocation_summary.xlsx) — Excel workbook containing pivot table summaries and gross margin analysis.
