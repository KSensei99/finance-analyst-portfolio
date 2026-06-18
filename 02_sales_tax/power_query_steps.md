# Power Query Steps: Sales Tax Reconciliation

In a real-world scenario, you might not use Python to build your Excel dashboard. Instead, you would connect Excel directly to the database using Power Query. This ensures your dashboard automatically updates when new data is added to the database.

Here is the exact step-by-step process you can use to build this directly in Excel, which you can mention in the Upwork case study.

## 1. Connect to the Database
1. Open a blank Excel workbook.
2. Go to **Data** > **Get Data** > **From Database** > **From ODBC** (if you have SQLite ODBC installed) OR export the SQLite tables to CSVs and use **From Text/CSV**.
3. Select `orders` and `state_tax_rates`.

## 2. Merge Data in Power Query
1. In the Power Query Editor, select the `orders` query.
2. Go to **Home** > **Merge Queries**.
3. Merge `orders` with `state_tax_rates` using the `state` column from `orders` and `state_code` from `state_tax_rates`.
4. Expand the merged table to include `state_tax_rate` and `nexus`.

## 3. Add Custom Columns for Calculations
We need to calculate the *expected* tax and the *discrepancy*.
1. Go to **Add Column** > **Custom Column**.
2. **Name:** `Expected_Tax`
   **Formula:** `[taxable_amount] * [state_tax_rate] * [nexus]`
3. **Name:** `Discrepancy`
   **Formula:** `[sales_tax] - [Expected_Tax]`
4. Format both new columns as Currency.

## 4. Load to Data Model and Pivot
1. Go to **Home** > **Close & Load To...**
2. Select **PivotTable Report** (or **Only Create Connection** and check "Add this data to the Data Model" if you want to use Power Pivot).
3. In the PivotTable Field List:
   - **Rows:** `state`
   - **Values:** `Sum of taxable_amount`, `Sum of Expected_Tax`, `Sum of sales_tax`, `Sum of Discrepancy`.

## 5. Formatting the Dashboard
1. Apply the standard financial color coding (Blue for inputs, Black for formulas/data).
2. Use Number Formatting `_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)` to make zeros appear as dashes.
3. Add Conditional Formatting to the `Discrepancy` column:
   - Highlight Red if `> 0.05` or `< -0.05` (indicating an over/under collection error).
