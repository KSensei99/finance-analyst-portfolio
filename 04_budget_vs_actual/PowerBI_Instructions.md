# Phase 4: Budget vs Actual Variance Dashboard (Power BI)

This guide provides the exact steps, DAX formulas, and data models required to build a highly professional, interactive Budget vs Actual Variance Dashboard in Power BI Desktop. This is a critical asset for a Finance Analyst portfolio.

## 1. Import Data (Get Data)
Open Power BI Desktop.
Click **Get Data -> Text/CSV** and import the following three files from the `04_budget_vs_actual` folder:
1. `fact_finances.csv` (Ensure `Date` is formatted as a Date type, `amount` as Decimal Number).
2. `dim_department.csv`
3. `dim_account.csv`

## 2. Create the Calendar Dimension (DAX)
To enable Time Intelligence functions (YTD, Prior Year), we need a centralized Date table.
1. Go to the **Modeling** tab and click **New Table**.
2. Paste the following DAX:
```dax
dim_date = 
ADDCOLUMNS (
    CALENDAR (DATE(2023, 1, 1), DATE(2025, 12, 31)),
    "Year", YEAR([Date]),
    "Month Number", MONTH([Date]),
    "Month Name", FORMAT([Date], "MMM"),
    "Quarter", "Q" & FORMAT([Date], "Q"),
    "Year Month", FORMAT([Date], "YYYY-MM")
)
```
3. Mark this table as a **Date Table** (Right-click `dim_date` -> Mark as date table -> select `Date`).

## 3. Build the Star Schema (Data Model)
Go to the **Model View** (the relationship icon on the left panel). Connect the tables using a 1-to-Many (1:*) relationship from Dimensions to the Fact table:
* `dim_date[Date]` -> `fact_finances[Date]`
* `dim_department[department]` -> `fact_finances[department]`
* `dim_account[gl_account]` -> `fact_finances[gl_account]`

## 4. Write DAX Measures
Create a dedicated "Measures" table (Home -> Enter Data -> Name it `_Measures`). Right-click inside this table to create the following New Measures:

### Core Amounts
```dax
Total Actuals = 
CALCULATE(
    SUM(fact_finances[amount]), 
    fact_finances[Scenario] = "Actual"
)
```

```dax
Total Budget = 
CALCULATE(
    SUM(fact_finances[amount]), 
    fact_finances[Scenario] = "Budget"
)
```

### Variances
```dax
Variance ($) = [Total Actuals] - [Total Budget]
```

```dax
Variance (%) = DIVIDE([Variance ($)], [Total Budget], 0)
```

*(Format Variance (%) as a Percentage, and Variance ($) as Currency).*

### Time Intelligence (Year-to-Date)
```dax
YTD Actuals = TOTALYTD([Total Actuals], dim_date[Date])
```

```dax
YTD Budget = TOTALYTD([Total Budget], dim_date[Date])
```

## 5. Build the Dashboard Pages

### Page 1: Executive Summary
- **KPI Cards (Top row):** Add 4 Cards showing `Total Actuals`, `Total Budget`, `Variance ($)`, and `Variance (%)`.
- **Waterfall Chart (Center):** 
  - *Category:* `dim_department[department]`
  - *Y-axis:* `Variance ($)`
  - This visual immediately shows which departments drove the variance.
- **Line & Clustered Column Chart (Bottom):**
  - *X-axis:* `dim_date[Year Month]`
  - *Column Y-axis:* `Total Actuals` and `Total Budget`
  - *Line Y-axis:* `Variance (%)`
- **Slicers (Left panel):** Add a dropdown slicer for `dim_date[Year]` and another for `dim_date[Month Name]`.

### Page 2: Department Drill-Down
- **Matrix Table:**
  - *Rows:* `dim_department[department]`, then nested under it: `dim_account[gl_category]`, then `dim_account[gl_account]`
  - *Values:* `Total Actuals`, `Total Budget`, `Variance ($)`, `Variance (%)`
- **Conditional Formatting:** Apply background color conditional formatting to the `Variance (%)` column in the Matrix. (Red = Over budget, Green = Under budget).

## 6. Publish to Portfolio
1. Save the file as `Nexoria_BVA_Dashboard.pbix` in the `04_budget_vs_actual` folder.
2. Publish to your Power BI Service workspace.
3. Generate a "Publish to Web" public link to include in your Upwork Portfolio Case Study.
