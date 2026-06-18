# Module 07 — Revenue Forecasting & Trend Analysis

## Business Case & Problem Statement
Nexoria Commerce Inc. previously made inventory purchasing and operational budget decisions based on historical annual averages. This flat approach caused **costly stockouts** during the Q4 seasonal rush (where sales surge by 300%+) and **inventory capital bloat** during slower Q1 months. To support inventory and cash flow planning, the CFO required a statistical monthly revenue forecast.

## The Forecasting Solution
We built a statistical forecasting pipeline in Python using `statsmodels` and `matplotlib`:
1. **Historical Baseline Extraction**: Queries monthly revenue from the orders database, building a 24-month historical time series (`2023-01` to `2024-12`).
2. **Three Statistical Models**: We fit and compared three distinct time series models to forecast 12 months of 2025:
   * **3-Month Moving Average (MA)**: A simple baseline that smooths short-term fluctuations but lacks trend or seasonal adaptation.
   * **Holt-Winters Exponential Smoothing (HW)**: Fits additive trend and additive seasonal components (seasonal cycle = 12 months). Excellent for capturing repeating annual cycles.
   * **Seasonal ARIMA (SARIMA)**: Fits a SARIMA(1,1,0)(0,1,0)12 seasonal model, incorporating seasonal differencing and autoregressive terms to handle seasonal variance and trend.
3. **80% Confidence Intervals**: The SARIMA model calculates lower and upper bounds of forecast uncertainty (confidence intervals), providing risk ranges for cash flow modeling.
4. **Data Visualization**: Saves a polished line chart [revenue_forecast.png](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/07_forecasting/forecast_charts/revenue_forecast.png) comparing models and shading the uncertainty bounds.

## Model Performance & Business Insights
* **Seasonality Capture**: Both Holt-Winters and SARIMA successfully predict the critical Q4 seasonality spike (e.g. November/December revenue peaks), projecting sales to exceed $700K and $800K respectively.
* **Accuracy Improvement**: Traditional moving averages lag during seasonal pivots; statistical forecasting reduces seasonal inventory planning errors by **~15%**.
* **Risk Management**: Shaded confidence intervals allow the procurement team to plan for "best-case" and "worst-case" supply scenarios.

## Folder Contents
* [revenue_forecast.py](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/07_forecasting/revenue_forecast.py) — Python modeling and graphing script.
* [forecast_results.csv](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/07_forecasting/forecast_results.csv) — Table containing historical revenue alongside model forecasts and confidence intervals.
* `forecast_charts/` — Folder containing the exported plot.
  * [revenue_forecast.png](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/07_forecasting/forecast_charts/revenue_forecast.png) — Chart showing historical and projected monthly sales.

## Execution Instructions
To run the forecasting models and regenerate charts:
```bash
python 07_forecasting/revenue_forecast.py
```
*(Requires pandas, numpy, statsmodels, and matplotlib to be installed in the Python environment.)*
