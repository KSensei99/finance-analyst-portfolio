import sqlite3
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from statsmodels.tsa.statespace.sarimax import SARIMAX
import warnings

# Suppress convergence and statsmodels warnings
warnings.filterwarnings("ignore")

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(base_dir, "..", "data", "nexoria.db")
    charts_dir = os.path.join(base_dir, "forecast_charts")
    os.makedirs(charts_dir, exist_ok=True)
    
    # 1. Load data from SQLite
    conn = sqlite3.connect(db_path)
    query = """
        SELECT 
            SUBSTR(order_date, 1, 7) as ym, 
            SUM(invoice_total) as revenue 
        FROM orders 
        WHERE status != 'Cancelled' 
        GROUP BY ym 
        ORDER BY ym;
    """
    df = pd.read_sql_query(query, conn)
    conn.close()

    # Format dataframe
    df['ym'] = pd.to_datetime(df['ym'] + '-01')
    df.set_index('ym', inplace=True)
    df.index.freq = 'MS' # Monthly Start frequency
    
    # 2. Forecasting Setup
    n_forecast = 12
    forecast_index = pd.date_range(start=df.index[-1] + pd.offsets.MonthBegin(1), periods=n_forecast, freq='MS')
    
    # --- Model 1: Moving Average (3-Month MA) ---
    ma_val = df['revenue'].rolling(window=3).mean().iloc[-1]
    forecast_ma = pd.Series([ma_val] * n_forecast, index=forecast_index)
    
    # --- Model 2: Holt-Winters Exponential Smoothing (Additive) ---
    # With 24 months, seasonal_periods=12 works
    hw_model = ExponentialSmoothing(
        df['revenue'], 
        trend='add', 
        seasonal='add', 
        seasonal_periods=12
    ).fit()
    forecast_hw = hw_model.forecast(n_forecast)
    
    # --- Model 3: Seasonal ARIMA (SARIMA) ---
    # Simple order due to limited dataset size: order=(1,1,0), seasonal_order=(0,1,0,12)
    # This captures the seasonal delta and first-order trend nicely
    sarima_model = SARIMAX(
        df['revenue'],
        order=(1,1,0),
        seasonal_order=(0,1,0,12),
        enforce_stationarity=False,
        enforce_invertibility=False
    ).fit(disp=False)
    
    sarima_res = sarima_model.get_forecast(steps=n_forecast)
    forecast_sarima = sarima_res.predicted_mean
    conf_int = sarima_res.conf_int(alpha=0.20) # 80% confidence interval for visualization
    
    # 3. Export CSV Results
    results_df = pd.DataFrame(index=pd.date_range(start=df.index[0], end=forecast_index[-1], freq='MS'))
    results_df = results_df.join(df.rename(columns={'revenue': 'Historic Revenue'}), how='left')
    
    # Create forecast series matching full results index
    ma_series = pd.Series(index=results_df.index)
    ma_series[forecast_index] = forecast_ma
    
    hw_series = pd.Series(index=results_df.index)
    hw_series[forecast_index] = forecast_hw
    
    sarima_series = pd.Series(index=results_df.index)
    sarima_series[forecast_index] = forecast_sarima
    
    ci_lower = pd.Series(index=results_df.index)
    ci_lower[forecast_index] = conf_int.iloc[:, 0]
    
    ci_upper = pd.Series(index=results_df.index)
    ci_upper[forecast_index] = conf_int.iloc[:, 1]
    
    results_df['MA Forecast'] = ma_series.round(2)
    results_df['HW Forecast'] = hw_series.round(2)
    results_df['SARIMA Forecast'] = sarima_series.round(2)
    results_df['SARIMA Lower CI'] = ci_lower.round(2)
    results_df['SARIMA Upper CI'] = ci_upper.round(2)
    
    csv_out_path = os.path.join(base_dir, "forecast_results.csv")
    results_df.to_csv(csv_out_path)
    print(f"Forecasting comparison CSV exported: {csv_out_path}")

    # 4. Generate Plot (Matplotlib)
    # Using premium design standards: dark theme background-like clean design
    plt.figure(figsize=(11, 6), dpi=150)
    plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
    
    # Historical Revenue
    plt.plot(df.index, df['revenue'], label='Historical Revenue', color='#1B3A6B', linewidth=2.5, marker='o', markersize=4)
    
    # Forecasts
    plt.plot(forecast_index, forecast_ma, label='3-Month Moving Average', color='#6B7280', linestyle='--', linewidth=1.8)
    plt.plot(forecast_index, forecast_hw, label='Holt-Winters Exp Smoothing', color='#D97706', linewidth=2.0, marker='^', markersize=4)
    plt.plot(forecast_index, forecast_sarima, label='SARIMA(1,1,0)(0,1,0)12', color='#2E86AB', linewidth=2.2, marker='s', markersize=4)
    
    # Confidence Interval
    plt.fill_between(
        forecast_index, 
        conf_int.iloc[:, 0], 
        conf_int.iloc[:, 1], 
        color='#2E86AB', 
        alpha=0.12, 
        label='SARIMA 80% Confidence Interval'
    )
    
    # Title & Labels
    plt.title('Nexoria Commerce Inc. — 12-Month Revenue Forecasting (2025)', fontsize=13, fontweight='bold', pad=15, color='#1B3A6B')
    plt.xlabel('Period', fontsize=10, fontweight='bold', labelpad=10)
    plt.ylabel('Monthly Revenue ($)', fontsize=10, fontweight='bold', labelpad=10)
    
    # Format axes
    plt.gca().yaxis.set_major_formatter(plt.FuncFormatter(lambda x, loc: "{:,}".format(int(x))))
    plt.xticks(pd.date_range(start=df.index[0], end=forecast_index[-1], freq='3MS'), rotation=45)
    plt.xlim(df.index[0] - pd.offsets.MonthBegin(1), forecast_index[-1] + pd.offsets.MonthBegin(1))
    
    # Premium Legend
    plt.legend(frameon=True, facecolor='white', edgecolor='none', shadow=True, loc='upper left', fontsize=9)
    plt.tight_layout()
    
    plot_out_path = os.path.join(charts_dir, "revenue_forecast.png")
    plt.savefig(plot_out_path)
    plt.close()
    print(f"Forecasting chart exported: {plot_out_path}")

if __name__ == "__main__":
    main()
