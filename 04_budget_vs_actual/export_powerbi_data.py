import sqlite3
import pandas as pd
import os

base_dir = os.path.dirname(os.path.abspath(__file__))
db_path = os.path.join(base_dir, "..", "data", "nexoria.db")
output_dir = base_dir

if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Read SQLite data
conn = sqlite3.connect(db_path)
budget = pd.read_sql_query("SELECT * FROM budget", conn)
actuals = pd.read_sql_query("SELECT * FROM actuals", conn)
conn.close()

# Add Scenario
budget['Scenario'] = 'Budget'
actuals['Scenario'] = 'Actual'

# Combine into Fact Table
fact_finances = pd.concat([budget, actuals], ignore_index=True)

# Create a proper Date column (First of the month) for Power BI time intelligence
fact_finances['Date'] = pd.to_datetime(fact_finances['year'].astype(str) + '-' + fact_finances['month'].astype(str) + '-01').dt.strftime('%Y-%m-%d')

# Clean up Fact Table
fact_finances = fact_finances[['Date', 'department', 'gl_account', 'Scenario', 'amount']]
fact_finances.to_csv(os.path.join(output_dir, 'fact_finances.csv'), index=False)

# Extract Dimension: Department
dim_department = pd.DataFrame(fact_finances['department'].unique(), columns=['department'])
dim_department['department_head'] = 'TBD' # Placeholder for extra dim attributes
dim_department.to_csv(os.path.join(output_dir, 'dim_department.csv'), index=False)

# Extract Dimension: Account
# Get unique combinations of gl_category and gl_account
dim_account = pd.concat([budget, actuals])[['gl_category', 'gl_account']].drop_duplicates()
dim_account.to_csv(os.path.join(output_dir, 'dim_account.csv'), index=False)

print(f"Exported Power BI Star Schema CSVs to {output_dir}")
