import sqlite3
import os
import csv

def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(base_dir, "..", "data", "nexoria.db")
    
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # Extract actuals
    cur.execute("SELECT year, month, department, gl_category, gl_account, amount FROM actuals ORDER BY year, month, gl_category, gl_account;")
    actuals = cur.fetchall()
    actuals_csv_path = os.path.join(base_dir, "actuals.csv")
    with open(actuals_csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["year", "month", "department", "gl_category", "gl_account", "amount"])
        writer.writerows(actuals)
    print(f"Extracted {len(actuals)} actuals rows to {actuals_csv_path}")

    # Extract budget
    cur.execute("SELECT year, month, department, gl_category, gl_account, amount FROM budget ORDER BY year, month, gl_category, gl_account;")
    budget = cur.fetchall()
    budget_csv_path = os.path.join(base_dir, "budget.csv")
    with open(budget_csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["year", "month", "department", "gl_category", "gl_account", "amount"])
        writer.writerows(budget)
    print(f"Extracted {len(budget)} budget rows to {budget_csv_path}")

    conn.close()

if __name__ == "__main__":
    main()
