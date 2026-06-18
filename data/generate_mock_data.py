"""
=============================================================
  Nexoria Commerce Inc. — Mock Data Generator
  Finance Analyst Portfolio Project | Phase 0
=============================================================
  Generates realistic mock financial data and loads it into
  a SQLite database (nexoria.db) for use across all modules.

  Run this script ONCE from the /data directory.
  Requirements: pip install pandas numpy faker openpyxl
=============================================================
"""

import sqlite3
import os
import random
import string
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
from faker import Faker

# ─────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────
random.seed(42)
np.random.seed(42)
fake = Faker("en_US")
Faker.seed(42)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "nexoria.db")
PROCESSED_DIR = os.path.join(BASE_DIR, "processed")
os.makedirs(PROCESSED_DIR, exist_ok=True)

# Date range: 2 full years of data
START_DATE = datetime(2023, 1, 1)
END_DATE = datetime(2024, 12, 31)

N_CUSTOMERS = 500
N_PRODUCTS = 80
N_ORDERS = 3500
N_CARRIERS = 5

print("=" * 60)
print("  Nexoria Commerce Inc. — Data Generator")
print("=" * 60)

# ─────────────────────────────────────────────────────────
# 1. STATE TAX RATES (All 50 US States)
# ─────────────────────────────────────────────────────────
print("\n[1/9] Generating state tax rates...")

STATE_TAX_RATES = {
    "AL": 0.04, "AK": 0.00, "AZ": 0.056, "AR": 0.065, "CA": 0.0725,
    "CO": 0.029, "CT": 0.0635, "DE": 0.00, "FL": 0.06, "GA": 0.04,
    "HI": 0.04, "ID": 0.06, "IL": 0.0625, "IN": 0.07, "IA": 0.06,
    "KS": 0.065, "KY": 0.06, "LA": 0.0445, "ME": 0.055, "MD": 0.06,
    "MA": 0.0625, "MI": 0.06, "MN": 0.06875, "MS": 0.07, "MO": 0.04225,
    "MT": 0.00, "NE": 0.055, "NV": 0.0685, "NH": 0.00, "NJ": 0.066,
    "NM": 0.05125, "NY": 0.04, "NC": 0.0475, "ND": 0.05, "OH": 0.0575,
    "OK": 0.045, "OR": 0.00, "PA": 0.06, "RI": 0.07, "SC": 0.06,
    "SD": 0.045, "TN": 0.07, "TX": 0.0625, "UT": 0.0485, "VT": 0.06,
    "VA": 0.053, "WA": 0.065, "WV": 0.06, "WI": 0.05, "WY": 0.04,
    "DC": 0.06,
}

STATE_NAMES = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas",
    "CA": "California", "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware",
    "FL": "Florida", "GA": "Georgia", "HI": "Hawaii", "ID": "Idaho",
    "IL": "Illinois", "IN": "Indiana", "IA": "Iowa", "KS": "Kansas",
    "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine", "MD": "Maryland",
    "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota", "MS": "Mississippi",
    "MO": "Missouri", "MT": "Montana", "NE": "Nebraska", "NV": "Nevada",
    "NH": "New Hampshire", "NJ": "New Jersey", "NM": "New Mexico", "NY": "New York",
    "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio", "OK": "Oklahoma",
    "OR": "Oregon", "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina",
    "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas", "UT": "Utah",
    "VT": "Vermont", "VA": "Virginia", "WA": "Washington", "WV": "West Virginia",
    "WI": "Wisconsin", "WY": "Wyoming", "DC": "District of Columbia",
}

df_tax = pd.DataFrame([
    {"state_code": k, "state_name": STATE_NAMES[k], "state_tax_rate": v, "nexus": 1 if v > 0 else 0}
    for k, v in STATE_TAX_RATES.items()
])
df_tax.to_csv(os.path.join(PROCESSED_DIR, "state_tax_rates.csv"), index=False)
print(f"   ✓ {len(df_tax)} state tax records")

# ─────────────────────────────────────────────────────────
# 2. PRODUCTS
# ─────────────────────────────────────────────────────────
print("\n[2/9] Generating products...")

CATEGORIES = ["Electronics", "Home & Garden", "Clothing", "Sports", "Books", "Toys", "Office Supplies", "Beauty"]
CATEGORY_WEIGHTS = [0.20, 0.15, 0.18, 0.12, 0.10, 0.10, 0.08, 0.07]

products = []
for i in range(1, N_PRODUCTS + 1):
    cat = random.choices(CATEGORIES, weights=CATEGORY_WEIGHTS)[0]
    unit_cost = round(random.uniform(5, 200), 2)
    margin = random.uniform(0.30, 0.65)
    unit_price = round(unit_cost / (1 - margin), 2)
    weight_lbs = round(random.uniform(0.1, 25.0), 2)
    products.append({
        "product_id": f"PRD-{i:04d}",
        "product_name": fake.catch_phrase().title()[:50],
        "category": cat,
        "unit_cost": unit_cost,
        "unit_price": unit_price,
        "weight_lbs": weight_lbs,
        "is_taxable": 1 if cat not in ["Books"] else 0,
        "sku": "".join(random.choices(string.ascii_uppercase + string.digits, k=8)),
    })

df_products = pd.DataFrame(products)
df_products.to_csv(os.path.join(PROCESSED_DIR, "products.csv"), index=False)
print(f"   ✓ {len(df_products)} products")

# ─────────────────────────────────────────────────────────
# 3. CUSTOMERS
# ─────────────────────────────────────────────────────────
print("\n[3/9] Generating customers...")

# Weight toward high-volume states
HIGH_VOLUME_STATES = ["CA", "TX", "FL", "NY", "IL", "PA", "OH", "GA", "NC", "WA"]
state_pool = HIGH_VOLUME_STATES * 5 + list(STATE_TAX_RATES.keys())

customers = []
for i in range(1, N_CUSTOMERS + 1):
    ctype = random.choices(["B2C", "B2B"], weights=[0.70, 0.30])[0]
    state = random.choice(state_pool)
    customers.append({
        "customer_id": f"CUST-{i:05d}",
        "company_name": fake.company() if ctype == "B2B" else None,
        "first_name": fake.first_name(),
        "last_name": fake.last_name(),
        "email": fake.email(),
        "phone": fake.phone_number()[:20],
        "address": fake.street_address(),
        "city": fake.city(),
        "state": state,
        "zip_code": fake.zipcode(),
        "customer_type": ctype,
        "credit_limit": round(random.uniform(5000, 50000), 2) if ctype == "B2B" else 1000,
        "payment_terms": random.choice(["Net 30", "Net 60", "Due on Receipt"]) if ctype == "B2B" else "Due on Receipt",
        "created_date": fake.date_between(start_date="-3y", end_date="-1y").strftime("%Y-%m-%d"),
    })

df_customers = pd.DataFrame(customers)
df_customers.to_csv(os.path.join(PROCESSED_DIR, "customers.csv"), index=False)
print(f"   ✓ {len(df_customers)} customers ({df_customers['customer_type'].value_counts().to_dict()})")

# ─────────────────────────────────────────────────────────
# 4. ORDERS + ORDER ITEMS
# ─────────────────────────────────────────────────────────
print("\n[4/9] Generating orders and order items...")

CARRIERS = ["FedEx", "UPS", "USPS", "DHL", "Amazon Logistics"]
CARRIER_BASE_RATES = {"FedEx": 8.50, "UPS": 7.80, "USPS": 5.20, "DHL": 9.10, "Amazon Logistics": 4.50}
CARRIER_PER_LB = {"FedEx": 0.45, "UPS": 0.40, "USPS": 0.30, "DHL": 0.55, "Amazon Logistics": 0.25}

def random_date(start, end):
    delta = end - start
    return start + timedelta(days=random.randint(0, delta.days))

def seasonal_weight(dt):
    """Higher weight for Q4 (holiday) and Q2 (spring)"""
    month = dt.month
    if month in [11, 12]: return 2.5
    if month in [3, 4, 5]: return 1.5
    if month in [1, 2]: return 0.8
    return 1.0

# Generate dates with seasonality
all_dates = [random_date(START_DATE, END_DATE) for _ in range(N_ORDERS * 3)]
weights = [seasonal_weight(d) for d in all_dates]
selected_dates = random.choices(all_dates, weights=weights, k=N_ORDERS)
selected_dates.sort()

orders = []
order_items = []
item_id = 1

for i, order_date in enumerate(selected_dates, 1):
    cust = df_customers.sample(1).iloc[0]
    carrier = random.choice(CARRIERS)
    status = random.choices(
        ["Paid", "Paid", "Paid", "Overdue", "Pending", "Cancelled"],
        weights=[50, 50, 50, 15, 10, 5]
    )[0]

    # 1–5 items per order
    n_items = random.choices([1, 2, 3, 4, 5], weights=[40, 30, 15, 10, 5])[0]
    items = df_products.sample(n_items)

    subtotal = 0
    total_weight = 0
    for _, prod in items.iterrows():
        qty = random.randint(1, 10)
        line_total = round(prod["unit_price"] * qty, 2)
        subtotal += line_total
        total_weight += prod["weight_lbs"] * qty
        order_items.append({
            "item_id": f"ITEM-{item_id:06d}",
            "order_id": f"ORD-{i:06d}",
            "product_id": prod["product_id"],
            "quantity": qty,
            "unit_price": prod["unit_price"],
            "unit_cost": prod["unit_cost"],
            "line_total": line_total,
            "is_taxable": prod["is_taxable"],
        })
        item_id += 1

    # Shipping cost
    shipping_cost = round(
        CARRIER_BASE_RATES[carrier] + (CARRIER_PER_LB[carrier] * total_weight),
        2
    )

    # Sales tax (on taxable items only)
    taxable_amount = sum(
        it["line_total"] for it in order_items
        if it["order_id"] == f"ORD-{i:06d}" and it["is_taxable"]
    )
    tax_rate = STATE_TAX_RATES.get(cust["state"], 0)
    sales_tax = round(taxable_amount * tax_rate, 2)

    invoice_total = round(subtotal + shipping_cost + sales_tax, 2)
    due_date = order_date + timedelta(days=30 if cust["payment_terms"] == "Net 30" else
                                       60 if cust["payment_terms"] == "Net 60" else 0)

    # Payment date logic
    if status == "Paid":
        days_to_pay = random.randint(1, 45)
        payment_date = (order_date + timedelta(days=days_to_pay)).strftime("%Y-%m-%d")
    else:
        payment_date = None

    orders.append({
        "order_id": f"ORD-{i:06d}",
        "customer_id": cust["customer_id"],
        "order_date": order_date.strftime("%Y-%m-%d"),
        "due_date": due_date.strftime("%Y-%m-%d"),
        "payment_date": payment_date,
        "status": status,
        "carrier": carrier,
        "shipping_cost": shipping_cost,
        "total_weight_lbs": round(total_weight, 2),
        "subtotal": round(subtotal, 2),
        "taxable_amount": round(taxable_amount, 2),
        "tax_rate": tax_rate,
        "sales_tax": sales_tax,
        "invoice_total": invoice_total,
        "state": cust["state"],
        "payment_terms": cust["payment_terms"],
    })

df_orders = pd.DataFrame(orders)
df_order_items = pd.DataFrame(order_items)
df_orders.to_csv(os.path.join(PROCESSED_DIR, "orders.csv"), index=False)
df_order_items.to_csv(os.path.join(PROCESSED_DIR, "order_items.csv"), index=False)
print(f"   ✓ {len(df_orders)} orders | {len(df_order_items)} line items")
print(f"   ✓ Total revenue: ${df_orders['subtotal'].sum():,.2f}")

# ─────────────────────────────────────────────────────────
# 5. BUDGET DATA (Monthly, by Department)
# ─────────────────────────────────────────────────────────
print("\n[5/9] Generating budget vs. actuals...")

DEPARTMENTS = ["Sales", "Marketing", "Operations", "Finance", "HR", "IT", "Customer Service"]
GL_ACCOUNTS = {
    "Revenue": ["Product Sales", "Service Revenue", "Shipping Revenue"],
    "COGS": ["Product Cost", "Freight In", "Packaging"],
    "OpEx": ["Salaries", "Marketing Spend", "Software & SaaS", "Office Rent",
             "Travel & Entertainment", "Professional Services", "Utilities", "Depreciation"],
}

budget_rows = []
actual_rows = []

for year in [2023, 2024]:
    for month in range(1, 13):
        for dept in DEPARTMENTS:
            for category, accounts in GL_ACCOUNTS.items():
                for account in accounts:
                    # Skip revenue for non-sales depts
                    if category == "Revenue" and dept not in ["Sales"]:
                        continue

                    base = random.uniform(10000, 120000)
                    # Revenue grows ~15% YoY with seasonality
                    if category == "Revenue":
                        seasonal = 1 + 0.4 * abs(month - 6.5) / 6.5
                        yoy = 1.15 if year == 2024 else 1.0
                        base = base * seasonal * yoy

                    budget = round(base, 2)
                    # Actual = budget ± variance (5–25%)
                    variance_pct = random.uniform(-0.20, 0.25)
                    actual = round(budget * (1 + variance_pct), 2)

                    budget_rows.append({
                        "year": year, "month": month,
                        "department": dept, "gl_category": category,
                        "gl_account": account, "amount": budget,
                    })
                    actual_rows.append({
                        "year": year, "month": month,
                        "department": dept, "gl_category": category,
                        "gl_account": account, "amount": actual,
                    })

df_budget = pd.DataFrame(budget_rows)
df_actuals = pd.DataFrame(actual_rows)
df_budget.to_csv(os.path.join(PROCESSED_DIR, "budget.csv"), index=False)
df_actuals.to_csv(os.path.join(PROCESSED_DIR, "actuals.csv"), index=False)
print(f"   ✓ {len(df_budget)} budget rows | {len(df_actuals)} actual rows")

# ─────────────────────────────────────────────────────────
# 6. GL / JOURNAL ENTRIES (for P&L Automation)
# ─────────────────────────────────────────────────────────
print("\n[6/9] Generating GL journal entries...")

gl_entries = []
entry_id = 1

for _, order in df_orders.iterrows():
    date = order["order_date"]
    if order["status"] == "Cancelled":
        continue

    # Revenue entry
    gl_entries.append({"entry_id": f"JE-{entry_id:07d}", "date": date,
                        "account_code": "4000", "account_name": "Product Sales",
                        "debit": 0, "credit": order["subtotal"],
                        "reference": order["order_id"], "description": "Order revenue"})
    entry_id += 1

    # COGS entry (use order items)
    order_items_sub = df_order_items[df_order_items["order_id"] == order["order_id"]]
    cogs = round((order_items_sub["unit_cost"] * order_items_sub["quantity"]).sum(), 2)
    gl_entries.append({"entry_id": f"JE-{entry_id:07d}", "date": date,
                        "account_code": "5000", "account_name": "Cost of Goods Sold",
                        "debit": cogs, "credit": 0,
                        "reference": order["order_id"], "description": "Order COGS"})
    entry_id += 1

    # Shipping revenue
    gl_entries.append({"entry_id": f"JE-{entry_id:07d}", "date": date,
                        "account_code": "4100", "account_name": "Shipping Revenue",
                        "debit": 0, "credit": order["shipping_cost"],
                        "reference": order["order_id"], "description": "Shipping charged to customer"})
    entry_id += 1

    # Sales tax collected
    if order["sales_tax"] > 0:
        gl_entries.append({"entry_id": f"JE-{entry_id:07d}", "date": date,
                            "account_code": "2100", "account_name": "Sales Tax Payable",
                            "debit": 0, "credit": order["sales_tax"],
                            "reference": order["order_id"], "description": "Sales tax collected"})
        entry_id += 1

df_gl = pd.DataFrame(gl_entries)
df_gl.to_csv(os.path.join(PROCESSED_DIR, "gl_journal_entries.csv"), index=False)
print(f"   ✓ {len(df_gl)} GL journal entries")

# ─────────────────────────────────────────────────────────
# 7. OPEN INVOICES (for AR Aging)
# ─────────────────────────────────────────────────────────
print("\n[7/9] Generating open invoice / AR data...")

# Only unpaid orders become AR
today = datetime(2025, 1, 15)  # Simulated "today" for aging
ar_orders = df_orders[df_orders["status"].isin(["Overdue", "Pending"])].copy()
ar_orders["days_outstanding"] = ar_orders["due_date"].apply(
    lambda x: max(0, (today - datetime.strptime(x, "%Y-%m-%d")).days)
)
ar_orders["aging_bucket"] = ar_orders["days_outstanding"].apply(
    lambda d: "Current" if d == 0 else
              "1-30 Days" if d <= 30 else
              "31-60 Days" if d <= 60 else
              "61-90 Days" if d <= 90 else
              "90+ Days"
)
ar_orders.to_csv(os.path.join(PROCESSED_DIR, "ar_open_invoices.csv"), index=False)
print(f"   ✓ {len(ar_orders)} open invoices | Total AR: ${ar_orders['invoice_total'].sum():,.2f}")

# ─────────────────────────────────────────────────────────
# 8. SHIPPING DATA (for Allocation Engine)
# ─────────────────────────────────────────────────────────
print("\n[8/9] Generating shipment data...")

shipments = []
for i, order in df_orders.iterrows():
    if order["status"] == "Cancelled":
        continue
    # Some orders are bundled in multi-order shipments
    is_bundled = random.random() < 0.15  # 15% bundled
    shipments.append({
        "shipment_id": f"SHIP-{len(shipments)+1:06d}",
        "order_id": order["order_id"],
        "carrier": order["carrier"],
        "ship_date": order["order_date"],
        "total_weight_lbs": order["total_weight_lbs"],
        "billed_shipping_cost": order["shipping_cost"],
        "is_bundled_shipment": int(is_bundled),
        "tracking_number": "".join(random.choices(string.ascii_uppercase + string.digits, k=12)),
        "zone": random.randint(1, 8),
        "delivery_days": random.randint(1, 10),
    })

df_shipments = pd.DataFrame(shipments)
df_shipments.to_csv(os.path.join(PROCESSED_DIR, "shipments.csv"), index=False)
print(f"   ✓ {len(df_shipments)} shipment records")

# ─────────────────────────────────────────────────────────
# 9. LOAD ALL DATA INTO SQLITE
# ─────────────────────────────────────────────────────────
print("\n[9/9] Loading all data into nexoria.db...")

conn = sqlite3.connect(DB_PATH)

tables = {
    "state_tax_rates": df_tax,
    "products": df_products,
    "customers": df_customers,
    "orders": df_orders,
    "order_items": df_order_items,
    "budget": df_budget,
    "actuals": df_actuals,
    "gl_journal_entries": df_gl,
    "ar_open_invoices": ar_orders,
    "shipments": df_shipments,
}

for table_name, df in tables.items():
    df.to_sql(table_name, conn, if_exists="replace", index=False)
    print(f"   ✓ {table_name}: {len(df):,} rows loaded")

conn.close()

# ─────────────────────────────────────────────────────────
# SUMMARY
# ─────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("  ✅ Phase 0 Complete!")
print("=" * 60)
print(f"\n  Database: {DB_PATH}")
print(f"  CSVs:     {PROCESSED_DIR}")
print("\n  Summary:")
print(f"    Orders:          {len(df_orders):,}")
print(f"    Order Items:     {len(df_order_items):,}")
print(f"    Customers:       {len(df_customers):,}")
print(f"    Products:        {len(df_products):,}")
print(f"    GL Entries:      {len(df_gl):,}")
print(f"    Open AR Items:   {len(ar_orders):,}")
print(f"    Shipments:       {len(df_shipments):,}")
print(f"\n    Total Revenue:   ${df_orders['subtotal'].sum():>12,.2f}")
print(f"    Total Tax Coll:  ${df_orders['sales_tax'].sum():>12,.2f}")
print(f"    Total Shipping:  ${df_orders['shipping_cost'].sum():>12,.2f}")
print(f"    Total AR Open:   ${ar_orders['invoice_total'].sum():>12,.2f}")
print("\n  Next step: Phase 1 — Invoicing Automation")
print("=" * 60)
