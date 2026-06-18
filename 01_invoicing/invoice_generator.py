"""
=============================================================
  Nexoria Commerce Inc. — Invoice Generator
  Finance Analyst Portfolio | Phase 1 | python-pro skill
=============================================================
  Reads orders + customers + order_items from nexoria.db
  Generates one professional PDF invoice per order.
  
  Usage:
    python invoice_generator.py              # All orders
    python invoice_generator.py --limit 20   # First 20 orders
    python invoice_generator.py --order ORD-000001  # Single order
=============================================================
"""

from __future__ import annotations

import argparse
import os
import sqlite3
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

# reportlab imports
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, HRFlowable
)

# ─── Paths ───────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
DB_PATH    = BASE_DIR.parent / "data" / "nexoria.db"
OUTPUT_DIR = BASE_DIR / "invoices"
OUTPUT_DIR.mkdir(exist_ok=True)

# ─── Brand Colors ─────────────────────────────────────────────
NAVY    = colors.HexColor("#1B3A6B")
TEAL    = colors.HexColor("#2E86AB")
LIGHT   = colors.HexColor("#F0F4F8")
WHITE   = colors.white
GREY    = colors.HexColor("#6B7280")
BLACK   = colors.black
GREEN   = colors.HexColor("#059669")  # for "PAID" status
RED     = colors.HexColor("#DC2626")  # for "OVERDUE"

# ─── DB helpers ──────────────────────────────────────────────
def get_order(conn: sqlite3.Connection, order_id: str) -> Optional[dict]:
    cur = conn.execute(
        "SELECT o.*, c.first_name, c.last_name, c.company_name, "
        "c.email, c.phone, c.address, c.city, c.state, c.zip_code, "
        "c.customer_type, c.payment_terms "
        "FROM orders o JOIN customers c USING(customer_id) "
        "WHERE o.order_id = ?", (order_id,)
    )
    row = cur.fetchone()
    if not row:
        return None
    return dict(zip([d[0] for d in cur.description], row))

def get_items(conn: sqlite3.Connection, order_id: str) -> list[dict]:
    cur = conn.execute(
        "SELECT oi.*, p.product_name, p.category, p.sku "
        "FROM order_items oi JOIN products p USING(product_id) "
        "WHERE oi.order_id = ?", (order_id,)
    )
    cols = [d[0] for d in cur.description]
    return [dict(zip(cols, row)) for row in cur.fetchall()]

def get_all_order_ids(conn: sqlite3.Connection, limit: Optional[int]) -> list[str]:
    query = "SELECT order_id FROM orders WHERE status != 'Cancelled' ORDER BY order_date"
    if limit:
        query += f" LIMIT {limit}"
    return [r[0] for r in conn.execute(query).fetchall()]

# ─── PDF Generation ──────────────────────────────────────────
def fmt_currency(val) -> str:
    try:
        return f"${float(val):,.2f}"
    except (TypeError, ValueError):
        return "-"

def fmt_pct(val) -> str:
    try:
        return f"{float(val)*100:.3f}%"
    except (TypeError, ValueError):
        return "-"

def status_color(status: str) -> colors.Color:
    return GREEN if status == "Paid" else RED if status == "Overdue" else GREY

def generate_pdf(order: dict, items: list[dict], out_path: Path) -> None:
    doc = SimpleDocTemplate(
        str(out_path),
        pagesize=letter,
        rightMargin=0.5 * inch,
        leftMargin=0.5 * inch,
        topMargin=0.5 * inch,
        bottomMargin=0.5 * inch,
    )

    styles = getSampleStyleSheet()
    story = []

    # ── HEADER ────────────────────────────────────────────────
    header_data = [
        [
            Paragraph(
                '<font name="Helvetica-Bold" size="20" color="#FFFFFF">NEXORIA COMMERCE INC.</font><br/>'
                '<font name="Helvetica" size="8" color="#FFFFFF">1234 Commerce Way, Suite 100 | Austin, TX 78701</font><br/>'
                '<font name="Helvetica" size="8" color="#FFFFFF">billing@nexoriacommerce.com | (800) 555-0100</font>',
                styles["Normal"]
            ),
            Paragraph(
                '<font name="Helvetica-Bold" size="30" color="#FFFFFF">INVOICE</font>',
                ParagraphStyle("inv", alignment=2)
            )
        ]
    ]
    header_table = Table(header_data, colWidths=[4.5 * inch, 2.5 * inch])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 14),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 14),
        ("LEFTPADDING", (0, 0), (0, 0), 12),
        ("RIGHTPADDING", (1, 0), (1, 0), 12),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 4))

    # Teal accent strip
    story.append(HRFlowable(width="100%", thickness=5, color=TEAL, spaceAfter=6))

    # ── META INFO ─────────────────────────────────────────────
    status = order.get("status", "")
    scol = status_color(status)

    meta_data = [
        ["Invoice #:", order["order_id"],  "Status:",
         Paragraph(f'<font color="{scol.hexval()}" name="Helvetica-Bold">{status.upper()}</font>', styles["Normal"])],
        ["Order Date:", order["order_date"],  "Due Date:", order["due_date"]],
        ["Payment Terms:", order.get("payment_terms", "N/A"), "Carrier:", order.get("carrier", "N/A")],
        ["Customer ID:", order["customer_id"], "State:", order.get("state", "N/A")],
    ]
    meta_table = Table(meta_data, colWidths=[1.4*inch, 2.1*inch, 1.4*inch, 2.1*inch])
    meta_table.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Helvetica", 8),
        ("FONT", (0, 0), (0, -1), "Helvetica-Bold", 8),
        ("FONT", (2, 0), (2, -1), "Helvetica-Bold", 8),
        ("BACKGROUND", (0, 0), (0, -1), LIGHT),
        ("BACKGROUND", (2, 0), (2, -1), LIGHT),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#D1D5DB")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 8))

    # ── BILL TO ───────────────────────────────────────────────
    name = f"{order.get('first_name','')} {order.get('last_name','')}".strip()
    company = order.get("company_name") or ""
    addr = f"{order.get('address','')}, {order.get('city','')}, {order.get('state','')} {order.get('zip_code','')}"

    bill_data = [
        [Paragraph('<b>BILL TO</b>', styles["Normal"]), ""],
        [name + (f"\n{company}" if company else ""), addr],
    ]
    bill_table = Table(bill_data, colWidths=[3.5*inch, 3.5*inch])
    bill_table.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), "Helvetica", 8),
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8),
        ("TEXTCOLOR", (0, 0), (-1, 0), TEAL),
        ("BACKGROUND", (0, 0), (-1, 0), LIGHT),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#D1D5DB")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("SPAN", (0, 0), (-1, 0)),
    ]))
    story.append(bill_table)
    story.append(Spacer(1, 10))

    # ── LINE ITEMS ────────────────────────────────────────────
    col_headers = ["SKU", "Product", "Category", "Qty", "Unit Price", "Line Total", "Taxable"]
    item_rows = [col_headers]
    for it in items:
        item_rows.append([
            it.get("sku", ""),
            it.get("product_name", "")[:30],
            it.get("category", ""),
            str(int(it.get("quantity", 0))),
            fmt_currency(it.get("unit_price")),
            fmt_currency(it.get("line_total")),
            "Yes" if it.get("is_taxable") else "No",
        ])

    items_table = Table(
        item_rows,
        colWidths=[0.9*inch, 2.1*inch, 1.0*inch, 0.5*inch, 0.85*inch, 0.9*inch, 0.65*inch],
        repeatRows=1
    )
    items_style = TableStyle([
        # Header row
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        # Data rows
        ("FONT", (0, 1), (-1, -1), "Helvetica", 8),
        ("ALIGN", (3, 1), (5, -1), "RIGHT"),
        ("ALIGN", (0, 1), (2, -1), "LEFT"),
        ("ALIGN", (6, 1), (6, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#D1D5DB")),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, LIGHT]),
    ])
    items_table.setStyle(items_style)
    story.append(items_table)
    story.append(Spacer(1, 10))

    # ── TOTALS ─────────────────────────────────────────────────
    subtotal   = float(order.get("subtotal", 0))
    tax_rate   = float(order.get("tax_rate", 0))
    sales_tax  = float(order.get("sales_tax", 0))
    shipping   = float(order.get("shipping_cost", 0))
    total      = float(order.get("invoice_total", 0))
    taxable    = float(order.get("taxable_amount", 0))

    totals_data = [
        ["", "Subtotal:",           fmt_currency(subtotal)],
        ["", "Taxable Amount:",     fmt_currency(taxable)],
        ["", f"Tax Rate ({order.get('state','')}):", fmt_pct(tax_rate)],
        ["", "Sales Tax:",          fmt_currency(sales_tax)],
        ["", "Shipping:",           fmt_currency(shipping)],
        ["", "INVOICE TOTAL:",      fmt_currency(total)],
    ]
    totals_table = Table(totals_data, colWidths=[4.0*inch, 1.8*inch, 1.2*inch])
    totals_table.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -2), "Helvetica", 9),
        ("FONT", (1, -1), (-1, -1), "Helvetica-Bold", 10),
        ("FONT", (1, 0), (1, -2), "Helvetica-Bold", 8),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("ALIGN", (2, 0), (2, -1), "RIGHT"),
        ("BACKGROUND", (1, -1), (-1, -1), NAVY),
        ("TEXTCOLOR", (1, -1), (-1, -1), WHITE),
        ("BACKGROUND", (1, 0), (-1, -2), LIGHT),
        ("GRID", (1, 0), (-1, -1), 0.3, colors.HexColor("#D1D5DB")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (1, 0), (1, -1), 6),
        ("RIGHTPADDING", (2, 0), (2, -1), 8),
    ]))
    story.append(totals_table)
    story.append(Spacer(1, 16))

    # ── FOOTER ────────────────────────────────────────────────
    story.append(HRFlowable(width="100%", thickness=3, color=TEAL, spaceBefore=4, spaceAfter=4))
    story.append(Paragraph(
        '<font name="Helvetica" size="7" color="#6B7280">'
        'Thank you for your business! | Make checks payable to: Nexoria Commerce Inc. | '
        'Bank: First National Bank | Routing: 021000021 | Account: 1234567890 | '
        'Questions? billing@nexoriacommerce.com'
        '</font>',
        ParagraphStyle("footer", alignment=1)
    ))

    doc.build(story)


# ─── Main ─────────────────────────────────────────────────────
def main() -> None:
    parser = argparse.ArgumentParser(description="Nexoria Invoice Generator")
    parser.add_argument("--limit", type=int, default=None, help="Max invoices to generate")
    parser.add_argument("--order", type=str, default=None, help="Single order ID")
    args = parser.parse_args()

    if not DB_PATH.exists():
        print(f"ERROR: Database not found at {DB_PATH}")
        print("Run data/generate_mock_data.py first.")
        sys.exit(1)

    conn = sqlite3.connect(DB_PATH)

    if args.order:
        order_ids = [args.order]
    else:
        order_ids = get_all_order_ids(conn, args.limit)

    print(f"Generating {len(order_ids)} invoice(s)...")
    success = 0
    errors  = 0

    for oid in order_ids:
        try:
            order = get_order(conn, oid)
            if not order:
                print(f"  [SKIP] {oid} — not found")
                continue
            items = get_items(conn, oid)
            out_path = OUTPUT_DIR / f"{oid}.pdf"
            generate_pdf(order, items, out_path)
            success += 1
            if success % 100 == 0:
                print(f"  ... {success} invoices generated")
        except Exception as e:
            print(f"  [ERROR] {oid}: {e}")
            errors += 1

    conn.close()

    print(f"\nDone!")
    print(f"  Generated: {success} invoices")
    print(f"  Errors:    {errors}")
    print(f"  Output:    {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
