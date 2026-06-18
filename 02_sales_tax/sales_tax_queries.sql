-- Query 1: Order-Level Tax Reconciliation
-- This query checks each order's collected tax against the expected tax based on state nexus rules.
SELECT 
    o.order_id,
    o.order_date,
    o.customer_id,
    o.state,
    t.nexus AS has_nexus,
    o.taxable_amount,
    t.state_tax_rate AS correct_rate,
    ROUND(o.taxable_amount * t.state_tax_rate * t.nexus, 2) AS expected_tax,
    o.sales_tax AS collected_tax,
    ROUND(o.sales_tax - (o.taxable_amount * t.state_tax_rate * t.nexus), 2) AS discrepancy
FROM orders o
JOIN state_tax_rates t ON o.state = t.state_code;

-- Query 2: State-Level Summary
-- This query summarizes the tax liability and discrepancies by state.
SELECT 
    o.state,
    t.state_name,
    t.nexus AS has_nexus,
    SUM(o.taxable_amount) AS total_taxable_sales,
    SUM(ROUND(o.taxable_amount * t.state_tax_rate * t.nexus, 2)) AS expected_tax_liability,
    SUM(o.sales_tax) AS actual_tax_collected,
    SUM(ROUND(o.sales_tax - (o.taxable_amount * t.state_tax_rate * t.nexus), 2)) AS net_discrepancy
FROM orders o
JOIN state_tax_rates t ON o.state = t.state_code
GROUP BY o.state, t.state_name, t.nexus
ORDER BY ABS(SUM(ROUND(o.sales_tax - (o.taxable_amount * t.state_tax_rate * t.nexus), 2))) DESC;
