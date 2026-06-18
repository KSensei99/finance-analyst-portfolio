# Upwork Project Catalog Listing Guide (For Sale Project)
## Service: Custom Excel & Power BI Financial Dashboard Automated with Python & VBA

This guide explains how to package this portfolio project into a **Project Catalog** listing (a pre-packaged, fixed-price service) that clients can browse and buy directly from your profile.

---

## 🚪 Step-by-Step Project Catalog Creation

### Step 1: Overview
1. Go to your Upwork **Project Catalog** dashboard and click **Create a Project**.
2. **Project Title**: `You will get a custom Excel and Power BI financial dashboard automated with Python or VBA`
   > *Note*: Upwork titles must start with "You will get..." and are capped at 75 characters. Keep it clear, active, and keyword-rich.
3. **Category**: `Finance & Accounting` -> `Financial Planning & Analysis (FP&A)` or `Data Visualization & Dashboards`.
4. **Project Attributes**:
   * **Dashboard Tools**: Excel, Power BI, Google Sheets.
   * **Programming Languages**: Python, VBA, M, DAX.
   * **Industry Focus**: E-commerce, Retail, B2B Distribution.

---

### Step 2: Pricing Tiers
Configure three distinct pricing tiers to capture clients at different budgets:

| Tier Features | Basic ($150) | Standard ($350) | Advanced ($650) |
| :--- | :--- | :--- | :--- |
| **Tier Name** | *Automated Excel template* | *Full BI Dashboard + Excel Reports* | *Enterprise Automation & Forecasting* |
| **Description** | Custom Excel dashboard, VBA macro automated invoicing, and multi-state tax lookups. | Dynamic Star-Schema Power BI dashboard, automated P&L workbook, and AR Aging pivot sheets. | End-to-End Suite: Power BI, P&L, AR, SARIMA forecasting script, and Google Workspace pipeline. |
| **Delivery Time**| 3 Days | 5 Days | 8 Days |
| **Revisions** | 1 Revision | 2 Revisions | 3 Revisions |
| **Source File** | ✅ Yes | ✅ Yes | ✅ Yes |
| **Data Cleaning** | ❌ No | ✅ Yes (up to 10K rows) | ✅ Yes (unlimited rows) |
| **Documentation** | ✅ Setup instructions | ✅ Walkthrough guide | ✅ Full documentation + setup support |

---

### Step 3: Description & Steps
* **Project Description**:
  Copy and paste the following detailed description:
  ```markdown
  Are manual processes, lagged reporting, and data entry errors slowing down your finance operations? Let's automate them.
  
  I will engineer a customized, automated financial analysis and dashboard suite tailored for your business. Whether you need structured Power BI dashboards, automated Excel reports via VBA, or cloud integrations, I provide enterprise-grade solutions.
  
  What this project can include:
  1. Transactional Billing: Automated invoice compilation & PDF exports (VBA or Python).
  2. Multi-State Tax Compliance: Power Query lookups with category-specific tax exemptions.
  3. Freight Shipping Cost Allocation: Multi-carrier weighted cost distribution back to individual items to determine true SKU gross margins.
  4. Executive BI Dashboard: Premium, dark-themed Power BI reports (Executive Summaries & Drill-Downs).
  5. Operational Reporting: Automated P&L workbooks and Accounts Receivable (AR) aging sheets with color-coded alerts.
  6. Demand Forecasting: Statistical forecast models (Holt-Winters, Seasonal ARIMA) to predict sales and optimize capital.
  7. Google Cloud Integrations: Automated quote sheets to Gmail/Drive pipeline via Apps Script.
  
  Let's transform your spreadsheets from manual entry bottlenecks into an automated, decision-ready data system.
  ```

* **Project Steps (Workflow)**:
  * **Step 1: Scope & Data Gathering** (1 Day): Reviewing your database schema, General Ledger structure, and business rules.
  * **Step 2: Database & Automation Engineering** (2-4 Days): Setting up clean SQLite/Excel data layers, writing VBA/Python automation scripts, and compiling reports.
  * **Step 3: Dashboard Design & Handoff** (2 Days): Designing the visual dashboards in Power BI/Excel, setting up the Google Apps Script pipeline, running verifications, and delivering files.

---

### Step 4: Gallery & Media
* **Cover Image**: Upload a high-resolution 1000x750px cover image. Show a mockup of the custom dark-mode Power BI dashboard with Heritage Gold highlights alongside the Excel P&L.
* **Additional Images**: Take screenshots of your Excel tables, the matplotlib forecasting plot, and the SVG data flow blueprint.
* **Sample Document**: Upload `Nexoria_BVA_Dashboard.pdf` as your PDF portfolio asset.

---

### Step 5: FAQs (High-Converting)

#### Q: Do I need a paid Power BI Pro license?
**A:** No. You can open, view, and interact with the `.pbix`/`.pbip` files locally using the free **Power BI Desktop** application. A paid Power BI Pro/Premium license is only required if you want to publish the dashboard to the cloud and share it with multiple users inside your company.

#### Q: Can you integrate this with QuickBooks, Xero, or Stripe?
**A:** Yes. We can write Python scripts or Power Query pipelines that connect directly to API endpoints or process standard CSV exports from QuickBooks, Xero, Stripe, or any ERP system to automate your reports.

#### Q: Can I run Python/VBA automation on a Mac?
**A:** VBA macros and Python scripts run on both Windows and macOS. However, Windows COM automation (pywin32) is exclusive to Windows. If you are on Mac, we will build standard, cross-platform `openpyxl` Python scripts or shift the automation to Google Sheets using Google Apps Script.

#### Q: Is the data secure?
**A:** Absolutely. The database (`nexoria.db`) is stored locally on your machine. The scripts do not transmit your financial data to external servers.

---

### Step 6: Requirements for the Client
In the **Requirements** section, write the following list of what the client must provide to start the project:
1. Sample transaction data or raw CSV ledger reports (e.g. Sales, Expenses, Shipping).
2. Chart of Accounts (GL categories and code structures).
3. Desired color themes or company brand guidelines.
4. Business rules (e.g., specific states with tax exemptions, freight allocation weightings).
