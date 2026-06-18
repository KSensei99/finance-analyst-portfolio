# 📊 Nexoria Commerce Inc. — Finance Automation & Analytics Suite
> **A Premium FP&A, Process Automation, and Business Intelligence Portfolio Project for B2B Distribution**

This portfolio demonstrates end-to-end finance engineering and data analyst capabilities. It solves critical manual bottlenecks for **Nexoria Commerce Inc.**, a B2B product distributor managing 32,000+ transactional rows across 12 US states.

---

## 🏢 The Client Scenario
Nexoria Commerce Inc. distributes office supplies, furniture, electronics, and safety equipment. The company's 5-person finance team spent **60% of their time** on manual invoicing, state-by-state tax calculations, freight cost reconciliations, and copy-pasting Excel templates. We implemented an integrated automation suite to streamline operations, reduce billing discrepancies, and provide real-time executive visibility.

### Key Financial Impact Metrics
| Performance Metric | Before Automation | After Automation | Business Impact |
|--------------------|:-----------------:|:----------------:|:----------------|
| **Weekly Invoicing Labor** | 40+ Hours | < 3 Minutes | **99.8% Time Saved** (reallocated to analysis) |
| **Sales Tax Billing Errors** | ~$8,000/year | $0.00 | **$8K/yr Saved** in credit memo write-offs |
| **Gross Margin Accuracy** | 78% | 97%+ | **Surgical Margin Tracking** at SKU level |
| **Executive Reporting Lag** | 2 Days | 20 Minutes | **Real-Time Data** for agile CFO decisions |
| **Quote-to-Invoice Cycle** | 3 Days | < 5 Minutes | **Faster cash collection** (reduced DSO) |

---

## 🧩 Project Modules & Roadmap

| Module | Core Tooling | Status | business Impact |
| :--- | :--- | :---: | :--- |
| **[01. Invoice Automation](./01_invoicing/)** | Excel VBA / Python (`ReportLab`) | ✅ Complete | Auto-generates 200+ monthly invoices & PDFs; logs runs. |
| **[02. Sales Tax Engine](./02_sales_tax/)** | Power Query / Excel Formulas | ✅ Complete | Multi-state tax lookup with automated category exemptions. |
| **[03. Shipping Reconciler](./03_shipping_allocation/)** | Python (`pandas`) / Excel | ✅ Complete | Reconciles freight bills; weighted cost allocation. |
| **[04. Executive BI Dashboard](./04_budget_vs_actual/)** | Power BI / DAX / M Query | ✅ Complete | 2-page star-schema dashboard in Midnight Navy theme. |
| **[05. P&L Automation](./05_pl_automation/)** | Python (`openpyxl` + COM) | ✅ Complete | Dynamic P&L workbook with VBA/COM month selectors. |
| **[06. AR Aging Validation](./06_ar_aging/)** | Python / Excel VBA | ✅ Complete | Aging bucket calculator with green-yellow-red warning system. |
| **[07. Revenue Forecasting](./07_forecasting/)** | Python (`statsmodels` SARIMA) | ✅ Complete | 12-month statistical revenue forecasting & chart exports. |
| **[08. Quote-to-Invoice GAS](./08_quotes_pipeline/)** | Google Apps Script / Gmail | ✅ Complete | One-click sheets quote conversion to emailed PDF invoices. |
| **[09. Case Study & Upwork Copy](./case_study/)** | Markdown / Docs | ✅ Complete | Comprehensive write-up and copy-paste proposal template. |

---

## 🛠️ Technology Stack
* **Database & Storage**: SQLite (nexoria.db with 10 tables and 32,000+ rows)
* **Spreadsheet & Macros**: Microsoft Excel (COM Automation, VBA macros, Power Query)
* **Business Intelligence**: Power BI Desktop & Service (DAX, Star Schema, custom JSON themes)
* **Data Science & ML**: Python 3.14 (pandas, numpy, statsmodels, openpyxl, reportlab)
* **Cloud Automation**: Google Apps Script, Gmail API, Google Drive API

---

## 📁 Repository Structure
```
finance-analyst-portfolio/
├── README.md                          ← Main portfolio portal
├── data/
│   ├── generate_mock_data.py          ← Populates central nexoria.db database
│   └── nexoria.db                     ← SQLite relational data warehouse
├── 01_invoicing/                      ← Module 1: VBA & Python invoicing
├── 02_sales_tax/                      ← Module 2: Power Query tax calculator
├── 03_shipping_allocation/            ← Module 3: Freight cost allocator
├── 04_budget_vs_actual/               ← Module 4: Power BI executive dashboard
├── 05_pl_automation/                  ← Module 5: Excel P&L automation
├── 06_ar_aging/                       ← Module 6: Accounts Receivable aging
├── 07_forecasting/                    ← Module 7: SARIMA forecasting
├── 08_quotes_pipeline/                ← Module 8: Apps Script quote emailer
└── case_study/                        ← Module 9: Case study & Upwork copy
```

---

## 🚀 Quick Start
To run the automated scripts locally:

1. **Clone the repository**:
   ```bash
   git clone https://github.com/YOUR_USERNAME/finance-analyst-portfolio.git
   cd finance-analyst-portfolio
   ```

2. **Generate the SQLite relational database**:
   ```bash
   cd data
   python generate_mock_data.py
   cd ..
   ```

3. **Navigate to the individual module directories** and follow their respective `README.md` setup instructions.

---

*This portfolio is maintained for Upwork Finance Analyst & FP&A Specializations.*
