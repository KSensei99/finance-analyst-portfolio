# Module 08 — Google Workspace Quote-to-Invoice Pipeline

> 🌐 **Live Demo Spreadsheet**: [Google Sheet Template](https://docs.google.com/spreadsheets/d/1TkUIZxnkKs1wcitzRc9f8xej_EulyQB4gbdTVBXP-mc/edit?usp=sharing)

## Business Case & Problem Statement
Sales representatives at Nexoria Commerce Inc. created and negotiated customer quotes in Google Sheets. When a deal was closed ("Won"), the sales rep notified the finance team, who manually re-entered the quote data into an invoice template, converted it to PDF, and drafted an email to the customer. This manual handoff took **3+ days**, introducing data entry errors and delaying cash collections (DSO).

## The Automation Solution
We built an automated, one-click quote conversion pipeline using Google Apps Script:
1. **Google Sheets Interface**: A custom menu item `Nexoria Tools -> Convert Quote to Invoice` is injected into the spreadsheet UI.
2. **Transactional Validation**: The script validates that a row is selected and that the quote stage is marked as `Won` before initiating billing.
3. **Template Compilation**: Copies an styled Google Doc invoice template, dynamically performs tax calculations based on the customer's state via a `TaxRates` worksheet, and replaces placeholders (e.g. `{{CUSTOMER}}`, `{{INVOICE_ID}}`, `{{TOTAL}}`).
4. **Drive & Email Delivery**: Saves the invoice as a PDF in a specified Google Drive folder and sends it to the customer's billing email via the `GmailApp` service.
5. **Audit logging**: Appends a log row to a `Pipeline` worksheet tracking the invoice ID, total, status (`Sent`), and timestamp, while marking the quote row as `Invoiced`.

## Performance & Business Impact
* **Process Acceleration**: Reduced the quote-to-invoice processing cycle **from 3 days to under 5 seconds**.
* **Billing Accuracy**: Pulls directly from the negotiated quote and central tax matrix, eliminating billing discrepancies.
* **Audit Trail**: Logs all transactions to a central pipeline tracker sheet automatically.

## Folder Contents
* [QuoteToInvoice.gs](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/08_quotes_pipeline/QuoteToInvoice.gs) — Google Apps Script macro code.
* [quotes_sample.csv](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/08_quotes_pipeline/quotes_sample.csv) — Sample quotes data to populate the test sheet.

---

## Step-by-Step Setup Instructions

### Step 1: Set Up the Google Spreadsheet
1. Create a new Google Spreadsheet and name it **Nexoria Quote-to-Invoice Tracker**.
2. Create three worksheets (tabs) named:
   * **Quotes**: Paste the contents of [quotes_sample.csv](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/08_quotes_pipeline/quotes_sample.csv) here.
   * **TaxRates**: Paste the state-to-tax-rate mappings (State in Col A, Rate in Col B).
   * **Pipeline**: Add headers in Row 1: `Invoice ID`, `Quote ID`, `Customer`, `Total ($)`, `Status`, `Date`.

### Step 2: Create the Google Drive Invoices Folder
1. Open your Google Drive.
2. Create a new folder named **Nexoria Invoices**.
3. Open the folder and copy the Folder ID from the URL (the string of characters after `/folders/`).
4. Paste this ID into `QuoteToInvoice.gs` as the value for `INVOICES_FOLDER_ID`.

### Step 3: Create the Google Doc Invoice Template
1. Create a new Google Document named **Nexoria Invoice Template**.
2. Design the invoice template. Insert the following placeholders exactly where you want the dynamic text to appear:
   * `{{INVOICE_ID}}`
   * `{{DATE}}`
   * `{{CUSTOMER}}`
   * `{{STATE}}`
   * `{{SUBTOTAL}}`
   * `{{TAX_RATE}}`
   * `{{TAX_AMOUNT}}`
   * `{{TOTAL}}`
3. Copy the Document ID from the URL (the string of characters between `/d/` and `/edit`).
4. Paste this ID into `QuoteToInvoice.gs` as the value for `DOC_TEMPLATE_ID`.

### Step 4: Inject the Apps Script
1. In your Google Spreadsheet, click **Extensions -> Apps Script**.
2. Clear any default code in the editor.
3. Paste the contents of [QuoteToInvoice.gs](file:///C:/Users/marsh/Downloads/Upwork/finance-analyst-portfolio/08_quotes_pipeline/QuoteToInvoice.gs).
4. Save the project and click **Run** (you will be prompted to authorize permissions for Google Sheets, Docs, Drive, and Gmail).
5. Reload the Google Spreadsheet. You will see a new menu item **Nexoria Tools** appear.

### Step 5: Test the Pipeline
1. Select a row in the **Quotes** sheet that is marked as **Won** (e.g. row 2: Summit Solutions).
2. Click **Nexoria Tools -> Convert Quote to Invoice**.
3. Check your designated Google Drive folder for the generated PDF invoice.
4. Check your Gmail "Sent" folder to verify the invoice email was delivered.
5. Check the **Pipeline** sheet to confirm the transaction log row was appended.
