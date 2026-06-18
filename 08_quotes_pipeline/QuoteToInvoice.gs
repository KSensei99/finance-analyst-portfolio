// Extensions → Apps Script → paste this code → Save → Run onOpen once to authorize
// Replace the constants below with your actual Google Drive Folder and Document Template IDs.

const DOC_TEMPLATE_ID = 'YOUR_GOOGLE_DOC_TEMPLATE_ID'; // Copy from Document Template URL
const INVOICES_FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID';       // Copy from Invoices Folder URL

/**
 * Creates a custom menu in the Google Sheets interface on open.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Nexoria Tools')
    .addItem('Convert Quote to Invoice', 'convertQuoteToInvoice')
    .addToUi();
}

/**
 * Converts the active row's "Won" quote into a PDF invoice, emails it, and logs the pipeline status.
 */
function convertQuoteToInvoice() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const qSht   = ss.getSheetByName('Quotes');
  const logSht = ss.getSheetByName('Pipeline');
  const activeCell = qSht.getActiveCell();
  const row    = activeCell.getRow();

  // Validate selection row
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Error: Please select a valid quote row (row 2 or below) first.');
    return;
  }

  // Retrieve values from the selected row
  // Expected columns: A:Quote_ID, B:Company_Name, C:Email, D:State, E:Subtotal, F:Stage, G:Rep, H:Expiry_Date
  const [quoteId, customer, email, state, subtotal, stage] =
    qSht.getRange(row, 1, 1, 6).getValues()[0];

  // Stage validation (only convert 'Won' quotes)
  if (stage !== 'Won') {
    SpreadsheetApp.getUi().alert('Error: Only "Won" quotes can be converted. Current stage: ' + stage);
    return;
  }

  // Tax lookup from the 'TaxRates' tab
  const taxSht = ss.getSheetByName('TaxRates');
  if (!taxSht) {
    SpreadsheetApp.getUi().alert('Error: "TaxRates" worksheet not found. Please create it first.');
    return;
  }
  const taxData  = taxSht.getDataRange().getValues();
  const taxMap   = Object.fromEntries(taxData.slice(1).map(r => [r[0], r[1]]));
  const taxRate  = taxMap[state] || 0;
  
  // Financial Calculations
  const taxAmt   = subtotal * taxRate;
  const total    = subtotal + taxAmt;
  const invId    = 'INV-' + Utilities.formatDate(new Date(), 'UTC', 'yyyyMMdd') + '-' + quoteId;

  // Verify Drive access
  try {
    const folder = DriveApp.getFolderById(INVOICES_FOLDER_ID);
    const templateFile = DriveApp.getFileById(DOC_TEMPLATE_ID);
    
    // Copy the Doc Template and open it
    const copy = templateFile.makeCopy('Invoice - ' + customer + ' - ' + invId, folder);
    const doc  = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    // Set template placeholders
    const replacements = {
      '{{INVOICE_ID}}': invId,
      '{{DATE}}':       Utilities.formatDate(new Date(), 'UTC', 'MMMM dd, yyyy'),
      '{{CUSTOMER}}':   customer,
      '{{STATE}}':      state,
      '{{SUBTOTAL}}':   '$' + subtotal.toFixed(2),
      '{{TAX_RATE}}':   (taxRate * 100).toFixed(2) + '%',
      '{{TAX_AMOUNT}}': '$' + taxAmt.toFixed(2),
      '{{TOTAL}}':      '$' + total.toFixed(2)
    };

    // Perform replacements in the document body
    Object.entries(replacements).forEach(([k, v]) => body.replaceText(k, v));
    doc.saveAndClose();

    // Export the Google Doc copy to PDF
    const pdfBlob = copy.getAs('application/pdf');
    pdfBlob.setName(invId + '.pdf');

    // Send the email via Gmail API
    GmailApp.sendEmail(
      email,
      'Invoice ' + invId + ' from Nexoria Commerce Inc.',
      'Dear Accounts Payable Team at ' + customer + ',\n\nPlease find attached your invoice ' + invId + ' for order conversion.\n\nThank you for your business,\nNexoria Commerce Inc.',
      {
        attachments: [pdfBlob],
        name: 'Nexoria Commerce Inc.',
        replyTo: 'billing@nexoriacommerce.com'
      }
    );

    // Log the transaction to Pipeline tracking worksheet
    if (logSht) {
      logSht.appendRow([invId, quoteId, customer, total.toFixed(2), 'Sent', new Date()]);
    }

    // Flag the row stage as processed
    qSht.getRange(row, 6).setValue('Invoiced');

    SpreadsheetApp.getUi().alert('Success! Invoice ' + invId + ' generated, saved to Drive, and emailed to ' + email);

  } catch (err) {
    SpreadsheetApp.getUi().alert('Execution Error: ' + err.toString());
  }
}
