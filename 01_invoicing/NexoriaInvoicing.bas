Attribute VB_Name = "NexoriaInvoicing"
' ==============================================================================
'   Nexoria Commerce Inc. — VBA Invoicing Automation
'   Finance Analyst Portfolio | Phase 1 | Module: xlsx-official
' ==============================================================================
'   Automates the generation of customer PDF invoices from the Orders worksheet.
'   Integrates with a dynamic TaxRates lookup sheet and logs all output files.
'
'   Setup:
'     1. Open NexoriaInvoicing.xlsx
'     2. Save the workbook as Excel Macro-Enabled Workbook (*.xlsm)
'     3. Press Alt+F11 to open the VBA Editor
'     4. Click File -> Import File... -> Select NexoriaInvoicing.bas
'     5. Create an "Invoices\" folder next to the workbook
'     6. Select the "Orders" sheet and run the GenerateInvoices macro
' ==============================================================================

Sub GenerateInvoices()
    Dim wsO As Worksheet:   Set wsO = ThisWorkbook.Sheets("Orders")
    Dim wsT As Worksheet:   Set wsT = ThisWorkbook.Sheets("InvoiceTemplate")
    Dim wsL As Worksheet:   Set wsL = ThisWorkbook.Sheets("Log")
    Dim wsI As Worksheet
    
    Dim lastRow As Long:    lastRow = wsO.Cells(wsO.Rows.Count, "A").End(xlUp).Row
    Dim i As Long, j As Long
    Dim currentOrderId As String
    Dim invoiceNum As String, pdfPath As String
    Dim folderPath As String: folderPath = ThisWorkbook.Path & "\Invoices\"
    
    ' Optimize Excel performance during batch run
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    ' Create Invoices directory next to workbook if absent
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    
    ' Process unique orders sequentially
    Dim orderCount As Long: orderCount = 0
    i = 2
    Do While i <= lastRow
        currentOrderId = wsO.Cells(i, "A").Value
        
        ' Skip if row is empty or order already processed
        If currentOrderId = "" Then
            i = i + 1
            GoTo ContinueLoop
        End If
        
        If wsO.Cells(i, "L").Value = "Invoiced" Then
            i = i + 1
            GoTo ContinueLoop
        End If
        
        ' Generate invoice ID based on Order Ref
        invoiceNum = "INV-" & Replace(currentOrderId, "ORD-", "")
        
        ' Copy template to a new temporary worksheet
        wsT.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set wsI = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        wsI.Name = invoiceNum
        
        ' Populate metadata and address details
        With wsI
            .Range("B5").Value = invoiceNum
            .Range("B6").Value = currentOrderId
            .Range("B7").Value = wsO.Cells(i, "M").Value ' Order Date (Col M)
            .Range("B8").Value = wsO.Cells(i, "N").Value ' Due Date (Col N)
            
            .Range("F5").Value = "PAID"
            .Range("F6").Value = wsO.Cells(i, "O").Value ' Payment Terms (Col O)
            .Range("F7").Value = wsO.Cells(i, "K").Value ' Carrier (Col K)
            .Range("F8").Value = wsO.Cells(i, "P").Value ' Tracking # (Col P)
            
            ' Populate Bill To details (writes to cell A10 in template)
            .Range("A10").Value = wsO.Cells(i, "C").Value & vbCrLf & _
                                  "Attn: Accounts Payable" & vbCrLf & _
                                  wsO.Cells(i, "Q").Value & vbCrLf & _
                                  wsO.Cells(i, "R").Value & ", " & wsO.Cells(i, "D").Value & " " & wsO.Cells(i, "S").Value
        End With
        
        ' Populate Line Items (handles multiple products under same Order ID)
        Dim itemRow As Long: itemRow = 12
        Dim shippingCost As Double: shippingCost = 0
        Dim stateCode As String: stateCode = wsO.Cells(i, "D").Value
        
        j = i
        Do While j <= lastRow And wsO.Cells(j, "A").Value = currentOrderId
            ' Template has 15 rows (12 to 26)
            If itemRow <= 26 Then
                wsI.Cells(itemRow, "A").Value = wsO.Cells(j, "E").Value ' SKU (Col E)
                wsI.Cells(itemRow, "B").Value = wsO.Cells(j, "F").Value ' Product Name (Col F)
                wsI.Cells(itemRow, "C").Value = wsO.Cells(j, "T").Value ' Category (Col T)
                wsI.Cells(itemRow, "D").Value = wsO.Cells(j, "G").Value ' Quantity (Col G)
                wsI.Cells(itemRow, "E").Value = wsO.Cells(j, "H").Value ' Unit Price (Col H)
                wsI.Cells(itemRow, "G").Value = IIf(wsO.Cells(j, "U").Value = 1, "Yes", "No") ' Taxable (Col U)
                itemRow = itemRow + 1
            End If
            
            shippingCost = wsO.Cells(j, "V").Value ' Shipping Cost (Col V)
            
            ' Mark row as invoiced
            wsO.Cells(j, "L").Value = "Invoiced"
            j = j + 1
        Loop
        
        ' Set tax rate and shipping cost in totals block
        wsI.Range("G30").Value = GetTaxRate(stateCode)
        wsI.Range("G32").Value = shippingCost
        
        ' Force sheet calculation to update formula values before export
        wsI.Calculate
        
        ' Export invoice to PDF
        pdfPath = folderPath & invoiceNum & ".pdf"
        wsI.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
                                Quality:=xlQualityStandard, OpenAfterPublish:=False
        
        ' Log details to run history sheet
        Dim logRow As Long: logRow = wsL.Cells(wsL.Rows.Count, "A").End(xlUp).Row + 1
        wsL.Cells(logRow, 1).Value = invoiceNum
        wsL.Cells(logRow, 2).Value = wsO.Cells(i, "C").Value
        wsL.Cells(logRow, 3).Value = Now
        wsL.Cells(logRow, 4).Value = pdfPath
        
        ' Clean up temporary sheet
        wsI.Delete
        orderCount = orderCount + 1
        
        ' Advance index to next unique Order ID
        i = j
        GoTo SkipIncrement
        
ContinueLoop:
        i = i + 1
SkipIncrement:
    Loop
    
    ' Restore Excel settings
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Execution Complete!" & vbCrLf & _
           "Successfully processed " & orderCount & " unique orders and generated PDFs.", _
           vbInformation, "Nexoria Finance Automation"
End Sub

' ── Helper Function for Sales Tax Rate Lookup ─────────────────────────
Function GetTaxRate(state As String) As Double
    Dim rng As Range: Set rng = ThisWorkbook.Sheets("TaxRates").Range("A:B")
    Dim v As Variant: v = Application.VLookup(state, rng, 2, False)
    GetTaxRate = IIf(IsError(v), 0, v)
End Function
