Option Explicit

Sub BookDeal_Click()

    ' TRADE BOOKING & HISTORY MANAGEMENT

    Dim wsPricer As Worksheet
    Dim wsHist As Worksheet
    Dim NextRow As Long
    Dim Counterparty As String
    Dim Volume As Double
    Dim Price As Double
    Dim Margin As Double
    Dim TotalPnL As Double
    Dim UserResponse As Integer
    
    ' 1. Setup Worksheets
    Set wsPricer = ThisWorkbook.Sheets("PRICER")
    Set wsHist = ThisWorkbook.Sheets("DB_TRADES")
    
    ' 2. Data Retrieval (Simulating Deal Entry)
    On Error Resume Next
    Counterparty = InputBox("Counterparty Name:", "New Trade Entry")
    If Counterparty = "" Then Exit Sub ' Cancel if empty
    
    Volume = InputBox("Volume (Metric Tons):", "Quantity")
    If Volume = 0 Then Exit Sub ' Cancel if 0
    On Error GoTo 0
    
    ' Retrieve calculated data from the sheet
    Price = wsPricer.Range("B12").Value
    Margin = wsPricer.Range("B11").Value
    TotalPnL = Margin * Volume
    
    ' 3. Trader Confirmation
    UserResponse = MsgBox("Confirm Trade Execution?" & vbNewLine & _
                          "Counterparty : " & Counterparty & vbNewLine & _
                          "Volume : " & Volume & " MT" & vbNewLine & _
                          "P&L : $" & Format(TotalPnL, "#,##0.00"), vbYesNo + vbQuestion, "CONFIRMATION")
    
    If UserResponse = vbNo Then Exit Sub
    
    ' 4. SAVE TO HISTORY (LOG)
    NextRow = wsHist.Cells(wsHist.Rows.Count, "A").End(xlUp).Row + 1
    
    With wsHist
        .Cells(NextRow, 1).Value = "TRD-" & NextRow ' Generates unique ID
        .Cells(NextRow, 2).Value = Now
        .Cells(NextRow, 3).Value = Counterparty
        .Cells(NextRow, 4).Value = Volume
        .Cells(NextRow, 5).Value = Price
        .Cells(NextRow, 6).Value = Margin
        .Cells(NextRow, 7).Value = TotalPnL
        
        ' Formatting ($)
        .Range("E" & NextRow & ":G" & NextRow).Style = "Currency"
    End With
    
    MsgBox "Trade successfully booked in DB_TRADES.", vbInformation

End Sub
