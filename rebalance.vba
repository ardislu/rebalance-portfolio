Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim cashRow As Integer
    cashRow = ActiveWorkbook.Sheets("Data").Range("B1").Value

    ' Marlett checkbox instead of form checkbox for performance
    If Not Intersect(Target, Range("I7:I" & cashRow - 1)) Is Nothing Then
            Cancel = True 'Prevent going into Edit Mode
            Target.Font.Name = "Marlett"
            Target.HorizontalAlignment = xlCenter
                If Target = vbNullString Then
                    Target = "a"
                Else
                    Target = vbNullString
                End If
    End If
End Sub

Private Sub DeleteAllStocks()
    Dim ws As Worksheet
    Dim sShape As Shape
    Dim shapeRow As Integer
    Dim cashRow As Integer
    
    Set ws = Sheets("Data")
    cashRow = ws.Range("B1").Value
    
    Set ws = Sheets("Rebalance")
    For Each sShape In ws.Shapes
        shapeRow = sShape.TopLeftCell.Row
        If shapeRow > 6 And shapeRow < cashRow Then
            sShape.Delete
        End If
    Next
    
    ws.Rows("7:" & cashRow - 1).Delete
End Sub

Public Sub AddStock(symbol As String, quantity As Double, price As Currency)
    Dim dataWS As Worksheet
    Dim rebalanceWS As Worksheet
    Dim cashRow As Integer
    Dim cellRefString As String
    Dim cellRef As Range
    Dim btn As Button
    
    Set dataWS = Sheets("Data")
    cashRow = dataWS.Range("B1").Value
    
    Set rebalanceWS = Sheets("Rebalance")
    cellRefString = "A" & cashRow
    rebalanceWS.Range(cellRefString).EntireRow.Insert ' Insert new row above cash
    
    ' Note: with new row inserted, old cashRow becomes new stock row
    
    ' Add delete row button
    cellRefString = "B" & cashRow
    Set cellRef = rebalanceWS.Range(cellRefString)
    Set btn = rebalanceWS.Buttons.Add(cellRef.Left, cellRef.Top, cellRef.Width, cellRef.Height)
    With btn
      .OnAction = "DeleteRowButton_Click"
      .Caption = "X"
      .Font.Bold = True
      .Name = "Btn" & cellRefString
    End With
    
    ' Hardcode given values
    cellRefString = "C" & cashRow & ":E" & cashRow
    rebalanceWS.Range(cellRefString).Value = Array(symbol, quantity, price)
    
    ' Set all input formulas
    cellRefString = "F" & cashRow & ":G" & cashRow
    rebalanceWS.Range(cellRefString).Formula = Array("=D" & cashRow & "*E" & cashRow, _
        "=F" & cashRow & "/$F$" & cashRow + 2)
    
    ' Set all optimal group formulas
    cellRefString = "K" & cashRow & ":N" & cashRow
    rebalanceWS.Range(cellRefString).Formula = Array("=$F$" & cashRow + 2 & "*($H" & cashRow & _
        "-$G" & cashRow & ")/$E" & cashRow, "=$D" & cashRow & "+K" & cashRow, "=$E" & cashRow & _
        "*L" & cashRow, "=M" & cashRow & "/$F$" & cashRow + 2)
    
    ' Set simple rounding group formulas
    cellRefString = "P" & cashRow & ":S" & cashRow
    rebalanceWS.Range(cellRefString).Formula = Array("=IF(ISBLANK(I" & cashRow & "),ROUND(K" & cashRow & _
        ",0),K" & cashRow & ")", "=$D" & cashRow & "+P" & cashRow, "=$E" & cashRow & "*Q" & cashRow, _
        "=R" & cashRow & "/$F$" & cashRow + 2)
        
    ' Set solver group formulas
    cellRefString = "V" & cashRow & ":X" & cashRow
    rebalanceWS.Range(cellRefString).Formula = Array("=$D" & cashRow & "+U" & cashRow, "=$E" & cashRow & _
        "*V" & cashRow, "=W" & cashRow & "/$F$" & cashRow + 2)
End Sub

Public Sub AddTotalFormulas()
    Dim dataWS As Worksheet
    Dim rebalanceWS As Worksheet
    Dim cashRow As Integer
    Dim cellRefString As String
    Dim cellRef As Range
    
    Set dataWS = Sheets("Data")
    cashRow = dataWS.Range("B1").Value
    
    Set rebalanceWS = Sheets("Rebalance")
    
    ' Set input group
    cellRefString = "F" & cashRow + 1 & ":H" & cashRow + 1
    rebalanceWS.Range(cellRefString).Formula = Array("=SUM(F7:F" & cashRow & ")", "=SUM(G7:G" & cashRow & ")", _
        "=SUM(H7:H" & cashRow & ")")
    
    ' Fix sumproducts
    rebalanceWS.Range("K" & cashRow).Formula = "=SUMPRODUCT($E7:$E" & cashRow - 1 & ",K7:K" & cashRow - 1 & ")*-1"
    rebalanceWS.Range("P" & cashRow).Formula = "=SUMPRODUCT($E7:$E" & cashRow - 1 & ",P7:P" & cashRow - 1 & ")*-1"
    rebalanceWS.Range("U" & cashRow).Formula = "=SUMPRODUCT($E7:$E" & cashRow - 1 & ",U7:U" & cashRow - 1 & ")*-1"
    
    ' Fix portfolio drifts
    rebalanceWS.Range("N" & cashRow + 4).FormulaArray = "=SUM(ABS(N7:N" & cashRow & "-$H7:$H" & cashRow & "))"
    rebalanceWS.Range("S" & cashRow + 4).FormulaArray = "=SUM(ABS(S7:S" & cashRow & "-$H7:$H" & cashRow & "))"
    rebalanceWS.Range("X" & cashRow + 4).FormulaArray = "=SUM(ABS(X7:X" & cashRow & "-$H7:$H" & cashRow & "))"
End Sub

Public Sub ResetFormatting()
    Dim dataWS As Worksheet
    Dim rebalanceWS As Worksheet
    Dim cashRow As Integer
    Dim dataCells As Range
    Dim quantityCells As Range
    Dim currencyCells As Range
    Dim percentageCells As Range
    Dim highlightCells As Range
    
    Set dataWS = Sheets("Data")
    cashRow = dataWS.Range("B1").Value
    
    Set rebalanceWS = Sheets("Rebalance")
    Set dataCells = rebalanceWS.Range("A7:" & "X" & cashRow)
    Set quantityCells = rebalanceWS.Range("D7:D" & cashRow - 1 & ", K7:L" & cashRow - 1 & _
        ", P7:Q" & cashRow - 1 & ", U7:V" & cashRow - 1)
    Set currencyCells = rebalanceWS.Range("E7:F" & cashRow - 1 & ", M7:M" & cashRow - 1 & _
        ", R7:R" & cashRow - 1 & ", W7:W" & cashRow - 1)
    Set percentageCells = rebalanceWS.Range("G7:H" & cashRow - 1 & ", N7:N" & cashRow - 1 & _
        ", S7:S" & cashRow - 1 & ", X7:X" & cashRow - 1)
    Set highlightCells = rebalanceWS.Range("D7:E" & cashRow - 1 & ", H7:H" & cashRow - 1)
    
    With dataCells
        .Font.Bold = False
        .HorizontalAlignment = xlLeft
    End With
    
    With quantityCells
        .NumberFormat = "0.0000"
    End With
    
    With currencyCells
        .NumberFormat = "$0.00"
    End With
    
    With percentageCells
        .NumberFormat = "0.00%"
    End With
    
    With highlightCells
        .Interior.Color = RGB(255, 255, 0)
    End With
End Sub

Private Sub ImportButton_Click()
    ' For file parsing
    Dim fd As Office.FileDialog
    Dim filePath As String
    Dim rowNumber As Integer
    Dim rowText As String
    Dim rowContent() As String
    ' Outputs
    Dim symbols() As String
    Dim quantities() As String
    Dim prices() As String
    Dim cash As Currency
    ' Writing outputs
    Dim ws As Worksheet
    Dim cashRow As Integer

    ' Select CSV file to open
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Select file to import"
        .Filters.Clear
        .Filters.Add "CSV (Comma delimited)", "*.csv"
        If .Show = True Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    ' Read file
    Open filePath For Input As #1
    
    rowNumber = 0
    ' Extend arrays to fit new data
    ReDim Preserve symbols(rowNumber)
    ReDim Preserve quantities(rowNumber)
    ReDim Preserve prices(rowNumber)
    
    Do Until EOF(1)
        Line Input #1, rowText
        rowContent = Split(rowText, ",")
        
        If rowNumber > 2 Then
            If rowContent(0) = """Cash & Cash Investments""" Then
                cash = CCur(Split(rowContent(6), Chr(34))(1))
                Exit Do
            End If
            
            ReDim Preserve symbols(rowNumber - 3)
            ReDim Preserve quantities(rowNumber - 3)
            ReDim Preserve prices(rowNumber - 3)
            
            symbols(rowNumber - 3) = rowContent(0)
            quantities(rowNumber - 3) = rowContent(2)
            prices(rowNumber - 3) = rowContent(3)
        End If
        
        rowNumber = rowNumber + 1
    Loop
    
    Close #1
    
    ' Writing outputs to sheet
    DeleteAllStocks
    
    Dim i As Long
    Dim symbol As String
    Dim quantity As Double
    Dim price As Currency
    For i = LBound(symbols) To UBound(symbols)
        symbol = Split(symbols(i), Chr(34))(1)
        quantity = CDbl(Split(quantities(i), Chr(34))(1))
        price = CCur(Split(prices(i), Chr(34))(1))
        AddStock symbol, quantity, price
    Next i
    
    Dim dataWS As Worksheet
    Dim rebalanceWS As Worksheet
    Set dataWS = Sheets("Data")
    Set rebalanceWS = Sheets("Rebalance")
    
    cashRow = dataWS.Range("B1").Value
    rebalanceWS.Range("F" & cashRow).Value = cash
    
    AddTotalFormulas
    ResetFormatting
End Sub

Private Sub AddStockButton_Click()
    AddStockForm.Show
End Sub

Private Sub RefreshPrices_Click()
    Dim mbResult As Integer
    mbResult = MsgBox("Warning: each stock will require 1 second of loading (API limitation). Continue?", vbYesNo)
    Select Case mbResult
        Case vbYes
        Case vbNo
            Exit Sub
    End Select

    ' Get data
    Dim dataWS As Worksheet
    Dim rebalanceWS As Worksheet
    Dim symbolsRange As Range
    Dim symbols As Variant
    Dim prices() As Currency
    Dim symbol As String
    Dim APIKey As String
    Dim cashRow As Integer
    
    Set dataWS = Sheets("Data")
    cashRow = dataWS.Range("B1").Value
    
    ' Max 5 calls/minute
    ' https://www.alphavantage.co/premium/
    If cashRow - 7 > 4 Then
        MsgBox "More than 4 stocks detected, quitting..." & vbLf & _
            "Source: https://www.alphavantage.co/premium/"
        Exit Sub
    End If
    
    Set rebalanceWS = Sheets("Rebalance")
    Set symbolsRange = rebalanceWS.Range("C7:" & "C" & cashRow - 1)
    
    ReDim symbols(1 To cashRow - 7)
    ReDim prices(1 To cashRow - 7)
    
    symbols = Application.WorksheetFunction.Transpose(symbolsRange.Value)
    
    ' Using Alpha Vantage API to get real-time stock data
    ' Reference: https://www.alphavantage.co/documentation/
    ' Ex. https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol=SCHE&interval=1min&apikey=EXAMPLE_XXXXXXXX&datatype=csv
    Dim query As String
    Dim req As New WinHttpRequest ' Tools > References > Microsoft WinHTTP Services
    APIKey = "EXAMPLE_XXXXXXXX"
    Dim i As Long
    For i = 1 To UBound(symbols)
        query = "https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&" & _
            "symbol=" & symbols(i) & "&interval=1min&apikey=" & APIKey & "&datatype=csv"
        req.Open "GET", query, False
        req.Send
        prices(i) = CCur(Split(CStr(Split(req.ResponseText, vbLf)(1)), ",")(4))
        Application.Wait Now + 0.00001  ' Must space each API call by 1 second
    Next i
    
    rebalanceWS.Range("E7:E" & cashRow - 1).Value = Application.WorksheetFunction.Transpose(prices)
End Sub

Private Sub SolverButton_Click()
    ' Get data
    Dim dataWS As Worksheet
    Dim rebalanceWS As Worksheet
    Dim cashRow As Integer
    Dim targetTotal As Double
    
    Set dataWS = Sheets("Data")
    cashRow = dataWS.Range("B1").Value
    
    Set rebalanceWS = Sheets("Rebalance")
    targetTotal = rebalanceWS.Range("H" & cashRow + 1).Value
    
    If targetTotal <> 1 Then
        MsgBox "Sum of target % must equal 100%!"
        Exit Sub
    End If

    rebalanceWS.Activate
    ' Tools > References > Solver
    SolverReset
    SolverOptions precision:=0.0001, _
        assumeNonNeg:=False, _
        multiStart:=False
    ' Minimize composite score by changing deltas
    SolverOK setCell:=Range("X" & cashRow + 5), _
        maxMinVal:=2, _
        byChange:=Range("U7:U" & cashRow - 1)
    ' Excess cash and composite score must be >= 0
    SolverAdd cellRef:=Range("X" & cashRow + 3), _
        relation:=3, _
        formulaText:=0
    SolverAdd cellRef:=Range("X" & cashRow + 5), _
        relation:=3, _
        formulaText:=0
    ' Non-fractional stocks must be integers
    Dim i As Long
    For i = 7 To cashRow - 1
        Range("U" & i).Value = Range("K" & i).Value ' Set initial delta to optimal to ballpark est.
        If Not Range("I" & i).Value = "a" Then
            SolverAdd cellRef:=Range("U" & i), _
                relation:=4
        End If
    Next i
    ' Set deltas to be >= -100 and <= 100 (required for evolutionary solve)
    ' SolverAdd cellRef:=Range("U7:U" & cashRow - 1), _
    '     relation:=3, _
    '     formulaText:=-100
    ' SolverAdd cellRef:=Range("U7:U" & cashRow - 1), _
    '    relation:=1, _
    '    formulaText:=100
    SolverSolve userFinish:=True
End Sub
