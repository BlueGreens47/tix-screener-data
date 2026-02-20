Attribute VB_Name = "CWE"
Option Explicit

'*****CWE Trading Signals*****

Sub cweGenerateTradingSignals()
    Dim wsData As Worksheet, wsDash As Worksheet, wsSignals As Worksheet, wsRpt As Worksheet
    Dim lastRow As Long, i As Long
    Dim compositeScore As Double, price As Double, volume As Double
       
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsDash = ThisWorkbook.Sheets("DashBoard")
    Set wsRpt = ThisWorkbook.Sheets("Reports")
    
    On Error Resume Next
    Set wsSignals = ThisWorkbook.Sheets("cweSignals")
    If wsSignals Is Nothing Then
        Set wsSignals = ThisWorkbook.Sheets.Add
        wsSignals.Name = "cweSignals"
    End If
    On Error GoTo 0
    
    ' Clear and setup signals sheet
    wsSignals.Cells.Clear
    With wsSignals
        .Range("A1").value = "Ticker"
        .Range("B1").value = "Signal"
        .Range("C1").value = "Strength"
        .Range("D1").value = "Price"
        .Range("E1").value = "Composite Score"
        .Range("F1").value = "RSI"
        .Range("G1").value = "MACD Diff"
        .Range("H1").value = "Trend"
        .Range("I1").value = "Timestamp"
        
        .Range("A1:I1").Font.Bold = True
        .Columns.AutoFit
    End With
    
    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).row
    If lastRow < 2 Then Exit Sub
    
    Dim signalRow As Long: signalRow = 2
    Dim currentTicker As String: currentTicker = ""
    
    For i = lastRow To 2 Step -1 ' Process from most recent to oldest
        If wsData.Cells(i, 7).value <> currentTicker Then ' New ticker
            currentTicker = wsData.Cells(i, 7).value
            
            compositeScore = Nz(wsData.Cells(i, 14).value, 0)
            price = wsData.Cells(i, 5).value
            volume = wsData.Cells(i, 6).value
            
            Dim rsi As Double: rsi = Nz(wsData.Cells(i, 8).value, 50)
            Dim macd As Double: macd = Nz(wsData.Cells(i, 9).value, 0)
            Dim macdSignal As Double: macdSignal = Nz(wsData.Cells(i, 10).value, 0)
            Dim priceVsMA As Double: priceVsMA = Nz(wsData.Cells(i, 11).value, 0)
            Dim bbPosition As Double: bbPosition = Nz(wsData.Cells(i, 12).value, 0.5)
            Dim volumeSpike As Double: volumeSpike = Nz(wsData.Cells(i, 13).value, 1)
            
            Dim signal As String, strength As String
            Call GenerateSignal(compositeScore, rsi, macd, macdSignal, priceVsMA, bbPosition, volumeSpike, signal, strength)
            
            If signal <> "HOLD" Then ' Only record actionable signals
                wsSignals.Cells(signalRow, 1).value = currentTicker
                wsSignals.Cells(signalRow, 2).value = signal
                wsSignals.Cells(signalRow, 3).value = strength
                wsSignals.Cells(signalRow, 4).value = price
                wsSignals.Cells(signalRow, 5).value = compositeScore
                wsSignals.Cells(signalRow, 6).value = rsi
                wsSignals.Cells(signalRow, 7).value = macd - macdSignal
                wsSignals.Cells(signalRow, 8).value = priceVsMA
                wsSignals.Cells(signalRow, 9).value = wsDash.Range("H5").value
                
                signalRow = signalRow + 1
            End If
        End If
    Next i
    
    ' Apply conditional formatting
    With wsSignals.Range("B2:B" & signalRow)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""BUY"""
        .FormatConditions(1).Interior.Color = RGB(198, 239, 206) ' Green
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""SELL"""
        .FormatConditions(2).Interior.Color = RGB(255, 199, 206) ' Red
    End With
    
    If signalRow > 2 Then
    Dim nextlogrow As Integer
         With wsRpt
            .Range("A4:O" & lastRow).ClearContents
            .Range("A4:A" & signalRow + 3).value = wsDash.Range("H5").value
            .Range("B4:B" & signalRow + 3).value = wsSignals.Range("A2:A" & signalRow).value
           
            nextlogrow = FindNextEmptyRow(wsRpt, "B", 4)
            .Range("A" & nextlogrow & ":O100").ClearContents
            
            Call ReportToDashOptimized
        End With
    End If
    
    msgbox "Trading signals generated: " & (signalRow - 2) & " actionable signals on Dash", vbInformation
End Sub

Sub cweFilteredDataFromBackupWithArrays()

    Dim wsDashboard As Worksheet, wsTJX As Worksheet, wsDash As Worksheet
    Dim wsBackupAll As Worksheet, wsData As Worksheet
    Dim maxPrice As Double, minPrice As Double, endDate As Date
    Dim lastRowTJX As Long, lastRowBackup As Long, lastRowData As Long
    Dim i As Long, j As Long, k As Long, outputRow As Long
    Dim ticker As String, tickerPrice As Double
    Dim tjxData As Variant, backupData As Variant
    Dim outputData() As Variant
    Dim tickerRecords() As Variant
    Dim recordCount As Long
    Dim tableName As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim startTime As Double
    
    DoEvents
    If gStopMacro Then
        msgbox "...E-Stopped!", vbInformation
        Exit Sub
    End If

    startTime = Timer
    
    ' OPTIMIZATION: Disable all unnecessary features
    Application.ScreenUpdating = False
    Application.CALCULATION = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Set&Prep worksheets
    With ThisWorkbook
        ' Get the table name from the dropdown
        tableName = .Sheets("DashBoard").Range("A1").value
        
        ' Find the worksheet that contains this table
        Set wsTJX = Nothing
        For Each ws In .Worksheets
            On Error Resume Next
            Set tbl = ws.ListObjects(tableName)
            On Error GoTo 0
            If Not tbl Is Nothing Then
                Set wsTJX = ws
                Exit For
            End If
        Next ws
        
        ' Check if table was found
        If wsTJX Is Nothing Then
            msgbox "Table '" & tableName & "' not found!", vbCritical
            Exit Sub
        End If
         ' Set worksheet references
        Set wsDash = .Sheets("DashBoard")
        Set wsData = .Sheets("Data")
        Set wsTJX = .Sheets("TJX")
        Set wsBackupAll = .Sheets("BackupAll")
      
    End With
    
    Call ClearAllFilters
    wsData.Range("A2:V" & wsData.Rows.count).ClearContents
     
    ' Get parameters
    minScore = wsDash.Range("W5").value
    
    maxPrice = wsDash.Range("Y5").value
    minPrice = wsTJX.Range("Y6").value
    endDate = wsDash.Range("H5").value

    Dim skipPrompt As Boolean
    skipPrompt = pubNotice Or perfTest
    
    If Not skipPrompt Then
        If Not ConfirmProcessing() Then
            If Not GetUserInputs(minScore, minPrice, maxPrice, endDate) Then
                Exit Sub
            End If
        End If
    End If
       
    ' Find last rows
    lastRowTJX = wsTJX.Cells(wsTJX.Rows.count, "A").End(xlUp).row
    lastRowBackup = wsBackupAll.Cells(wsBackupAll.Rows.count, "A").End(xlUp).row
    lastRowData = wsData.Cells(wsData.Rows.count, "A").End(xlUp).row
    
    ' Load all data into arrays
    tjxData = wsTJX.Range("A3:D" & lastRowTJX).value
    backupData = wsBackupAll.Range("A2:G" & lastRowBackup).value
    
    ' Start appending after existing data
    If lastRowData = 1 And wsData.Range("A1").value = "" Then
        outputRow = 2
    Else
        outputRow = lastRowData + 1
    End If
    
    ' Prepare output array (max size estimate)
    ReDim outputData(1 To UBound(tjxData, 1) * 50, 1 To 7)
    Dim outputIndex As Long
    outputIndex = 0
    
    ' Loop through each ticker in TJX table
    For i = 1 To UBound(tjxData, 1)
        ticker = tjxData(i, 1) ' Column A
        tickerPrice = tjxData(i, 4) ' Column D
        
        ' Check if price is within range
        If tickerPrice >= minPrice And tickerPrice <= maxPrice Then
            
            ' Collect matching records from BackupAll into temp array
            ReDim tickerRecords(1 To lastRowBackup, 1 To 7)
            recordCount = 0
            
            For j = 1 To UBound(backupData, 1)
                If backupData(j, 7) = ticker Then ' Column 7 is Ticker
                    If backupData(j, 1) <= endDate Then ' Column 1 is Date
                        recordCount = recordCount + 1
                        tickerRecords(recordCount, 1) = backupData(j, 1)
                        tickerRecords(recordCount, 2) = backupData(j, 2)
                        tickerRecords(recordCount, 3) = backupData(j, 3)
                        tickerRecords(recordCount, 4) = backupData(j, 4)
                        tickerRecords(recordCount, 5) = backupData(j, 5)
                        tickerRecords(recordCount, 6) = backupData(j, 6)
                        tickerRecords(recordCount, 7) = backupData(j, 7)
                    End If
                End If
            Next j
            
            ' Take the last 50 records (most recent)
            If recordCount > 0 Then
                Dim startIdx As Long
                startIdx = Application.WorksheetFunction.max(1, recordCount - 39) '40 Records
                
                For j = startIdx To recordCount
                    outputIndex = outputIndex + 1
                    outputData(outputIndex, 1) = tickerRecords(j, 1)
                    outputData(outputIndex, 2) = tickerRecords(j, 2)
                    outputData(outputIndex, 3) = tickerRecords(j, 3)
                    outputData(outputIndex, 4) = tickerRecords(j, 4)
                    outputData(outputIndex, 5) = tickerRecords(j, 5)
                    outputData(outputIndex, 6) = tickerRecords(j, 6)
                    outputData(outputIndex, 7) = tickerRecords(j, 7)
                Next j
            End If
        End If
    Next i
    
    ' Write all data at once if we have any
    If outputIndex > 0 Then
        wsData.Range("A" & outputRow).Resize(outputIndex, 7).value = outputData
        
    End If
    
    Application.ScreenUpdating = True
    Application.CALCULATION = xlCalculationAutomatic

    Call UpdateSystemWithATR_Complete
    Call CalculateIndicators  ' Fixed: was CalculateEnhancedIndicators (wrong column layout)
    Call cweGenerateTradingSignals

    Call DisplayCompletionMessage(startTime)
End Sub


Sub GenerateSignal(compositeScore As Double, rsi As Double, macd As Double, macdSignal As Double, _
                  priceVsMA As Double, bbPosition As Double, volumeSpike As Double, _
                  ByRef signal As String, ByRef strength As String)
    
    Dim macdDiff As Double: macdDiff = macd - macdSignal
    
    ' Strong Buy conditions
    If compositeScore >= 3 And rsi < 35 And macdDiff > 0 And priceVsMA < -2 And bbPosition < 0.3 And volumeSpike > 1.5 Then
        signal = "BUY": strength = "STRONG"
        Exit Sub
    End If
    
    ' Strong Sell conditions
    If compositeScore <= -3 And rsi > 65 And macdDiff < 0 And priceVsMA > 2 And bbPosition > 0.7 And volumeSpike > 1.5 Then
        signal = "SELL": strength = "STRONG"
        Exit Sub
    End If
    
    ' Moderate Buy conditions
    If compositeScore >= 2 And rsi < 40 And macdDiff > 0 And priceVsMA < 0 And bbPosition < 0.4 Then
        signal = "BUY": strength = "MODERATE"
        Exit Sub
    End If
    
    ' Moderate Sell conditions
    If compositeScore <= -2 And rsi > 60 And macdDiff < 0 And priceVsMA > 0 And bbPosition > 0.6 Then
        signal = "SELL": strength = "MODERATE"
        Exit Sub
    End If
    
    ' Weak Buy conditions
    If compositeScore >= 1.5 And ((rsi < 45 And macdDiff > 0) Or (priceVsMA < -1 And bbPosition < 0.5)) Then
        signal = "BUY": strength = "WEAK"
        Exit Sub
    End If
    
    ' Weak Sell conditions
    If compositeScore <= -1.5 And ((rsi > 55 And macdDiff < 0) Or (priceVsMA > 1 And bbPosition > 0.5)) Then
        signal = "SELL": strength = "WEAK"
        Exit Sub
    End If
    
    signal = "HOLD"
    strength = "NONE"
End Sub

