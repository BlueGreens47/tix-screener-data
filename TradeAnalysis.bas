Attribute VB_Name = "TradeAnalysis"
' Trading Performance Analyzer - Excel VBA
' This code should be placed in a VBA module in Excel
' Make sure to enable macros and add references to Microsoft Forms 2.0 Object Library

Option Explicit

' Global variables
Public Const TRADES_SHEET = "Trades"
Public Const METRICS_SHEET = "Metrics"
Public Const CHARTS_SHEET = "Charts"

Sub OneStepPerformance()

    Dim wsDash As Worksheet
    Dim wsRpt As Worksheet
    Dim wsTrades As Worksheet
    Dim wsTLog As Worksheet
    Dim wsPerForm As Worksheet
    Dim i As Long
    Dim startTime As Double
    Dim lastRow As Long
    Dim tstPeriod As Variant
    
    perfTest = True
    startTime = Timer
    
    On Error Resume Next
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
        
        Set wsDash = ThisWorkbook.Sheets("DashBoard")
        Set wsRpt = ThisWorkbook.Sheets("Reports")
        Set wsTLog = ThisWorkbook.Sheets("TRADE LOG")
        Set wsTrades = ThisWorkbook.Sheets("Trades")
        Set wsPerForm = ThisWorkbook.Sheets("PERFORMANCE")
        
        wsPerForm.Range("A12:X300").ClearContents
        wsTrades.Range("B2:L300").ClearContents
        
        Call CreateWorksheets
        
        Call FilterAndReport
        
        lastRow = wsDash.Cells(wsDash.Rows.count, "A").End(xlUp).row
        
        If lastRow <= 8 Then
             MsgBox "Not enough rows in Report..exiting!"
            Exit Sub
       
        End If
        
        Application.ScreenUpdating = True
         Sheets("DashBoard").Select
         If Not ConfirmProcessing Then Exit Sub
       ' Application.ScreenUpdating = False
        
        Call SetupTradeLog(wsTLog, lastRow)
        
        tstPeriod = Array(7, 14, 28, 56, 84, 112) '35, 42, 49, 56, 63, 70, 77, 84)
        For i = 0 To UBound(tstPeriod)
            wsTLog.Range("N2").value = tstPeriod(i)
            Application.Calculate
            Call UpdateTradePerformance
        Next i
       
        Call ApplyTradesIDs
        Call AnalyzeTrades
        
        perfTest = False
        Sheets("Charts").Select
        
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    DisplayCompletionMessage (startTime)

End Sub

' Create necessary worksheets
Sub CreateWorksheets()
    Dim ws As Worksheet
    Dim wsNames As Variant
    Dim i As Integer
    
    wsNames = Array(TRADES_SHEET, METRICS_SHEET, CHARTS_SHEET)
    
    For i = 0 To UBound(wsNames)
        Set ws = Nothing
        
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(wsNames(i))
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        On Error GoTo 0
        
       
            Set ws = ThisWorkbook.Worksheets.Add
            ws.Name = wsNames(i)
      
    Next i
    Call SetupTradesTable
    
End Sub

' Set up the trades table structure
Sub SetupTradesTable()
    Dim ws As Worksheet
    Dim headers As Variant
    Dim i As Integer
    
    Set ws = ThisWorkbook.Worksheets(TRADES_SHEET)
    'ws.Range("B2:W300").ClearContents
    
    headers = Array("ID", "Group", "Entry Date", "Exit Date", "Setup", _
                   "Conviction", "VBA", "Market Regime", "Outcome", _
                   "P&L", "Risk Amount", "R-Multiple", _
                   "Rank", "Boll", "VolSpike", "Hull", "DMI", "MA", "MACD", "RSI", "S&DSS", "Candles", "S&R")
    
    ' Add headers
    For i = 1 To UBound(headers) + 1
        ws.Cells(1, i).value = headers(i - 1)
        'ws.Cells(1, i).Font.Bold = True
        ws.Cells(1, i).Interior.Color = RGB(200, 220, 240)
    Next i
    
    ws.Range("A2").Formula = "=RIGHT(B2,2)"
    ' Format columns
    ws.Columns("C:D").NumberFormat = "ddd dd mmm yyyy"
    ws.Columns("G").NumberFormat = "0.0"
    ws.Columns("J:K").NumberFormat = "$#,##0.00"
    ws.Columns("L").NumberFormat = "0.00"
    
    ' Auto-fit columns
    ws.Columns.AutoFit
    
    ' Add data validation for dropdowns
    'Call AddDataValidation
    
    ' Add buttons
    Call AddControlButtons
  
End Sub

' Add control buttons
Sub AddControlButtons()
    Dim ws As Worksheet
    Dim btn As Button
    
    Set ws = ThisWorkbook.Worksheets(TRADES_SHEET)
    
    ' Add Trade button
   ' Set btn = ws.Buttons.Add(10, 10, 100, 30)
   ' btn.Caption = "Add Trade"
   ' btn.OnAction = "AddTradeForm"
    
    ' Bulk Import button
    Set btn = ws.Buttons.Add(1100, 1, 75, 20)
    btn.Caption = "Update"
    btn.OnAction = "UpdateTradePerformance"
    
    ' Refresh Metrics button
    Set btn = ws.Buttons.Add(1200, 1, 75, 20)
    btn.Caption = "ANALYZE"
    btn.OnAction = "AnalyzeTrades"
End Sub


' Helper subroutine to setup Trade Log
Sub SetupTradeLog(wsTLog As Worksheet, lastRow As Long)
            
    ThisWorkbook.Sheets("TRADE LOG (2)").Range("A1:AF3").Copy
    With wsTLog
        .Range("A1").PasteSpecial Paste:=xlPasteFormulas
         Application.CutCopyMode = False
        .Range("B4:AF53").ClearContents
        .Range("B1:AF1").Copy
        .Range("B4:AF" & lastRow).PasteSpecial Paste:=xlPasteFormulas
        .Range("A4:A" & lastRow).Formula = "= ""Group "" & " & "row()-3"
        .Range("N2").value = 7
        Application.Calculate
        Application.CutCopyMode = False
    End With
End Sub

Sub LogPerformance() ' Records Performance

    On Error Resume Next 'GoTo ErrorHandler
   
    ' Variable declarations
    Dim ws As Worksheet, wsTLog As Worksheet, wsRpt As Worksheet, wsTrades As Worksheet
    Dim lastRow As Long, nextLogRow As Long, nextTradesRow As Long
    Dim currentGroup As String, groupNumber As Long
    Dim valueRange As Range, templateRange As Range
    Dim originalCalculation As XlCalculation
    
    ' Store original calculation mode and disable automatic calculation for performance
    originalCalculation = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Set worksheet references with error checking
    Set ws = GetWorksheet("PERFORMANCE")
    Set wsTLog = GetWorksheet("TRADE LOG")
    Set wsRpt = GetWorksheet("Reports")
    Set wsTrades = GetWorksheet("Trades")
   
    ' Get last row of data in Reports sheet
    lastRow = wsTLog.Range("C2").value
    If lastRow < 4 Then
        If Not perfTest Then MsgBox "No data found!", vbExclamation
        GoTo Cleanup
    End If
    
   ' Find next empty row in sheet, column, starting from row
    nextLogRow = FindNextEmptyRow(ws, "F", 12)
    
    ' Get current group info before incrementing
    currentGroup = CStr(ws.Range("A" & nextLogRow - 4).value)
    
     ' Copy performance data to next row
    Call SetupPerformance(ws, nextLogRow)
          
    DoEvents
    
    Application.CutCopyMode = False
    ' Increment group number and update
    groupNumber = ExtractAndIncrementGroupNumber(currentGroup)
    
    ws.Range("A" & nextLogRow).value = "Test_Group_" & groupNumber
    Application.CutCopyMode = False
     
    ' Copy Performance to Trades
    nextTradesRow = FindNextEmptyRow(wsTrades, "B", 2)
    Call setupTrades(wsTLog, wsTrades, lastRow, nextTradesRow)
       
    ' Navigate to the new row
    'ws.Activate
    'ws.Range("A" & nextLogRow).Select
    'Application.Calculate
   ' Application.Wait Now + TimeValue("00:00:01")
    
    'If Not perfTest Then MsgBox "Performance data copied successfully to row " & nextLogRow, vbInformation
    
Cleanup:
    ' Restore application settings
    Application.Calculation = originalCalculation
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.CutCopyMode = False
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in LogPerformance: " & Err.Description & " (Line: " & Erl & ")", vbExclamation
    GoTo Cleanup
End Sub

Sub SetupPerformance(ws As Worksheet, lastRow As Long)
    With ws
    If Not perfTest Then ws.Range("B2:L300").ClearContents
        .Range("A1:X4").Copy
        .Range("A" & lastRow & ":X" & lastRow + 3).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        .Range("A" & lastRow & ":X" & lastRow + 3).PasteSpecial xlPasteFormats
        
        '.Range("B" & lastRow & ":X" & lastRow + 2).value = .Range("B1:X3").value
        Application.CutCopyMode = False
    End With
End Sub

Sub setupTrades(wsTLog As Worksheet, wsTrades As Worksheet, count As Long, nextTradesRow As Long)
          
    'wsTrades.Range("B" & nextTradesRow & ":K" & nextTradesRow + count).value = wsTLog.Range("A4:J" & count + 4).value
    
    With wsTrades
    Dim lastRow As Long
    lastRow = wsTrades.Cells(wsTrades.Rows.count, "B").End(xlUp).row
        .Range("A2").Copy
        .Range("A3:A" & lastRow).PasteSpecial Paste:=xlPasteFormulas
    End With
    Application.CutCopyMode = False
  
End Sub

Sub ApplyTradesIDs()

    Dim wsTrades As Worksheet
    Dim lastRow As Long
    Dim formulas() As String
    Dim i As Long

    Set wsTrades = ThisWorkbook.Sheets("Trades")
    lastRow = wsTrades.Cells(wsTrades.Rows.count, "C").End(xlUp).row
 
    ' Preload formulas into Column A
    ReDim formulas(1 To lastRow - 1)
    For i = 2 To lastRow
        formulas(i - 1) = "=RIGHT(B" & i & ",2)"
    Next i
    ' Write all formulas to column A
    wsTrades.Range("A2:A" & lastRow).Formula = Application.Transpose(formulas)
    
    ' Preload formulas into R-multiple column
    ReDim formulas(1 To lastRow - 1)
    For i = 2 To lastRow
        formulas(i - 1) = "=IFERROR(J" & i & "/K" & i & ", """")"
    Next i
    ' Write all formulas to column L
    wsTrades.Range("L2:L" & lastRow).Formula = Application.Transpose(formulas)
    wsTrades.Columns.AutoFit
    
End Sub

' Analyze Trades
Sub AnalyzeTrades()
    'Application.ScreenUpdating = False
    Application.Calculate
    
    Call SetupTradesTable
    Call CreateMetricsDashboard
    
    'Call setupTradesPage
    
    Call CalculateAllMetrics
    Call CreateAllCharts
    
    Application.ScreenUpdating = True
    
    MsgBox "All metrics and charts refreshed!", vbInformation
End Sub

' Gets data from Trade Log
Sub UpdateTradePerformance()

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
   
    Dim wsTLog As Worksheet
    Dim wsTrades As Worksheet
    Dim lastRow As Long, lastSourceRow As Long, i As Long
    Dim wschart As Worksheets
        
    Set wsTLog = ThisWorkbook.Worksheets("TRADE LOG")
    Set wsTrades = ThisWorkbook.Worksheets("Trades")
    
    If Not perfTest Then wsTrades.Range("B2:V300").ClearContents
    
    lastSourceRow = wsTLog.Range("C2").value + 3
    lastRow = wsTrades.Cells(wsTrades.Rows.count, "B").End(xlUp).row

    ' Copy data and paste only values and formatting
   
    With wsTrades
        wsTLog.Range("A4:J" & lastSourceRow).Copy
        .Range("B" & lastRow + 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        
        wsTLog.Range("V4:AF" & lastSourceRow).Copy
        .Range("M" & lastRow + 1).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        lastRow = wsTrades.Cells(wsTrades.Rows.count, "B").End(xlUp).row
        
        For i = 2 To lastRow
            ' Calculate R-Multiple
            wsTrades.Cells(i, 12).Formula = "=IFERROR(J" & i & "/K" & i & ", """")"
        Next i
    
    End With
    
    Call LogPerformance
    
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    'If Not perfTest Then MsgBox "Done"
End Sub

' Main initialization subroutine
Sub InitializeTradingAnalyzer()
    Application.ScreenUpdating = False
    
    ' Create worksheets if they don't exist
    Call CreateWorksheets
    
    ' Set up the trades table structure
    Call SetupTradesTable
    
    ' Create the metrics dashboard
    Call CreateMetricsDashboard
    
    ' Add data
    Call UpdateTradePerformance
    Call LogPerformance
    
    ' Calculate initial metrics
    Call CalculateAllMetrics
    
    ' Create charts
    Call CreateAllCharts
    
    Application.ScreenUpdating = True
    
    MsgBox "Trading Performance Analyzer initialized successfully!" & vbCrLf & _
           "Use the 'Add Trade' button to add new trades." & vbCrLf & _
           "Metrics and charts will update automatically.", vbInformation
End Sub


' Create metrics dashboard
Sub CreateMetricsDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(METRICS_SHEET)
    
    ' Title
    ws.Range("A1").value = "Trading Performance Metrics Dashboard"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    
    ' Key Metrics section
    ws.Range("A3").value = "Key Performance Metrics"
    ws.Range("A3").Font.Size = 14
    ws.Range("A3").Font.Bold = True
    
    ' Metric labels and formulas
    Dim metrics As Variant
    metrics = Array( _
        Array("Total Trades", "=COUNTA(Trades!B:B)-1"), _
        Array("Win Rate (%)", "=COUNTIF(Trades!I:I,""Win"")/COUNTA(Trades!I:I)*100"), _
        Array("Total P&L", "=SUM(Trades!J:J)"), _
        Array("Profit Factor", "=SUMIF(Trades!I:I,""Win"",Trades!J:J)/ABS(SUMIF(Trades!I:I,""Loss"",Trades!J:J))"), _
        Array("Average Win", "=AVERAGEIF(Trades!I:I,""Win"",Trades!J:J)"), _
        Array("Average Loss", "=ABS(AVERAGEIF(Trades!I:I,""Loss"",Trades!J:J))"), _
        Array("Max Drawdown", "'=CalcMaxDrawdown()"), _
        Array("Expectancy", "'=CalcExpectancy()") _
    )
    
    Dim i As Integer
    For i = 0 To UBound(metrics)
        ws.Cells(5 + i, 1).value = metrics(i)(0)
        ws.Cells(5 + i, 1).Font.Bold = True
        ws.Cells(5 + i, 2).Formula = metrics(i)(1)
        ws.Cells(5 + i, 2).NumberFormat = "#,##0.00"
    Next i
    
    ' Format metrics
    ws.Range("B6").NumberFormat = "0.00%" ' Win Rate
    ws.Range("B7,B9,B10,B11,B12").NumberFormat = "#,##0.00"
    
    ' Analysis tables
    Call CreateAnalysisTables
    
    ws.Columns.AutoFit
End Sub

' Create analysis tables for setup, conviction, and market regime
Sub CreateAnalysisTables()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(METRICS_SHEET)
    
    ' Setup Analysis
    'ws.Range("D3").value = "Setup Analysis"
   ' ws.Range("D3").Font.Bold = True
   ' ws.Range("D4:G4").value = Array("Setup", "Win Rate", "Total P&L", "Trade Count")
   ' ws.Range("D4:G4").Font.Bold = True
    
    ' Market Regime Analysis
    ws.Range("I3").value = "Market Regime Analysis"
    ws.Range("I3").Font.Bold = True
    ws.Range("I4:L4").value = Array("Regime", "Win Rate", "Total P&L", "Trade Count")
    ws.Range("I4:L4").Font.Bold = True
    
    ' Conviction Analysis
    'ws.Range("D18").value = "Conviction Analysis"
    'ws.Range("D18").Font.Bold = True
    'ws.Range("D19:G19").value = Array("Conviction", "Win Rate", "Total P&L", "Trade Count")
    'ws.Range("D19:G19").Font.Bold = True
    
    ' vba Analysis
    ws.Range("I18").value = "VBA Analysis"
    ws.Range("I18").Font.Bold = True
    ws.Range("I19:L19").value = Array("VBA", "Win Rate", "Total P&L", "Trade Count")
    ws.Range("I19:L19").Font.Bold = True
    
    ws.Range("B:B,J:J").NumberFormat = "0.00_ ;[Red]-0.00 "
End Sub

' Calculate all metrics and update tables
Sub CalculateAllMetrics()
    'Call UpdateSetupAnalysis
    Call UpdateMarketRegimeAnalysis
    'Call UpdateConvictionAnalysis
    Call UpdateVBAAnalysis
    Call UpdateEquityCurve
End Sub

' Update setup analysis table
Sub UpdateSetupAnalysis()
    Dim ws As Worksheet
    Dim tradesWs As Worksheet
    Dim setupList As Collection
    Dim setup As Variant
    Dim row As Integer
    
    Set ws = ThisWorkbook.Worksheets(METRICS_SHEET)
    Set tradesWs = ThisWorkbook.Worksheets(TRADES_SHEET)
    Set setupList = GetUniqueValues(tradesWs.Range("E:E"))
    
    row = 5
    For Each setup In setupList
        If setup <> "Setup" And setup <> "" Then
            ws.Cells(row, 4).value = setup
            ws.Cells(row, 5).Formula = "=COUNTIFS(Trades!E:E,""" & setup & """,Trades!I:I,""Win"")/COUNTIF(Trades!E:E,""" & setup & """)*100"
            ws.Cells(row, 6).Formula = "=SUMIF(Trades!E:E,""" & setup & """,Trades!J:J)"
            ws.Cells(row, 7).Formula = "=COUNTIF(Trades!E:E,""" & setup & """)"
            row = row + 1
        End If
    Next setup
    
    ws.Range("E5:E" & row - 1).NumberFormat = "0.00%"
    ws.Range("F5:F" & row - 1).NumberFormat = "$#,##0.00"
End Sub

' Update market regime analysis table
Sub UpdateMarketRegimeAnalysis()
    Dim ws As Worksheet
    Dim tradesWs As Worksheet
    Dim regimeList As Collection
    Dim regime As Variant
    Dim row As Integer
    
    Set ws = ThisWorkbook.Worksheets(METRICS_SHEET)
    Set tradesWs = ThisWorkbook.Worksheets(TRADES_SHEET)
    Set regimeList = GetUniqueValues(tradesWs.Range("H:H"))
    
    row = 5
    For Each regime In regimeList
        If regime <> "Market Regime" And regime <> "" Then
            ws.Cells(row, 9).value = regime
            ws.Cells(row, 10).Formula = "=COUNTIFS(Trades!H:H,""" & regime & """,Trades!I:I,""Win"")/COUNTIF(Trades!H:H,""" & regime & """)*100"
            ws.Cells(row, 11).Formula = "=SUMIF(Trades!H:H,""" & regime & """,Trades!J:J)"
            ws.Cells(row, 12).Formula = "=COUNTIF(Trades!H:H,""" & regime & """)"
            row = row + 1
        End If
    Next regime
    
    ws.Range("J5:J" & row - 1).NumberFormat = "0.00%"
    ws.Range("K5:K" & row - 1).NumberFormat = "$#,##0.00"
End Sub

' Update conviction analysis table
Sub UpdateConvictionAnalysis()
    Dim ws As Worksheet
    Dim tradesWs As Worksheet
    Dim convictionList As Collection
    Dim conviction As Variant
    Dim row As Integer
    
    Set ws = ThisWorkbook.Worksheets(METRICS_SHEET)
    Set tradesWs = ThisWorkbook.Worksheets(TRADES_SHEET)
    Set convictionList = GetUniqueValues(tradesWs.Range("F:F"))
    
    row = 20
    For Each conviction In convictionList
        If conviction <> "Conviction" And conviction <> "" Then
            ws.Cells(row, 4).value = conviction
            ws.Cells(row, 5).Formula = "=COUNTIFS(Trades!F:F,""" & conviction & """,Trades!I:I,""Win"")/COUNTIF(Trades!F:F,""" & conviction & """)*100"
            ws.Cells(row, 6).Formula = "=SUMIF(Trades!F:F,""" & conviction & """,Trades!J:J)"
            ws.Cells(row, 7).Formula = "=COUNTIF(Trades!F:F,""" & conviction & """)"
            row = row + 1
        End If
    Next conviction
    
    ws.Range("E20:E" & row - 1).NumberFormat = "0.00%"
    ws.Range("F20:F" & row - 1).NumberFormat = "$#,##0.00"
End Sub

' Update VBA analysis table
Sub UpdateVBAAnalysis()

    Dim ws As Worksheet
    Dim tradesWs As Worksheet
    Dim vbaList As Collection
    Dim vba As Variant
    Dim row As Integer
    
    Set ws = ThisWorkbook.Worksheets(METRICS_SHEET)
    Set tradesWs = ThisWorkbook.Worksheets(TRADES_SHEET)
    Set vbaList = GetUniqueValues(tradesWs.Range("G:G"))
    
    row = 20
    For Each vba In vbaList
        If vba <> "VBA" And vba <> "" Then
            ws.Cells(row, 9).value = vba
            ws.Cells(row, 10).Formula = "=COUNTIFS(Trades!G:G,""" & vba & """,Trades!I:I,""Win"")/COUNTIF(Trades!G:G,""" & vba & """)*100"
            ws.Cells(row, 11).Formula = "=SUMIF(Trades!G:G,""" & vba & """,Trades!J:J)"
            ws.Cells(row, 12).Formula = "=COUNTIF(Trades!G:G,""" & vba & """)"
            row = row + 1
        End If
    Next vba
    
    ws.Range("J5:J" & row - 1).NumberFormat = "0.00%"
    ws.Range("K5:K" & row - 1).NumberFormat = "$#,##0.00"
End Sub

' Update equity curve data
Sub UpdateEquityCurve()
    Dim ws As Worksheet
    Dim tradesWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cumulativePnL As Double
    
    Set ws = ThisWorkbook.Worksheets(METRICS_SHEET)
    Set tradesWs = ThisWorkbook.Worksheets(TRADES_SHEET)
    
    lastRow = tradesWs.Cells(tradesWs.Rows.count, "A").End(xlUp).row
    
    ' Clear existing equity curve data
    ws.Range("N3:P1000").Clear
    
    ' Add headers
    ws.Range("N3").value = "Equity Curve"
    ws.Range("N3").Font.Bold = True
    ws.Range("N4:P4").value = Array("Trade #", "Date", "Cumulative P&L")
    ws.Range("N4:P4").Font.Bold = True
    
    cumulativePnL = 0
    For i = 2 To lastRow
        If tradesWs.Cells(i, 1).value <> "" Then
            If tradesWs.Cells(i, 10).value <> "" Then
                cumulativePnL = cumulativePnL + tradesWs.Cells(i, 10).value
                ws.Cells(i + 3, 14).value = i - 1 ' Trade number
                ws.Cells(i + 3, 15).value = tradesWs.Cells(i, 4).value ' Exit date
                ws.Cells(i + 3, 16).value = cumulativePnL
            End If
        End If
    Next i
    
    ws.Range("O5:O" & lastRow + 3).NumberFormat = "mm/dd/yyyy"
    ws.Range("P5:P" & lastRow + 3).NumberFormat = "$#,##0.00"
End Sub

' Get unique values from a range
Function GetUniqueValues(rng As Range) As Collection
    Dim uniqueValues As New Collection
    Dim cell As Range
    Dim value As Variant
    
    On Error Resume Next
    For Each cell In rng
        If cell.value <> "" Then
            value = cell.value
            uniqueValues.Add value, CStr(value)
        End If
    Next cell
    On Error GoTo 0
    
    Set GetUniqueValues = uniqueValues
End Function

' Custom function to calculate max drawdown
Function CalcMaxDrawdown() As Double

    Dim tradesWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim runningPnL As Double
    Dim peak As Double
    Dim maxDrawdown As Double
    Dim currentDrawdown As Double
    
    Set tradesWs = ThisWorkbook.Worksheets(TRADES_SHEET)
    lastRow = tradesWs.Cells(tradesWs.Rows.count, "A").End(xlUp).row
    
    runningPnL = 0
    peak = 0
    maxDrawdown = 0
    
    For i = 2 To lastRow
        If tradesWs.Cells(i, 1).value <> "" Then
            runningPnL = runningPnL + tradesWs.Cells(i, 10).value
            If runningPnL > peak Then peak = runningPnL
            currentDrawdown = peak - runningPnL
            If currentDrawdown > maxDrawdown Then maxDrawdown = currentDrawdown
        End If
    Next i
    
    CalcMaxDrawdown = maxDrawdown
End Function

' Custom function to calculate expectancy
Function CalcExpectancy() As Double
    Dim tradesWs As Worksheet
    Dim winRate As Double
    Dim avgWin As Double
    Dim avgLoss As Double
    
    Set tradesWs = ThisWorkbook.Worksheets("Trades")
    
    ' Calculate win rate as decimal
    winRate = Application.WorksheetFunction.CountIf(tradesWs.Range("I:I"), "Win") / Application.WorksheetFunction.CountA(tradesWs.Range("I:I"))
    
    ' Calculate average win and loss
    avgWin = Application.WorksheetFunction.AverageIf(tradesWs.Range("I:I"), "Win", tradesWs.Range("J:J"))
    avgLoss = Abs(Application.WorksheetFunction.AverageIf(tradesWs.Range("I:I"), "Loss", tradesWs.Range("J:J")))
    
    CalcExpectancy = winRate * avgWin - (1 - winRate) * avgLoss
End Function

' Create all charts
Sub CreateAllCharts()
    Call CreateEquityCurveChart
    'Call CreateSetupPerformanceChart
    Call CreateMarketRegimeChart
    'Call CreateConvictionChart
    Call CreateVBAChart
End Sub

' Create equity curve chart
Sub CreateEquityCurveChart()
    Dim ws As Worksheet
    Dim chartsWs As Worksheet
    Dim chartObj As ChartObject
    Dim dataRange As Range
    
    Set ws = ThisWorkbook.Worksheets(METRICS_SHEET)
    Set chartsWs = ThisWorkbook.Worksheets(CHARTS_SHEET)
    
    ' Clear existing charts
    chartsWs.Cells.Clear
    
    ' Create equity curve chart
    Set dataRange = ws.Range("N4:P" & ws.Cells(ws.Rows.count, "N").End(xlUp).row)
    Set chartObj = chartsWs.ChartObjects.Add(10, 10, 600, 300)
    
    With chartObj.Chart
        .SetSourceData dataRange
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "Equity Curve"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Trade Number"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Cumulative P&L ($)"
        .SeriesCollection(1).Name = "Cumulative P&L"
    End With
End Sub

' Create setup performance chart
Sub CreateSetupPerformanceChart()
    Dim ws As Worksheet
    Dim chartsWs As Worksheet
    Dim chartObj As ChartObject
    Dim dataRange As Range
    
    Set ws = ThisWorkbook.Worksheets(METRICS_SHEET)
    Set chartsWs = ThisWorkbook.Worksheets(CHARTS_SHEET)
    
    ' Find the range with setup data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).row
    Set dataRange = ws.Range("D4:G" & lastRow)
    
    Set chartObj = chartsWs.ChartObjects.Add(630, 10, 500, 300)
    
    With chartObj.Chart
        .SetSourceData dataRange
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Setup Performance Analysis"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Setup Type"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Win Rate % / Trade Count"
    End With
End Sub

' Create market regime chart
Sub CreateMarketRegimeChart()
    Dim ws As Worksheet
    Dim chartsWs As Worksheet
    Dim chartObj As ChartObject
    Dim dataRange As Range
    
    Set ws = ThisWorkbook.Worksheets(METRICS_SHEET)
    Set chartsWs = ThisWorkbook.Worksheets(CHARTS_SHEET)
    
    ' Find the range with market regime data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
    Set dataRange = ws.Range("I4:L" & lastRow)
    
    Set chartObj = chartsWs.ChartObjects.Add(10, 330, 600, 300)
    
    With chartObj.Chart
        .SetSourceData dataRange
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Market Regime Performance"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Market Regime"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Win Rate % / Trade Count"
    End With
End Sub

' Create conviction analysis chart
Sub CreateConvictionChart()
    Dim ws As Worksheet
    Dim chartsWs As Worksheet
    Dim chartObj As ChartObject
    Dim dataRange As Range
    
    Set ws = ThisWorkbook.Worksheets(METRICS_SHEET)
    Set chartsWs = ThisWorkbook.Worksheets(CHARTS_SHEET)
    
    ' Find the range with conviction data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).row
    If lastRow > 20 Then
        Set dataRange = ws.Range("D16:G" & lastRow)
        
        Set chartObj = chartsWs.ChartObjects.Add(630, 330, 500, 300)
        
        With chartObj.Chart
            .SetSourceData dataRange
            .ChartType = xlColumnClustered
            .HasTitle = True
            .ChartTitle.Text = "Conviction Level Analysis"
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "Conviction Level"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "Win Rate % / Trade Count"
        End With
    End If
End Sub

' Create VBA analysis chart
Sub CreateVBAChart()
    Dim ws As Worksheet
    Dim chartsWs As Worksheet
    Dim chartObj As ChartObject
    Dim dataRange As Range
    
    Set ws = ThisWorkbook.Worksheets(METRICS_SHEET)
    Set chartsWs = ThisWorkbook.Worksheets(CHARTS_SHEET)
    
    ' Find the range with VBA data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "I").End(xlUp).row
    If lastRow > 20 Then
        Set dataRange = ws.Range("I16:L" & lastRow)
        
        Set chartObj = chartsWs.ChartObjects.Add(10, 660, 600, 300)
        
        With chartObj.Chart
            .SetSourceData dataRange
            .ChartType = xlColumnClustered
            .HasTitle = True
            .ChartTitle.Text = "VBA Level Analysis"
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "VBA Level"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "Win Rate % / Trade Count"
        End With
    End If
    With ws
        .Range("B11").Formula = "=CalcMaxDrawdown()"
        .Range("B12").Formula = "=CalcExpectancy()"

        .Range("B11").value = .Range("B11").value
        .Range("B12").value = .Range("B12").value
    End With
End Sub


