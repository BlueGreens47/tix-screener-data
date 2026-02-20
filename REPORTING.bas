Attribute VB_Name = "Reporting"
Sub theReporter()
    Dim wsRpt As Worksheet, wsRptHist As Worksheet, wsRptLog As Worksheet
    Dim visibleRange As Range, filteredData As Range
    Dim lastRow As Long, lastRptRow As Long, visibleCount As Long
    Dim endDate As Date, startDate As Date
    Dim minScoreCopy As Integer, iterationCount As Integer
    Dim dashBoard As Worksheet, wsTLog As Worksheet
    
    Const MAX_ROWS As Integer = 48 ' Max data rows (excludes headers)
    
    On Error GoTo ErrorHandler
    
    ' OPTIMIZATION: Disable all unnecessary Excel features
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With

    ' Set worksheet references once
    With ThisWorkbook
        Set wsRpt = .Worksheets("Reports")
        Set wsRptHist = .Worksheets("ReportHistory")
        Set wsRptLog = .Worksheets("ReportLog")
        Set dashBoard = .Worksheets("DashBoard")
        Set wsTLog = .Worksheets("TRADE LOG")
    End With
    
    ' OPTIMIZATION: Streamlined tidyReports call
    Call tidyReportsOptimized(wsRptHist)
    Call LogReports
    
    ' Get parameters once
    minScore = dashBoard.Range("W5").Value
    minScoreCopy = minScore
    endDate = dashBoard.Range("H5").Value
    startDate = IIf(pubNotice And weekday(Date, vbMonday) = 2, endDate - 7, endDate)
    
    ' OPTIMIZATION: Clear filters once at start
    dashBoard.AutoFilterMode = False
    wsRptHist.AutoFilterMode = False
    wsRptLog.AutoFilterMode = False
    wsTLog.AutoFilterMode = False
    
    ' OPTIMIZATION: Clear in single operation
    wsRpt.Range("A4:M400").ClearContents
    dashBoard.Range("A8:AP57").ClearContents
    
    ' Check data availability
    lastRow = wsRptHist.Cells(wsRptHist.Rows.count, "A").End(xlUp).row
    If lastRow < 3 Then
        If Not perfTest Then MsgBox "No data found in Report History.", vbExclamation
        GoTo Cleanup
    End If
    
    ' OPTIMIZATION: Use optimized filtering approach
    Set visibleRange = wsRptHist.Range("A3:J" & lastRow)
    visibleCount = OptimizedFilterAndCount(visibleRange, startDate, endDate, minScore, MAX_ROWS, dashBoard)
    
    
      '  If visibleCount > 50 Then visibleRange.AutoFilter Field:=3, Criteria1:=">=" & minScore
       ' visibleCount = Application.WorksheetFunction.Subtotal(103, visibleRange.Columns(1)) - 1
    
    
    ' Copy filtered data if any rows found
    If visibleCount > 0 Then
        On Error Resume Next
        Set filteredData = visibleRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo ErrorHandler
        
        If Not filteredData Is Nothing Then
            ' OPTIMIZATION: Direct copy without paste special
            filteredData.Copy
            wsRpt.Range("A3").PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            
            lastRptRow = wsRpt.Cells(wsRpt.Rows.count, "A").End(xlUp).row
            
            ' Ensure we don't exceed MAX_ROWS
            If lastRptRow - 2 > MAX_ROWS Then
                wsRpt.Rows(MAX_ROWS + 3 & ":" & lastRptRow).Delete
            End If
        End If
    End If
  
    ' Handle no data situation
    If visibleCount = 0 Or wsRpt.Cells(wsRpt.Rows.count, "A").End(xlUp).row < 4 Then
        dashBoard.Range("W5").Value = minScore
            If Not perfTest Then MsgBox "Few suggestions found...ending..!", vbExclamation
        Exit Sub
    End If
    
    
Cleanup:
    GoTo FinalCleanup

ErrorHandler:
    dashBoard.Range("W5").Value = minScore
    
FinalCleanup:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayAlerts = True
    End With
    
   'Call UpdateDashFromReport 'ReportToDashOptimized
    
    'If pubNotice = False And perfTest = False Then MsgBox "Reporting done", vbExclamation
    
    ' SendReport on Tuesdays when running in live/scheduled mode
    If Not perfTest Then If pubNotice And Weekday(Date, vbMonday) = 2 Then Call sendReport
        
End Sub

' OPTIMIZATION: New optimized filtering function
Private Function OptimizedFilterAndCount(ByRef visibleRange As Range, startDate As Date, endDate As Date, ByRef minScore As Variant, maxRows As Integer, dashBoard As Worksheet) As Long
    Dim iterationCount As Integer
    Dim visibleCount As Long
    
    iterationCount = 0
    
    Do
        ' Clear existing filters
        visibleRange.Parent.AutoFilterMode = False
        
        ' Apply filters
        With visibleRange
            .AutoFilter Field:=1, Criteria1:=">=" & Format(startDate, "m/d/yyyy"), _
                       Operator:=xlAnd, Criteria2:="<=" & Format(endDate, "m/d/yyyy")
            .AutoFilter Field:=3, Criteria1:=">=" & minScore, _
                       Operator:=xlOr, Criteria2:="<=" & minScore * -1
        End With
        
        ' Count visible rows
        visibleCount = Application.WorksheetFunction.Subtotal(103, visibleRange.Columns(1)) - 1
        
        If visibleCount > 50 Then visibleRange.AutoFilter Field:=3, Criteria1:=">=" & minScore
        visibleCount = Application.WorksheetFunction.Subtotal(103, visibleRange.Columns(1)) - 1
        
        If visibleCount < 10 Then visibleRange.AutoFilter Field:=3, Criteria1:=">=" & minScore * -1
        visibleCount = Application.WorksheetFunction.Subtotal(103, visibleRange.Columns(1)) - 1
        
        ' Exit if within limit or max iterations reached
        If visibleCount <= maxRows Or iterationCount >= 8 Then Exit Do
        
        ' Adjust minScore and iterate
        minScore = minScore + 1
        iterationCount = iterationCount + 1
        
        ' OPTIMIZATION: Update dashboard less frequently
        If iterationCount Mod 2 = 0 Then
            dashBoard.Range("W5").Value = minScore
            dashBoard.Range("W6").Value = iterationCount
        End If
        
    Loop
    
    ' Final dashboard update
    dashBoard.Range("W6").Value = iterationCount
    dashBoard.Range("W5").Value = minScore
    
    OptimizedFilterAndCount = visibleCount
End Function

' OPTIMIZATION: Streamlined tidyReports
Sub tidyReportsOptimized(ws As Worksheet)
    Dim lastRow As Long
    Dim dataRange As Range
    
    On Error GoTo ErrorHandler
    
    ' Determine the last row in column A
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lastRow < 4 Then GoTo Cleanup
    
    ' Set the main data range
    Set dataRange = ws.Range("A3:J" & lastRow)
    
    ' OPTIMIZATION: Apply all formulas in batch operations
    ApplyFormulasOptimized ws, lastRow
    
    ' OPTIMIZATION: Process data cleanup more efficiently
    ProcessDataCleanupOptimized ws, dataRange
    
    ' OPTIMIZATION: Apply final formatting in one go
    ApplyFinalFormattingOptimized ws, dataRange
    
Cleanup:
    Exit Sub
    
ErrorHandler:
    Resume Next
End Sub

 Sub ApplyFormulasOptimized(ByRef ws As Worksheet, ByVal lastRow As Long)
    ' OPTIMIZATION: Disable calculation during formula application
    'Application.Calculation = xlCalculationManual
    
    ' Copy formulas for columns D:E in one operation
    'ws.Range("D2:E2").Copy
    'ws.Range("D4:E" & lastRow).PasteSpecial Paste:=xlPasteFormulas
    
    ' Copy formulas for columns F:I in batch
    'ws.Range("F1:I1").Copy
    'ws.Range("F4:I" & lastRow).PasteSpecial Paste:=xlPasteFormulas
    
    Application.CutCopyMode = False
    
    ' OPTIMIZATION: Single calculation cycle
    'Application.Calculation = xlCalculationAutomatic
   ' DoEvents ' Allow calculation to complete
    
    ' Set column I to dashboard value
    
    ws.Range("F4:F" & lastRow).Value = ThisWorkbook.Sheets("DashBoard").Range("Q8:Q" & lastRow + 5).Value 'Regime
    ws.Range("G4:G" & lastRow).Value = ThisWorkbook.Sheets("DashBoard").Range("AO8:AO" & lastRow + 5).Value 'Setup
    ws.Range("H4:H" & lastRow).Value = ThisWorkbook.Sheets("DashBoard").Range("AP8:AP" & lastRow + 5).Value 'Rank
    ws.Range("I4:I" & lastRow).Value = ThisWorkbook.Sheets("DashBoard").Range("H1").Value 'Daily/Weekly
    ws.Range("J4:J" & lastRow).Value = ThisWorkbook.Sheets("DashBoard").Range("S8:S" & lastRow + 5).Value 'Origin
    'ws.Range("K4:K" & lastRow).Value = ThisWorkbook.Sheets("DashBoard").Range("R8:R" & lastRow + 5).Value2 'FundScore
    
    ' OPTIMIZATION: Convert to values in single operation
    Dim tempRange As Range
    Set tempRange = ws.Range("A3:J" & lastRow + 3)
    tempRange.Value = tempRange.Value
    
    ' Sort the combined range by ticker (column G) and then by Date (column A)
    tempRange.Sort key1:=ws.Range("C3"), Order1:=xlDescending, Header:=xlYes 'Key2:=wsTo.Range("A1"), Order2:=xlAscending, Header:=xlYes
    'Application.Calculation = xlCalculationManual
End Sub

Private Sub ProcessDataCleanupOptimized(ByRef ws As Worksheet, ByRef dataRange As Range)
    ' OPTIMIZATION: Remove duplicates before other operations
    dataRange.RemoveDuplicates Columns:=Array(1, 2, 9), Header:=xlYes ' 1 3 5 9
    
    ' Apply filters efficiently
    ws.AutoFilterMode = False
    dataRange.AutoFilter Field:=2, Criteria1:="<>"
End Sub

Private Sub ApplyFinalFormattingOptimized(ByRef ws As Worksheet, ByRef dataRange As Range)
    Dim lastDataRow As Long
    lastDataRow = dataRange.Rows.count + 2
    
    ' OPTIMIZATION: Apply all formatting in batch operations
    With ws
        ' Number formats
        .Range("A4:A" & lastDataRow).NumberFormat = "m/d/yyyy"
        .Range("D4:I" & lastDataRow).NumberFormat = "#,##0.00"
        
        ' Alignment - combine similar operations
        .Range("A3:B" & lastDataRow).HorizontalAlignment = xlCenter
        .Range("A4:D" & lastDataRow).HorizontalAlignment = xlLeft
        .Range("C4:E" & lastDataRow).HorizontalAlignment = xlRight
        
        ' Indentation
        .Range("B4:D" & lastDataRow).IndentLevel = 1
    End With
    
    ' OPTIMIZATION: Streamlined sorting
    ApplySortingOptimized ws, dataRange
End Sub

Private Sub ApplySortingOptimized(ByRef ws As Worksheet, ByRef dataRange As Range)
    ' OPTIMIZATION: Simplified sorting with fewer method calls
    dataRange.Sort key1:=ws.Range("B4"), Order1:=xlAscending, Header:=xlYes
End Sub

' ***RTD*******
Sub ReportToDashOptimized()
    Dim wsDash As Worksheet, wsRpt As Worksheet, wsTLog As Worksheet, wsTrade As Worksheet
    Dim lastRow As Long, dataCount As Long
    Dim sourcedata As Variant, targetRange As Range, sourceRange As Range
    Dim analysisDate As Date
    
    On Error GoTo ErrorHandler
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    With ThisWorkbook
        Set wsTLog = .Worksheets("TRADE LOG")
        Set wsDash = .Worksheets("DashBoard")
        Set wsRpt = .Worksheets("Reports")
        Set wsTrades = .Worksheets("Trades")
    End With
    
    ' OPTIMIZATION: Clear filters once at start
    wsDash.AutoFilterMode = False
    wsRpt.AutoFilterMode = False
    wsTrades.AutoFilterMode = False
           
    lastRow = wsRpt.Cells(wsRpt.Rows.count, "B").End(xlUp).row
    If lastRow < 4 Then GoTo Cleanup
    
    'Sort by Rank before copy
    Set sourceRange = wsRpt.Range("A3:I" & lastRow)
    With wsRpt.Sort
        .SortFields.Clear
        .SortFields.Add key:=wsRpt.Range("C3"), Order:=xlDescending
        .SetRange wsRpt.Range("A3:I" & lastRow)
        .Header = xlYes
        .Apply
    End With
    
    ' Date validation
    If wsDash.Range("H5").Value <> wsRpt.Range("A4").Value Then
        MsgBox "Check dates.. Exiting.", vbExclamation
        Exit Sub
    End If
    
    wsDash.Range("A8:A57").ClearContents
    analysisDate = wsDash.Range("H5").Value
    dataCount = Application.WorksheetFunction.min(lastRow - 3, 50)
    
    If dataCount > 0 Then
    ' OPTIMIZATION: Only copy formulas once for the batch size we need
        With wsDash
            .Range("A3:AP3").Copy
            .Range("A8:AP" & 7 + dataCount).PasteSpecial Paste:=xlPasteFormulas
            .Range("A8:AP" & 7 + dataCount).Font.Name = "Calibri"
            .Range("A8:AP" & 7 + dataCount).Font.Size = 8
        End With
        Application.CutCopyMode = False
        ' OPTIMIZATION: Direct array transfer
        sourcedata = wsRpt.Range("B4:B" & (3 + dataCount)).Value
        wsDash.Range("A8").Resize(dataCount, 1).Value = sourcedata
       
        Call ModDataFromBackup 'DataFromBackup(analysisDate)
        
        ' OPTIMIZATION: Batch update column C values
          
        Dim dashValues As Variant
        dashValues = wsDash.Range("AP8:AP" & (7 + dataCount)).Value
         
        Dim i As Long
        For i = 1 To dataCount
            wsRpt.Cells(i + 3, 8).Value = dashValues(i, 1)
        Next i
        
        Call LogReports
        'Application.CutCopyMode = False
        
        ' OPTIMIZATION: Write formulas to Trade Log
        Call SetupTradeLog(wsTLog, dataCount)
        
        'Call CalcPnL(dataCount)
    End If
    
    Call CreateHyperLinks(wsRpt)
    
    GoTo Cleanup
    
ErrorHandler:
    Resume Next
    
Cleanup:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    'Call seePerformance
    Sheets("DashBoard").Select
    'If Not pubNotice Or Not perfTest Then MsgBox "Report to Dash completed", vbExclamation
End Sub

Sub CreateTVHyperLinks(ws As Worksheet)
    Dim lastRow As Long
    Dim cell As Range
    Dim ticker As String
    Dim exchange As String
    Dim baseURL As String
       
    ws.AutoFilterMode = False
    ws.Range("B:B").Hyperlinks.Delete
    
    baseURL = "https://www.tradingview.com/symbols/"
   
    ' Find the last row in the column with stock tickers
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
 
    ' Loop through each cell in column B and create the hyperlink
    For Each cell In ws.Range("B4:B" & lastRow) ' Adjust range as needed
        ticker = cell.Value
        
        ' Get exchange from column A (adjust column as needed)
        exchange = ws.Cells(cell.row, 1).Value
        
        If ticker <> "" Then
            ' Default to NASDAQ if no exchange specified
            If exchange = "" Then exchange = "NASDAQ"
            
            ws.Hyperlinks.Add Anchor:=cell.Offset(0, 0), _
                Address:=baseURL & exchange & "-" & ticker & "/", _
                TextToDisplay:=ticker
        End If
        Application.CutCopyMode = False
    Next cell
End Sub

Sub CreateHyperLinks(ws As Worksheet)
    Dim lastRow As Long
    Dim cell As Range
    Dim ticker As String
    Dim baseURL As String
       
    ws.AutoFilterMode = False
    ws.Range("B:B").Hyperlinks.Delete
    
    baseURL = "https://finance.yahoo.com/quote/"
   
    ' Find the last row in the column with stock tickers
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
 
    ' Loop through each cell in column A and create the hyperlink in column B
    For Each cell In ws.Range("B4:B" & lastRow) ' Adjust range as needed
        ticker = cell.Value
        If ticker <> "" Then
            ws.Hyperlinks.Add Anchor:=cell.Offset(0, 0), Address:=baseURL & ticker, TextToDisplay:=ticker
        End If
         Application.CutCopyMode = False
    Next cell
End Sub

Sub sendReport()
    Dim OutlookApp As Object
    Dim MailItem As Object
    Dim MailBody As String
    Dim MailList As Worksheet
    Dim ws As Worksheet
    Dim dashBoard As Worksheet
    Dim ticker As String
    Dim i As Long, lastRow As Long
    Dim baseURL As String
    Dim cell As Range
    Dim myEmailList As Integer
    Dim fileName As String
    
    ' Set the worksheet containing the email addresses
    Set MailList = ThisWorkbook.Sheets("EmailList")
    ' Set the worksheet containing the report
    Set ws = ThisWorkbook.Sheets("Reports")
    ' Set the worksheet containing the additional data
    Set dashBoard = ThisWorkbook.Sheets("DashBoard")
    
    ' OPTIMIZATION: Clear filters once at start
    dashBoard.AutoFilterMode = False
    ws.AutoFilterMode = False
    
    ' Yahoo Finance base URL
    baseURL = "https://finance.yahoo.com/quote/"
    
    ' Find the last row in the report with tickers
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
       
    ' Check pubNotice and set email list
    myEmailList = 1
    'If pubNotice And Weekday(Date, vbMonday) = 2 Then myEmailList = 2
        
    ' Initialize Outlook application
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Loop through each email address in the list
    For i = 1 To MailList.Cells(Rows.count, myEmailList).End(xlUp).row
        Set MailItem = OutlookApp.CreateItem(0)
        
        ' Build the email body with embedded report content and hyperlinks
        MailBody = "<html><body style='font-size: 11pt; font-family: Arial, sans-serif;'><p>Hi Folks,</p>"
        MailBody = MailBody & "<p>This week's Stock Insights :</p>"
        
        ' Start table
        MailBody = MailBody & "<table border='1' style='font-size: 10pt; border-collapse: collapse; margin-bottom: 20px;'>"
        
        ' Add table headers
        MailBody = MailBody & "<tr style='background-color: #f2f2f2;'><th style='padding: 8px;'>Date</th><th>Ticker</th><th>Score</th><th>Name</th><th>Price</th>" ',<th>Stop</th><th>Entry</th><th>Exit</th></tr>"
        
        ' Add table rows with data and hyperlinks
        For Each cell In ws.Range("A4:A" & lastRow)
            ticker = ws.Cells(cell.row, 2).Value
            If ticker <> "" Then
                MailBody = MailBody & "<tr>"
                MailBody = MailBody & "<td style='padding: 8px;'>" & Format(ws.Cells(cell.row, 1).Value, "yyyy/mm/dd") & "</td>" ' Date
                MailBody = MailBody & "<td style='padding: 8px;'><a href='" & baseURL & ticker & "/chart'>" & ticker & "</a></td>" ' Ticker
                MailBody = MailBody & "<td style='padding: 8px; text-align: center;'>" & Format(ws.Cells(cell.row, 3).Value, "0") & "</td>" '  Score Value from wsRPT col 3 (orig 23-VBA/42-Rank)
                MailBody = MailBody & "<td style='padding: 8px;'>" & ws.Cells(cell.row, 4).Value & "</td>" 'Name
                MailBody = MailBody & "<td style='padding: 8px; text-align: right;'>" & Format(ws.Cells(cell.row, 5).Value, "0.00") & "</td>" ' Price
                'MailBody = MailBody & "<td style='padding: 8px; text-align: right;'>" & Format(ws.Cells(cell.row, 6).value, "0.00") & "</td>" ' Stop
                'MailBody = MailBody & "<td style='padding: 8px; text-align: right;'>" & Format(ws.Cells(cell.row, 7).value, "0.00") & "</td>" ' Entry
                'MailBody = MailBody & "<td style='padding: 8px; text-align: right;'>" & Format(ws.Cells(cell.row, 8).value, "0.00") & "</td>" ' Exit
                MailBody = MailBody & "</tr>"
            End If
        Next cell
        
        ' Close table
        MailBody = MailBody & "</table>"
        
        ' Additional insights section after the table
        MailBody = MailBody & "<div style='margin-top: 20px; padding: 15px; background-color: #f9f9f9; border-radius: 5px;'>"
        MailBody = MailBody & "<h3 style='color: #333;'>Key Insights:</h3>"
        MailBody = MailBody & "<p style='margin-bottom: 10px;'><strong>Growth Potential:</strong> Each of these picks is backed by strong indicators and market trends.</p>"
        MailBody = MailBody & "<p style='margin-bottom: 10px;'><strong>Technical Signals:</strong> Our analysis shows favorable conditions for entry or profit-taking.</p>"
        MailBody = MailBody & "<p style='margin-bottom: 10px;'><strong>Fundamental Strength:</strong> These stocks are backed by solid financials, ensuring lower risk.</p>"
        MailBody = MailBody & "</div>"
        
        MailBody = MailBody & "<p style='margin-top: 20px;'>Trade wisely, the Market is still unsettled. we'll be back next week with more Insights.</p>"
        MailBody = MailBody & "<tr>"
        MailBody = MailBody & "<p>Best regards,<br>TradeInsights</p></body></html>"
        
        ' Set and display or send email
        On Error Resume Next
        With MailItem
            .To = MailList.Cells(i, 1).Value
            .Subject = "Weekly Insights"
            .HTMLBody = MailBody
            .display
        End With
    Next i
    
    ' Clean up
    Set MailItem = Nothing
    Set OutlookApp = Nothing
            
    If pubNotice Then Call SaveAndClose
     
End Sub

Sub HistoryToDash()
    Dim wsDash As Worksheet, wsRpt As Worksheet
    Dim lastRow As Long, dataCount As Long
    Dim sourcedata As Variant, targetRange As Range
    Dim analysisDate As Date
    
    On Error GoTo ErrorHandler
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    With ThisWorkbook
        Set wsDash = .Worksheets("DashBoard")
        Set wsRpt = .Worksheets("ReportHistory")
    End With
    
    wsDash.AutoFilterMode = False
    wsRpt.AutoFilterMode = False
    lastRow = wsRpt.Cells(wsRpt.Rows.count, "B").End(xlUp).row
    
    If lastRow < 4 Then GoTo Cleanup
    If lastRow > 57 Then lastRow = 57
    
    analysisDate = wsDash.Range("H5")
    sourcedata = wsRpt.Range("B4:B" & lastRow).Value
    dataCount = Application.WorksheetFunction.min(lastRow - 3, 50)
    
    With wsDash.Range("A8:A57")
        .ClearContents
        .Font.Name = "Calibri"
        .Font.Size = 8
    End With
    
    Set targetRange = wsDash.Range("A8").Resize(dataCount, 1)
    
    ' Optimized data transfer using Index and Sequence
    targetRange.Value = Application.Index(sourcedata, Application.Sequence(dataCount), 1)
    
    With targetRange.Borders
        .LineStyle = xlContinuous
        
        With .item(xlEdgeLeft)
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .item(xlEdgeTop)
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .item(xlEdgeBottom)
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With .item(xlEdgeRight)
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        
        .item(xlInsideVertical).LineStyle = xlContinuous
        .item(xlInsideHorizontal).LineStyle = xlContinuous
    End With
    
    Call DataFromBackup(analysisDate)
        
    GoTo Cleanup
    
ErrorHandler:
    ' Call the error handling procedure
   ' HandleProcessingError "HistoryToDash", Err
    Resume Next
    
Cleanup:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    wsRpt.Range("$A$3:$H$3").AutoFilter
    'Call RecordPerformance
    
    If Not pubNotice Then MsgBox "History to Dash and Performance Record completed", vbExclamation
End Sub

Sub LogReports()
    Dim wsLog As Worksheet, wsRpt As Worksheet
    Dim lastRow As Long, nextLogRptRow As Long, dataCount As Long
    Dim logdataRange As Range

    'On Error GoTo ErrorHandler
    
    With ThisWorkbook
        Set wsLog = .Worksheets("ReportLog")
        Set wsRpt = .Worksheets("Reports") '.Worksheets("ReportHistory")
    End With
    
    wsLog.AutoFilterMode = False
    wsRpt.AutoFilterMode = False
    
    ' Find the last row in wsRpt
    lastRow = wsRpt.Range("A" & wsRpt.Rows.count).End(xlUp).row
    'lastRow = wsRpt.Cells(wsRpt.Rows.count, 1).End(xlUp).row
    
    ' Find the next available row in wsLog
    nextLogRptRow = wsLog.Range("A" & wsLog.Rows.count).End(xlUp).row + 1
    
     wsLog.Range("A" & nextLogRptRow & ":J" & nextLogRptRow + lastRow).Value = wsRpt.Range("A4:J" & lastRow + 4).Value
       ' Application.CutCopyMode = False
        
    ' Copy data and paste only values and formatting
    wsRpt.Range("A4:J" & lastRow).Copy
    wsLog.Range("A" & nextLogRptRow).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
     
    lastRow = nextLogRptRow + lastRow - 4
    
    Set logdataRange = wsLog.Range("A3:J" & lastRow)
    logdataRange.RemoveDuplicates Columns:=Array(1, 2, 9), Header:=xlYes
    
    With Application
        .CutCopyMode = False
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
      
    'If Not pubNotice Then MsgBox "Report Log Completed", vbExclamation
     
End Sub

