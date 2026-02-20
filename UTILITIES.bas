Attribute VB_Name = "UTILITIES"
Option Explicit
Sub testP()
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("TPage")
Call CreateTVHyperLinks(ws)
End Sub

Sub UpdateDashFromReport()
    Dim sourcedata As Variant
    Dim wsDash As Worksheet, wsRpt As Worksheet, wsTLog As Worksheet
    Dim lastRow As Long
    
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsRpt = ThisWorkbook.Sheets("Reports")
    Set wsTLog = ThisWorkbook.Sheets("TRADE LOG")
    
    If wsDash.AutoFilterMode Then wsDash.AutoFilterMode = False
        
    lastRow = wsRpt.Cells(wsRpt.Rows.count, "B").End(xlUp).row
   
    ' Handle no data situation
    If lastRow < 4 Or wsRpt.Cells(wsRpt.Rows.count, "A").End(xlUp).row < 4 Then
        wsDash.Range("W5").Value = minScore
            If Not perfTest Then MsgBox "Few suggestions found in call to UpdateDashFromReport...ending..!", vbExclamation
        Exit Sub
    End If
    
    sourcedata = wsRpt.Range("B4:B" & (3 + lastRow)).Value2
    With wsDash
        .Range("A8:A57").ClearContents
        .Range("B3:AQ3").Copy
        .Range("B8:AQ57").PasteSpecial Paste:=xlPasteFormulas
        .Range("A8:A57").Font.Size = 10
        .Range("A8").Resize(lastRow - 3, 1).Value = sourcedata
    End With
    
    Call ModDataFromBackup
     
    Call LogReports
    Application.CutCopyMode = False
    
    ' OPTIMIZATION: Write formulas to Trade Log
    'Call SetupTradeLog(wsTLog, lastRow - 3)
    
    'Call CalcPnL(lastrow-3)
 
    Sheets("DashBoard").Select
End Sub

Sub ModDataFromBackup()
    Dim wsDash As Worksheet, wsBackupAll As Worksheet, wsData As Worksheet
    Dim maxPrice As Double, minPrice As Double, minScore As Double, endDate As Date
    Dim lastRowDash As Long, lastRowBackup As Long
    Dim i As Long, j As Long, outputIndex As Long, startIdx As Long, recordCount As Long
    Dim ticker As String, tickerPrice As Double, currentDate As Date
    Dim DashData As Variant, backupData As Variant
    Dim outputData() As Variant, tickerRecords() As Variant
    Dim frequency As String, includeRecord As Boolean
    Dim isWeekly As Boolean
    
    'Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set worksheet references
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsBackupAll = ThisWorkbook.Sheets("BackupAll")
    Set wsData = ThisWorkbook.Sheets("Data")
    
    ' Clear existing data
    wsData.Range("A2:G100000").ClearContents
    
    ' Get filter criteria from Dashboard
    frequency = wsDash.Range("H1").Value2
    endDate = wsDash.Range("H5").Value2
    maxPrice = wsDash.Range("Y5").Value2
    minPrice = wsDash.Range("Y6").Value2
    minScore = wsDash.Range("W5").Value2
    
    ' Validate frequency selection
    If UCase(frequency) = "WEEKLY" Then
        isWeekly = True
    ElseIf UCase(frequency) = "DAILY" Then
        isWeekly = False
    Else
        MsgBox "Please select either 'DAILY' or 'WEEKLY' in cell H1", vbExclamation
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
    
    ' Find last rows
    lastRowDash = wsDash.Cells(wsDash.Rows.count, "A").End(xlUp).row
    lastRowBackup = wsBackupAll.Cells(wsBackupAll.Rows.count, "A").End(xlUp).row
    
    ' Validate data exists
    If lastRowDash < 8 Or lastRowBackup < 2 Then
        MsgBox "Insufficient data in Dashboard or BackupAll sheets", vbExclamation
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
    
    ' Load all data into arrays
    DashData = wsDash.Range("A8:C" & lastRowDash).Value2
    backupData = wsBackupAll.Range("A2:G" & lastRowBackup).Value2
    
    ' Prepare output array (estimate: 50 records per ticker)
    ReDim outputData(1 To UBound(DashData, 1) * 50, 1 To 7)
    ReDim tickerRecords(1 To lastRowBackup, 1 To 7)
    outputIndex = 0
    
    ' Loop through each ticker in Dashboard
    For i = 1 To UBound(DashData, 1)
        ticker = DashData(i, 1)
        tickerPrice = DashData(i, 3)
        
        ' Skip if price is outside range
        If tickerPrice < minPrice Or tickerPrice > maxPrice Then GoTo NextTicker
        
        ' Collect matching records from BackupAll
        recordCount = 0
        
        For j = 1 To UBound(backupData, 1)
            ' Skip if not matching ticker
            If backupData(j, 7) <> ticker Then GoTo NextBackupRow
            
            currentDate = backupData(j, 1)
            
            ' Skip if date is beyond end date
            If currentDate > endDate Then GoTo NextBackupRow
            
            ' Determine if record should be included
            includeRecord = False
            If isWeekly Then
                If weekday(currentDate) = weekday(endDate) Then includeRecord = True
                
                ' Include only Mondays (Weekday = 2 with vbSunday)
                'If weekday(currentDate, vbSunday) = 2 Then includeRecord = True
            Else
                ' Include all dates for DAILY
                includeRecord = True
            End If
            
            If includeRecord Then
                recordCount = recordCount + 1
                tickerRecords(recordCount, 1) = backupData(j, 1)
                tickerRecords(recordCount, 2) = backupData(j, 2)
                tickerRecords(recordCount, 3) = backupData(j, 3)
                tickerRecords(recordCount, 4) = backupData(j, 4)
                tickerRecords(recordCount, 5) = backupData(j, 5)
                tickerRecords(recordCount, 6) = backupData(j, 6)
                tickerRecords(recordCount, 7) = backupData(j, 7)
            End If
            
NextBackupRow:
        Next j
        
        ' Take the last 50 records (most recent) - ONLY if we have at least 50
        If recordCount >= 50 Then
            startIdx = recordCount - 49
            
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
        
NextTicker:
    Next i
    
    ' Write all data at once if we have any
    If outputIndex > 0 Then
        wsData.Range("A2").Resize(outputIndex, 7).Value2 = outputData
    End If
    
    Call UpdateDataWithATR_Complete
    Call UpdateDataWithEnhancedIndicators
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Optional: Uncomment to show completion message
     'MsgBox "ModDataFromBackup with ATR and Indicators complete! " & outputIndex & " records written using " & frequency & " frequency.", vbInformation
    
End Sub


Sub myNewGroup()
Dim ws As Worksheet, WKS As Worksheet
Dim lastRow As Long
Set ws = ThisWorkbook.Sheets("cweSignals")
'Set WKS = ThisWorkbook.Sheets("ATR Signals")
    

        
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).row
    Range("B1").Select
    ActiveWorkbook.Worksheets("cweSignals").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("cweSignals").Sort.SortFields.Add2 key:=Range( _
        "B1:B" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("cweSignals").Sort
        .SetRange Range("A2:I" & lastRow)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    'Call cweFilter
    'Call ApplySignalFilters(WKS)
    MsgBox " weektest done"
    Sheets("cweSignals").Select
    'Dim lastRow As Long, i As Long
    'Dim analysisDate As Date
    
    'Dim wsTrades As Worksheet, wsDash As Worksheet, wsTLog As Worksheet
    
    'analysisDate = wsDash.Range("H5").value
    
    'Call DataFromBackup(analysisDate)
End Sub
    
Sub setupTradesPage()

    Dim lastRow As Long, i As Long
    Dim analysisDate As Date

    Dim wsTrades As Worksheet, wsDash As Worksheet, wsTLog As Worksheet
    
    Set wsTrades = ThisWorkbook.Sheets("Trades")

    lastRow = wsTrades.Cells(wsTrades.Rows.count, "c").End(xlUp).row
    
    With wsTrades
        .Range("A2:A" & lastRow).Formula = "=ROW()"
        .Range("B2:B" & lastRow).Formula = "= ""Group "" & " & "row()"
    End With
     
End Sub

' Add data validation for dropdown lists
Sub AddDataValidation()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(TRADES_SHEET)
    
    ' Setup dropdown options
    With ws.Range("E:E") ' Setup column
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
                        AlertStyle:=xlValidAlertStop, _
                        Formula1:="Momentum,Price Action,Trend,Breakout,Mixed"
    End With
    
    With ws.Range("F:F") ' Conviction column
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
                        AlertStyle:=xlValidAlertStop, _
                        Formula1:="High,Medium,Standard,Low"
    End With
    
    With ws.Range("H:H") ' Market Regime column
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
                        AlertStyle:=xlValidAlertStop, _
                        Formula1:="Normal,Trending,Sideways,Volatile"
    End With
    
    With ws.Range("I:I") ' Outcome column
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, _
                        AlertStyle:=xlValidAlertStop, _
                        Formula1:="Win,Loss"
    End With
End Sub

Sub oldtest()

Dim wsTLog As Worksheet
Dim lastRow As Long
'Call SetupTradeLog(wsTLog, lastRow)
Set wsTLog = ThisWorkbook.Sheets("TRADE LOG")
    lastRow = 20
    ThisWorkbook.Sheets("TRADE LOG (2)").Range("A1:W3").Copy
    With wsTLog
        .Range("A1").PasteSpecial Paste:=xlPasteFormulas
         Application.CutCopyMode = False
        .Range("B4:V53").ClearContents
        .Range("B1:V1").Copy
        .Range("B4:V" & lastRow).PasteSpecial Paste:=xlPasteFormulas
        .Range("A4:A" & lastRow).Formula = "= ""Group "" & " & "row()-3"
        .Range("N2").Value = 7
        Application.Calculate
        Application.CutCopyMode = False
    End With
End Sub

Sub quicktest()

Dim keyDay As Date

keyDay = DateAdd("d", 1, Date)
MsgBox "Next workday is " & GetNextWorkday(keyDay), vbInformation
End Sub

Sub toggleTJXFilter()
 Range("A2:W2").AutoFilter
 Range("A1").Select
End Sub

Sub toggleDashFilter()
 Dim ws As Worksheet
 Set ws = ThisWorkbook.Sheets("DashBoard")
 
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    Else
        ws.Range("$A$7:$AQ$57").AutoFilter Field:=1, Criteria1:="<>"
        ThisWorkbook.Sheets("TRADE LOG").Range("N2").Formula = "=DashBoard!$AQ$5"
    End If
    
End Sub

Sub toggleReportFilter()
 Range("A3:j3").AutoFilter
 Range("C3").Select
End Sub

Sub toggleDataFilter()
 Range("A1:G1").AutoFilter
 Range("A1").Select
End Sub

Sub toggleTradeFilter()
 Range("B3:V3").AutoFilter
 Range("V3").Select
End Sub

Sub toggleIndicatorFilter()
 Range("A6:AF6").AutoFilter
 Range("A6").Select
End Sub
    
Sub timerTEST()
    Dim PauseTime As Double, startTime As Double, Finish As Double
    
    PauseTime = 0.005   ' Set duration.
    startTime = Timer    ' Set start time.
    Do While Timer < startTime + PauseTime
  '      DoEvents    ' Yield to other processes.
    Loop
    Finish = Timer    ' Set end time.
    'MsgBox "Paused for " & Finish - startTime & " seconds"

    Call DisplayCompletionMessage(startTime)
End Sub

Sub moretest()
Dim ws As Worksheet
Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets("Reports")
    'Set wsDash = ThisWorkbook.Sheets("DashBoard")
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    ApplyFormulasOptimized ws, lastRow
    'new
       'wsRpt.Range("C4:C" & lastRow).value = wsDash.Range("AP8:AP" & lastRow + 8).value
       'Dim i As Long
       'For i = 4 To lastRow
        '  wsRpt.Cells(i, 3).Value = wsDash.Cells(i + 4, 42).Value
       'Next i
    

End Sub

Sub nextsettest()
Dim PauseTime, Start, Finish, TotalTime
'If (MsgBox("Press Yes to pause for 5 seconds", 4)) = vbYes Then
   PauseTime = 5    ' Set duration.
    Start = Timer    ' Set start time.
  Do While Timer < Start + PauseTime
  '      DoEvents    ' Yield to other processes.
    Loop
    Finish = Timer    ' Set end time.
    TotalTime = Finish - Start    ' Calculate total time.
    MsgBox "Paused for " & TotalTime & " seconds"
     ' End
    'End If
'   Dim ws As Worksheet
'   Set ws = ThisWorkbook.Sheets("DataHistory")
    'ThisWorkbook.Save
    'ThisWorkbook.Sheets("DashBoard").Range("$A$7:$AE$57").AutoFilter '  .AutoFilter Field:=9, Criteria1:="1"
    
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ActiveWorkbook ' Or another specific workbook
    Set ws = wb.Sheets("ReportHistory")
    
    Call CreateHyperLinks(ws)
    
    Set ws = ThisWorkbook.Sheets("DataHistory")
    
    Call DeleteNARows(ws)
    
    wb.Close SaveChanges:=True
    Application.Quit

End Sub

Sub CheckQualityTJXData()
    Dim i As Long
    Dim lastRow As Long
    Dim cellValue As Variant
    Dim wsTJX As Worksheet
    
    ' Set the worksheet
    Set wsTJX = ThisWorkbook.Worksheets("TJX")
    
    ' Determine the last row (adjust column if needed)
    lastRow = wsTJX.Cells(wsTJX.Rows.count, "C").End(xlUp).row
    
    ' Loop through rows starting from C3
    For i = 3 To lastRow
        With wsTJX.Range("C" & i)
            If IsError(.Value) Then
                Debug.Print "Row " & i & Err; ": Skipped due to error in cell"
            Else
                cellValue = .Value
                Debug.Print "Row " & i & ": " & cellValue ' Replace with your logic
            End If
        End With
    Next i
End Sub

Sub speedTest()
    Dim ws As Worksheet
    Dim t As Double
    
    Application.Calculation = xlCalculationManual
    Debug.Print "Starting Calculation Time Check..."
    
    For Each ws In ThisWorkbook.Worksheets
        t = Timer
        ws.Calculate
        Debug.Print "Sheet: " & ws.Name & " - Time: " & Format(Timer - t, "0.00") & " sec"
    Next ws
    
    Application.Calculation = xlCalculationAutomatic
    Debug.Print "Done."
End Sub

Sub DeleteGhostTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    
    On Error Resume Next
    For Each ws In ThisWorkbook.Sheets
        For Each tbl In ws.ListObjects
            If tbl.Name = "Table1" Then
                tbl.Delete
            End If
        Next tbl
    Next ws
    On Error GoTo 0
    
    MsgBox "Table1 removed (if it existed).", vbInformation
End Sub

Sub ResizeTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRange As Range

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change as needed
    ' Reference the table
    Set tbl = ws.ListObjects("Table1")

    ' Define new table range
    Set newRange = ws.Range("A2:G1000") ' Adjust as needed

    ' Resize the table
    tbl.Resize newRange

    MsgBox "Table resized successfully!", vbInformation
End Sub

