Attribute VB_Name = "Performance"
Option Explicit

Sub PerformanceHISTORY() ' Records Performance
    On Error GoTo ErrorHandler
    
    ' Variable declarations
    Dim ws As Worksheet, wsTLog As Worksheet, wsRpt As Worksheet
    Dim lastRow As Long, nextLogRow As Long
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
    
   
    ' Get last row of data in Reports sheet
    lastRow = GetLastRow(wsRpt, "B")
    If lastRow < 2 Then
        MsgBox "No data found in Reports sheet!", vbExclamation
        GoTo Cleanup
    End If
    
    ' Find next empty row in Performance column A
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
       
    ' Navigate to the new row
    ws.Activate
    ws.Range("A" & nextLogRow).Select
    Application.Calculate
    Application.Wait Now + TimeValue("00:00:02")
    
    'If Not perfTest Then MsgBox "Performance data copied successfully to row " & nextLogRow, vbInformation
    
Cleanup:
    ' Restore application settings
    Application.Calculation = originalCalculation
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.CutCopyMode = False
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in CopyPerformance: " & Err.Description & " (Line: " & Erl & ")", vbExclamation
    GoTo Cleanup
End Sub

' Helper function to safely get worksheet reference
Private Function GetWorksheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
End Function

' Helper function to find last row with data
Private Function GetLastRow(ws As Worksheet, columnLetter As String) As Long
    GetLastRow = ws.Cells(ws.Rows.count, columnLetter).End(xlUp).row
End Function

' Helper function to find next empty row
Private Function FindNextEmptyRow(ws As Worksheet, columnLetter As String, startRow As Long) As Long
    Dim checkRow As Long
    checkRow = startRow
    
    Do While Not IsEmpty(ws.Cells(checkRow, columnLetter))
        checkRow = checkRow + 1
        ' Safety check to prevent infinite loop
        If checkRow > ws.Rows.count Then Exit Do
    Loop
    
    FindNextEmptyRow = checkRow
End Function

' Helper subroutine to setup Trade Log
Sub SetupTradeLog(wsTLog As Worksheet, lastRow As Long)
    With wsTLog
        .Range("B4:V53").ClearContents
        .Range("B1:V1").Copy
        .Range("B4:V" & lastRow).PasteSpecial Paste:=xlPasteFormulas
        Application.CutCopyMode = False
    End With
End Sub

Private Sub SetupPerformance(ws As Worksheet, lastRow As Long)
    With ws
        .Range("A1:X5").Copy
        .Range("A" & lastRow & ":X" & lastRow + 5).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        
        '.Range("B" & lastRow & ":X" & lastRow + 2).value = .Range("B1:X3").value
        Application.CutCopyMode = False
    End With
End Sub
' Helper function to validate template formulas
Private Function ValidateTemplateFormulas(ws As Worksheet) As Boolean
    ValidateTemplateFormulas = (Application.CountA(ws.Range("I1:AT1")) > 0)
End Function

' Helper subroutine to copy performance data
Private Sub origCopyPerformanceData(ws As Worksheet, nextLogRow As Long)
    With ws
        ' Write test group label in column A
        .Range("A" & nextLogRow).value = .Range("A1").value
        
        ' Copy formulas from named range or template row
        If NameExists("TemplateFormulas") Then
            ' Use named range if it exists
            ThisWorkbook.Names("TemplateFormulas").RefersToRange.Copy
        Else
            ' Fall back to row 1 template
            .Range("B1:X3").Copy
        End If
        
        ' Paste formulas to the target row
        .Range("B" & nextLogRow & ":X" & nextLogRow).PasteSpecial Paste:=xlPasteFormulas
        Application.CutCopyMode = False
    End With
End Sub

' Helper function to check if named range exists
Private Function NameExists(rangeName As String) As Boolean
    On Error Resume Next
    Dim testRange As Range
    Set testRange = ThisWorkbook.Names(rangeName).RefersToRange
    NameExists = Not testRange Is Nothing
    On Error GoTo 0
End Function

Sub performanceFormulasToValues(ws As Worksheet, newLrow As Long)
    Dim lastRow As Long, lastCol As Long
    Dim formulaRange As Range
    Set formulaRange = ws.UsedRange
    
    With ws
        lastRow = .Cells(.Rows.count, "I").End(xlUp).row
        lastCol = .Cells(9, .Columns.count).End(xlToLeft).Column
        .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).value = .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).value
    End With
End Sub
' Placeholder for the PerformanceFormulasToValues subroutine
Private Sub testperformanceFormulasToValues(ws As Worksheet, formulaRange As Range)
    ' This should contain the logic from your existing performanceFormulasToValues function
    ' Example implementation:
    'Dim formulaRange As Range
    'Set formulaRange = ws.UsedRange

    On Error Resume Next
    formulaRange.value = formulaRange.value
    On Error GoTo 0
End Sub

' Your existing ExtractAndIncrementGroupNumber function would go here
' (Include the actual implementation of this function)
Private Function ExtractAndIncrementGroupNumber(groupText As String) As Long
    ' Placeholder - implement your existing logic here
    ' Example basic implementation:
    Dim parts() As String
    Dim numberPart As String
    
    parts = Split(groupText, "_")
    If UBound(parts) >= 2 Then
        numberPart = parts(2)
        If IsNumeric(numberPart) Then
            ExtractAndIncrementGroupNumber = CLng(numberPart) + 1
        Else
            ExtractAndIncrementGroupNumber = 1
        End If
    Else
        ExtractAndIncrementGroupNumber = 1
    End If
End Function

Sub TISystemPerformance()
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim wsDash As Worksheet, ws As Worksheet, wsTJX As Worksheet
    Dim startTime As Double
    Dim totalIterations As Long
    Dim currentIteration As Long
    Dim initialGroupValue As String
    Dim isWeeklyAnalysis As Boolean
    Dim minPrice As Double, maxPrice As Double
    
    ' Error handling
    On Error GoTo ErrorHandler
    
    startTime = Timer
    perfTest = True
    
    If Not ConfirmProcessing() Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("PERFORMANCE")
    ws.Range("A12:X300").ClearContents
    
    Set wsDash = ThisWorkbook.Sheets("DashBoard")
    Set wsTJX = ThisWorkbook.Sheets("TJX")
    
    ' Determine analysis type
    isWeeklyAnalysis = (UCase(Trim(wsDash.Range("H1").value)) = "WEEKLY")
    
    minScore = 5
    startDate = wsDash.Range("H5").value
    minPrice = wsTJX.Range("C1").value
    maxPrice = wsTJX.Range("E1").value
   
    ' Fix: Set appropriate end date based on analysis type
    If isWeeklyAnalysis Then
        endDate = DateAdd("m", 12, startDate) ' Monthly range for weekly analysis
        If endDate > GetPreviousMonday(Date) Then endDate = GetPreviousMonday(Date)
    Else
        endDate = DateAdd("d", 90, startDate) ' Max 90 days for daily analysis
        If endDate > GetPreviousWorkday(Date) Then endDate = GetPreviousWorkday(Date)
    End If
          
    If Not GetUserInputs(minScore, minPrice, maxPrice, startDate) Then Exit Sub
    
    endDate = InputBox("Optional: Change or Confirm endDate):", "endDate set at ", Format(endDate, "yyyy-mm-dd"))

    If startDate >= endDate Then
        MsgBox "Invalid date range: Start date must be before end date.. Resetting endate", vbCritical
        endDate = IIf(isWeeklyAnalysis, endDate = GetPreviousMonday(Date), endDate = GetPreviousWorkday(Date))
    End If
    
    currentDate = startDate
    
    ' Fix: Calculate total iterations based on analysis type
    If isWeeklyAnalysis Then
        totalIterations = DateDiff("ww", startDate, endDate) + 1
    Else
        totalIterations = DateDiff("d", startDate, endDate) + 1
    End If
    
    currentIteration = 0
    
    ' Main processing loop
    Do While currentDate <= endDate
        currentIteration = currentIteration + 1
        
        ' Update status bar for progress tracking
        Application.StatusBar = "Processing backtest... " & currentIteration & " of " & totalIterations & " (" & Format((currentIteration / totalIterations) * 100, "0") & "%)"
        
        ' Set the current date in the dashboard
        wsDash.Range("H5").value = currentDate
        pubNotice = True
        
        ' Process the current iteration
        Call FilterAndReport
        Call PerformanceHISTORY
        
        ' Move to next date
        If isWeeklyAnalysis Then
            currentDate = GetNextMonday(currentDate)
        Else
            currentDate = GetNextWorkday(currentDate)
        End If
        
        Debug.Print "Analysis Type: " & IIf(isWeeklyAnalysis, "WEEKLY", "DAILY") & _
                   " | Date range: " & startDate & " to " & endDate & " | Current: " & currentDate
        
        ' Safety check to prevent infinite loop
        If currentIteration > totalIterations * 2 Then
            MsgBox "Loop exceeded expected iterations. Stopping for safety.", vbExclamation
            Exit Do
        End If
    Loop
    
    ' Clear status bar and show completion message
    Application.StatusBar = False
    MsgBox "TIS completed up to: " & Format(DateAdd("d", IIf(isWeeklyAnalysis, -7, -1), currentDate), "mm/dd/yyyy") & vbCrLf & _
           "Procedure completed in " & Round((Timer - startTime) / 60, 2) & " minutes." & vbCrLf & _
           "Total iterations processed: " & currentIteration, vbInformation
    
    perfTest = False
       
ErrorHandler:
    Application.StatusBar = False
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical
End Sub

' Helper function to check if worksheet exists
Function WorksheetExists(wsName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' Placeholder for missing functions - you'll need to implement these
Function newIsValidGroupFormat(groupValue As String) As Boolean
    ' Add your validation logic here
    ' For example: check if it matches "Test_Group_X" pattern
    IsValidGroupFormat = (Len(groupValue) > 0 And InStr(groupValue, "Group") > 0)
End Function

Function GetNextMonday(currentDate As Date) As Date
    ' Return the next Monday after currentDate
    Dim daysToMonday As Integer
    daysToMonday = (9 - Weekday(currentDate)) Mod 7
    If daysToMonday = 0 Then daysToMonday = 7 ' If it's already Monday, go to next Monday
    GetNextMonday = DateAdd("d", daysToMonday, currentDate)
End Function

Function GetNextWorkday(currentDate As Date) As Date
    ' Return the next workday (Monday-Friday)
    Dim nextDay As Date
    nextDay = DateAdd("d", 1, currentDate)
    
    ' Skip weekends
    Do While Weekday(nextDay) = vbSaturday Or Weekday(nextDay) = vbSunday
        nextDay = DateAdd("d", 1, nextDay)
    Loop
    
    GetNextWorkday = nextDay
End Function

Function IsValidGroupFormat(groupValue As String) As Boolean
    ' Check if the group value follows the expected format
    Dim regex As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^Test_Group_\d+$"
    regex.Global = False
    
    IsValidGroupFormat = regex.test(groupValue)
End Function

' Additional helper function to reset group numbering if needed
Sub ResetGroupNumbering()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PERFORMANCE DASHBOARD")
    
    ws.Range("B1").value = "Test_Group_1"
    MsgBox "Group numbering has been reset to Test_Group_1", vbInformation
End Sub




