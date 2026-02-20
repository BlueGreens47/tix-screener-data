Attribute VB_Name = "Scheduling"
Option Explicit

' Constants for configuration
Private Const SCHEDULED_TIME As String = "03:00"
Private Const LOG_SHEET As String = "SchedulerLog"  ' Optional: for logging

' This sub is called when the workbook is opened
Sub ScheduledRun()
    On Error GoTo ErrorHandler
    
    ' Schedule the next run
    Application.OnTime TimeValue(SCHEDULED_TIME), "RunScheduledTask"

    Exit Sub
    
ErrorHandler:
    LogMessage "Error in ScheduleNextRun: " & Err.Description
    MsgBox "Error scheduling task: " & Err.Description, vbCritical
End Sub

Sub RunScheduledTask()
    On Error GoTo ErrorHandler
    
    LogMessage "RunScheduledTask started at: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
       
    ' Initialize global variables
    gStopMacro = False
    pubNotice = True
    
    Call ClearAllFilters
    
    ' Set dashboard values
    With ThisWorkbook.Sheets("DashBoard")
        .Range("H1").Value = "DAILY"
        .Range("H5").Value = GetPreviousWorkday(Date)
    End With
    
    ' Run Tuesday calculation if needed (only if Tuesday is a workday)
    ' Refreshes fundamental scores (PE, ROE, FCF etc.) on the TJX sheet weekly
    If weekday(Date, vbMonday) = 2 And IsWorkday(Date) Then
        LogMessage "Running Tuesday fundamentals calculation..."
        Call CalculateTJXScore  ' FairValue.bas: DCF + PE/PB/ROE ratings for all TJX tickers
    End If
    
    ' Process main tasks
    Call ProcessAll
    Call FilterAndReport ' AllSignals
    
    LogMessage "RunScheduledTask completed successfully"
    
    Exit Sub
    
ErrorHandler:
    LogMessage "Error in RunScheduledTask: " & Err.Description
    ' Still schedule next run even if there was an error

End Sub

Private Function GetNextRunTime() As Date
    Dim nextRun As Date
    Dim currentDateTime As Date
    
    currentDateTime = Now
    nextRun = DateSerial(year(currentDateTime), month(currentDateTime), Day(currentDateTime)) + TimeValue(SCHEDULED_TIME)
    
    ' If it's past today's scheduled time, start checking from tomorrow
    If nextRun <= currentDateTime Then
        nextRun = nextRun + 1
    End If
    
    ' Skip weekends and holidays
    Do While Not IsWorkday(nextRun)
        nextRun = nextRun + 1
    Loop
    
    GetNextRunTime = nextRun
End Function

Private Function IsWorkday(ByVal checkDate As Date) As Boolean
    ' Check if it's a weekday (Monday-Friday) AND not a holiday
    IsWorkday = (weekday(checkDate, vbMonday) <= 5) And (Not IsHoliday(checkDate))
End Function

Private Function GetPreviousWorkday(ByVal fromDate As Date) As Date
    Dim previousDay As Date
    previousDay = fromDate - 1
    
    ' Keep going back until we find a workday
    Do While Not IsWorkday(previousDay)
        previousDay = previousDay - 1
    Loop
    
    GetPreviousWorkday = previousDay
End Function

Public Function GetNextWorkday(ByVal fromDate As Date) As Date
    Dim nextDay As Date
    nextDay = fromDate
    
    ' Keep going forward until we find a workday
    Do While Not IsWorkday(nextDay)
        nextDay = nextDay + 1
    Loop
    
    GetNextWorkday = nextDay
End Function

Private Function IsHoliday(ByVal checkDate As Date) As Boolean
    Dim holidays() As Date
    Dim i As Long
    Dim yearValue As Integer
    
    yearValue = year(checkDate)
    
    ' List of US Federal Holidays
    ReDim holidays(0 To 9)
    holidays(0) = DateSerial(yearValue, 1, 1)           ' New Year's Day
    holidays(1) = GetMLKDay(yearValue)                  ' Martin Luther King Jr. Day
    holidays(2) = GetPresidentsDay(yearValue)           ' Presidents' Day
    holidays(3) = GetMemorialDay(yearValue)             ' Memorial Day
    holidays(4) = DateSerial(yearValue, 7, 4)           ' Independence Day
    holidays(5) = GetLaborDay(yearValue)                ' Labor Day
    holidays(6) = GetColumbusDay(yearValue)             ' Columbus Day
    holidays(7) = DateSerial(yearValue, 11, 11)         ' Veterans Day
    holidays(8) = GetThanksgivingDay(yearValue)         ' Thanksgiving Day
    holidays(9) = DateSerial(yearValue, 12, 25)         ' Christmas Day
    
    ' Handle holidays that fall on weekends (observed on different days)
    For i = LBound(holidays) To UBound(holidays)
        Dim observedDate As Date
        observedDate = GetObservedHolidate(holidays(i))
        If checkDate = observedDate Then
            IsHoliday = True
            Exit Function
        End If
    Next i
    
    IsHoliday = False
End Function

Private Function GetObservedHolidate(ByVal holidayDate As Date) As Date
    ' Handle holidays that are observed on different days when they fall on weekends
    Select Case weekday(holidayDate, vbSunday)
        Case 1 ' Sunday - observed on Monday
            GetObservedHolidate = holidayDate + 1
        Case 7 ' Saturday - observed on Friday
            GetObservedHolidate = holidayDate - 1
        Case Else ' Weekday - observed on actual day
            GetObservedHolidate = holidayDate
    End Select
End Function

Private Function GetMLKDay(ByVal year As Integer) As Date
    ' Third Monday in January
    GetMLKDay = DateSerial(year, 1, Application.WorksheetFunction.RoundUp((20 - weekday(DateSerial(year, 1, 1), vbMonday)) / 7, 0) * 7 + 2)
End Function

Private Function GetPresidentsDay(ByVal year As Integer) As Date
    ' Third Monday in February
    GetPresidentsDay = DateSerial(year, 2, Application.WorksheetFunction.RoundUp((21 - weekday(DateSerial(year, 2, 1), vbMonday)) / 7, 0) * 7 + 1)
End Function

Private Function GetMemorialDay(ByVal year As Integer) As Date
    ' Last Monday in May
    GetMemorialDay = DateSerial(year, 5, Application.WorksheetFunction.RoundDown((31 - weekday(DateSerial(year, 5, 31), vbMonday)) / 7, 0) * 7 + 25)
End Function

Private Function GetLaborDay(ByVal year As Integer) As Date
    ' First Monday in September
    GetLaborDay = DateSerial(year, 9, Application.WorksheetFunction.RoundUp((7 - weekday(DateSerial(year, 9, 1), vbMonday)) / 7, 0) * 7 + 1)
End Function

Private Function GetColumbusDay(ByVal year As Integer) As Date
    ' Second Monday in October
    GetColumbusDay = DateSerial(year, 10, Application.WorksheetFunction.RoundUp((14 - weekday(DateSerial(year, 10, 1), vbMonday)) / 7, 0) * 7 + 1)
End Function

Private Function GetThanksgivingDay(ByVal year As Integer) As Date
    ' Fourth Thursday in November
    GetThanksgivingDay = DateSerial(year, 11, Application.WorksheetFunction.RoundUp((28 - weekday(DateSerial(year, 11, 1), vbThursday)) / 7, 0) * 7 + 1)
End Function

' Helper function to check if worksheet exists
Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not (ThisWorkbook.Sheets(sheetName) Is Nothing)
    On Error GoTo 0
End Function

' Logging function (optional but recommended)
Private Sub LogMessage(ByVal message As String)
    On Error Resume Next
    
    ' Try to log to a worksheet (create if it doesn't exist)
    Dim logSheet As Worksheet
    Set logSheet = ThisWorkbook.Sheets(LOG_SHEET)
    
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Sheets.Add
        logSheet.Name = LOG_SHEET
        logSheet.Range("A1").Value = "Timestamp"
        logSheet.Range("B1").Value = "Message"
    End If
    
    Dim lastRow As Long
    lastRow = logSheet.Cells(logSheet.Rows.count, 1).End(xlUp).row + 1
    
    logSheet.Cells(lastRow, 1).Value = Now
    logSheet.Cells(lastRow, 2).Value = message
    
    On Error GoTo 0
End Sub

' Helper functions for floating holidays
Private Function GetLastMondayOfMonth(ByVal year As Integer, ByVal month As Integer) As Date
    Dim lastDay As Date
    lastDay = DateSerial(year, month + 1, 0) ' Last day of the month
    
    ' Find the last Monday
    Do While weekday(lastDay, vbMonday) <> 1
        lastDay = lastDay - 1
    Loop
    
    GetLastMondayOfMonth = lastDay
End Function

Private Function GetFirstMondayOfMonth(ByVal year As Integer, ByVal month As Integer) As Date
    Dim firstDay As Date
    firstDay = DateSerial(year, month, 1)
    
    ' Find the first Monday
    Do While weekday(firstDay, vbMonday) <> 1
        firstDay = firstDay + 1
    Loop
    
    GetFirstMondayOfMonth = firstDay
End Function

Private Function GetNthWeekdayOfMonth(ByVal year As Integer, ByVal month As Integer, ByVal weekday As Integer, ByVal occurrence As Integer) As Date
    Dim firstDay As Date
    Dim targetDay As Date
    Dim count As Integer
    
    firstDay = DateSerial(year, month, 1)
    targetDay = firstDay
    count = 0
    
    ' Find the nth occurrence of the specified weekday
    Do While count < occurrence And month(targetDay) = month
        If weekday(targetDay, vbSunday) = weekday Then
            count = count + 1
            If count = occurrence Then
                GetNthWeekdayOfMonth = targetDay
                Exit Function
            End If
        End If
        targetDay = targetDay + 1
    Loop
    
    ' If not found, return 0
    GetNthWeekdayOfMonth = 0
End Function

' Call this sub to stop all scheduled tasks (useful for maintenance)
Sub StopScheduledTasks()
    On Error Resume Next
    Application.OnTime GetNextRunTime(), "RunScheduledTask", , False
    On Error GoTo 0
    LogMessage "Scheduled tasks stopped by user"
End Sub
Sub SaturdaySchedule()
    On Error Resume Next
    pubNotice = True
    Application.OnTime TimeValue(SCHEDULED_TIME), "ProcessALL"
    On Error GoTo 0
End Sub



