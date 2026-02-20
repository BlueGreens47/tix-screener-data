Attribute VB_Name = "Filtering"
Option Explicit
Sub ShowGroupForm()
    groupForm.Show
    'ResetApplicationSettings
    With ThisWorkbook.Sheets("DashBoard")
        .Range("W5").Value = 3
        .Range("W6").Value = 0 'Reset Iterations
        '.Range("Y5").Value = ThisWorkbook.Sheets("TJX").Range("E1").Value2
        '.Range("Y6").Value = ThisWorkbook.Sheets("TJX").Range("C1").Value2
    End With
End Sub

Sub FilterAndReport()
    Dim wsTJX As Worksheet, wsDash As Worksheet, wsRptLog As Worksheet, wsRptHist As Worksheet, wsTLog As Worksheet, wsRpt As Worksheet
    Dim lastRowTJX As Long, tickerCount As Long, i As Long
    Dim groupSize As Long
    Dim priceThreshold As Double, minpriceThreshold As Double
    Dim analysisDate As Date, minScore As Variant
    Dim filterArray() As Variant
    Dim startTime As Double
    Dim totalIterations As Long
    Dim currentIteration As Long
    Dim testDate As Date
       
    DoEvents
    If gStopMacro Then
        MsgBox "...E-Stopped!", vbInformation
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    startTime = Timer
    
    ' OPTIMIZATION: Disable all unnecessary features
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Set&Prep worksheets
    With ThisWorkbook
        Set wsTJX = .Sheets("TJX")
        Set wsDash = .Sheets("DashBoard")
        Set wsRptLog = .Sheets("ReportLog")
        Set wsRptHist = .Sheets("ReportHistory")
        Set wsTLog = .Sheets("TRADE LOG")
        Set wsRpt = .Sheets("Reports")
    End With
    
    Call ClearAllFilters
    
    ' Define constants
    groupSize = 50
    
    ' Get parameters
    minScore = wsDash.Range("W5").Value
    minpriceThreshold = wsDash.Range("Y6").Value
    priceThreshold = wsDash.Range("Y5").Value
    'priceThreshold = IIf(pubNotice Or perfTest, 100, wsTJX.Range("E1").value)
    
    analysisDate = wsDash.Range("H5").Value
    
    wsDash.Range("I4,N4").Value = "off"
    
    'Debug.Print "Analysis Date: " & analysisDate & " minScore: " & minScore & ", minPrice: " & minpriceThreshold & ", maxPrice: " & priceThreshold

    Dim skipPrompt As Boolean
    skipPrompt = pubNotice Or perfTest
    
    If Not skipPrompt Then
        If Not ConfirmProcessing() Then
            If Not GetUserInputs(minScore, minpriceThreshold, priceThreshold, analysisDate) Then
                Exit Sub
            End If
        End If
    End If

    ' OPTIMIZATION: Clear larger range in one operation
    wsRptHist.Range("A4:J100").ClearContents
    wsTLog.Range("B4:AF53").ClearContents
    
    ' OPTIMIZATION: Minimize dashboard operations and reduce range copying
    With wsDash
        .Range("A8:A57").ClearContents
        .Range("H5").Value = analysisDate
        .Range("W5").Value = minScore

        .Range("Y5").Value = CStr(priceThreshold)
        .Range("Y6").Value = CStr(minpriceThreshold)
        
        ' OPTIMIZATION: Only copy formulas once for the batch size we need
        .Range("A3:AQ3").Copy
        .Range("A8:AQ57").PasteSpecial Paste:=xlPasteFormulas
    End With
    Application.CutCopyMode = False
    
    ' OPTIMIZATION: Read all ticker data at once
    lastRowTJX = wsTJX.Cells(wsTJX.Rows.count, "A").End(xlUp).row
    tickerCount = lastRowTJX - 2
    
    ' Read all tickers into array in one operation
    Dim tickerData As Variant
    tickerData = wsTJX.Range("A3:A" & lastRowTJX).Value
    
    ReDim filterArray(1 To tickerCount)
    For i = 1 To tickerCount
        filterArray(i) = tickerData(i, 1)
    Next i

    totalIterations = Application.WorksheetFunction.Ceiling(tickerCount / groupSize, 1)
    currentIteration = 0

    ' OPTIMIZATION: Pre-allocate result collection array
    Dim allResults() As Variant
    Dim totalResultCount As Long
    ReDim allResults(1 To tickerCount, 1 To 5) ' Maximum possible size
    totalResultCount = 0

    ' Process in batches
    Dim batchStart As Long: batchStart = 1
    Do While batchStart <= tickerCount
        currentIteration = currentIteration + 1
        
        ' OPTIMIZATION: Update status less frequently
        If currentIteration Mod 10 = 0 Or currentIteration = totalIterations Then
            Application.StatusBar = "Processing... " & currentIteration & " of " & totalIterations & " (" & Format((currentIteration / totalIterations) * 100, "0") & "%)"
        End If

        Dim batchEnd As Long
        batchEnd = WorksheetFunction.min(batchStart + groupSize - 1, tickerCount)
        Dim batchSize As Long
        batchSize = batchEnd - batchStart + 1
        
        ' OPTIMIZATION: Clear only what we need
        wsDash.Range("A8:A57").ClearContents ' & (7 + batchSize)).ClearContents
        
        ' OPTIMIZATION: Load batch using range operations instead of loop
        Dim batchRange As Range
        Set batchRange = wsDash.Range("A8:A" & (7 + batchSize))
        
        Dim batchData() As Variant
        ReDim batchData(1 To batchSize, 1 To 1)
        
        For i = 1 To batchSize
            batchData(i, 1) = filterArray(batchStart + i - 1)
        Next i
        
        batchRange.Value = batchData

        Debug.Print "Processing batch " & currentIteration & "/" & totalIterations & " (" & batchStart & " to " & batchEnd & ")"

    '*** Fetch historical data for current batch
        Call ModDataFromBackup '(analysisDate)
    '***
        ' OPTIMIZATION: Calculate only once per batch
        Application.Calculation = xlCalculationAutomatic
        DoEvents ' Allow calculation to complete
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = True
        ' Process the batch and collect results
        Call ProcessTickersUltraFast(wsDash, analysisDate, minScore, minpriceThreshold, priceThreshold, batchSize, allResults, totalResultCount)
        
        batchStart = batchEnd + 1

        ' OPTIMIZATION: Reduce DoEvents frequency
        If currentIteration Mod 20 = 0 Then
            DoEvents
            If gStopMacro Then
                MsgBox "Macro stopped by user.", vbInformation
                GoTo CleanExit
            End If
        End If
    Loop

    ' OPTIMIZATION: Write all results at once at the end
    If totalResultCount > 0 Then
    Dim j As Integer
        Dim finalResults() As Variant
        ReDim finalResults(1 To totalResultCount, 1 To 5)
        
        For i = 1 To totalResultCount
            For j = 1 To 5
                finalResults(i, j) = allResults(i, j)
            Next j
        Next i
        
        wsRptHist.Range("A4").Resize(totalResultCount, 5).Value = finalResults
        Debug.Print "Total qualifying tickers: " & totalResultCount
    End If

    Application.Calculation = xlCalculationAutomatic
    
    Call theReporter
    Call UpdateDashFromReport
    
CleanExit:
    'If Not skipPrompt The
        Call DisplayCompletionMessage(startTime)
        If wsRpt.Range("B2").Value2 > 0 Then Call toggleDashFilter
        wsDash.Range("Y5").Value = priceThreshold
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    'End If
    
ErrorHandler:
    HandleProcessingError "FilterAndReport", Err
    Resume Next
    GoTo CleanExit
End Sub

' ULTRA-OPTIMIZED VERSION: ProcessTickersUltraFast
Sub ProcessTickersUltraFast(wsDash As Worksheet, aDate As Date, minScore As Variant, minpriceThreshold As Double, priceThreshold As Double, batchSize As Long, ByRef allResults() As Variant, ByRef totalResultCount As Long)
    Dim i As Long, j As Long
    Dim minScoreVal As Double
    Dim minCompScore As String
    
    
    minScoreVal = CDbl(minScore)

    Dim minCompScoreVal As Double
    minCompScoreVal = CDbl(wsDash.Range("S5").Value)

    ' OPTIMIZATION: Read all required data in single operations
    Dim tickers As Variant, values As Variant, prices As Variant, company As Variant, compScore As Variant

    tickers = wsDash.Range("A8:A" & (7 + batchSize)).Value
    company = wsDash.Range("B8:B" & (7 + batchSize)).Value
    values = wsDash.Range("G8:P" & (7 + batchSize)).Value
    prices = wsDash.Range("C8:C" & (7 + batchSize)).Value
    compScore = wsDash.Range("S8:S" & (7 + batchSize)).Value ' added for CompScore

    ' Process each ticker with optimized logic
    For i = 1 To batchSize
        If Not IsEmpty(tickers(i, 1)) And IsNumeric(prices(i, 1)) Then
            Dim price As Double, cScoreVal As Double ' added for CompScore
            price = CDbl(prices(i, 1))
            cScoreVal = IIf(IsNumeric(compScore(i, 1)), CDbl(compScore(i, 1)), 0) ' added for CompScore

            ' OPTIMIZATION: Quick price filter first (>= so scores above threshold pass)
            If price >= minpriceThreshold And price <= priceThreshold And cScoreVal >= minCompScoreVal Then ' added for CompScore
                ' OPTIMIZATION: Efficient score calculation
                Dim sumCL As Double, countCL As Double
                sumCL = 0: countCL = 0
                
                ' OPTIMIZATION: Unrolled and optimized loop
                Dim val1 As Double, val2 As Double, val3 As Double, val4 As Double, val5 As Double
                Dim val6 As Double, val7 As Double, val8 As Double, val9 As Double, val10 As Double
                
                If IsNumeric(values(i, 1)) Then
                    val1 = CDbl(values(i, 1))
                    If val1 <= -1 Or val1 >= 1 Then: sumCL = sumCL + val1: countCL = countCL + 1
                End If
                If IsNumeric(values(i, 2)) Then
                    val2 = CDbl(values(i, 2))
                    If val2 <= -1 Or val2 >= 1 Then: sumCL = sumCL + val2: countCL = countCL + 1
                End If
                If IsNumeric(values(i, 3)) Then
                    val3 = CDbl(values(i, 3))
                    If val3 <= -1 Or val3 >= 1 Then: sumCL = sumCL + val3: countCL = countCL + 1
                End If
                If IsNumeric(values(i, 4)) Then
                    val4 = CDbl(values(i, 4))
                    If val4 <= -1 Or val4 >= 1 Then: sumCL = sumCL + val4: countCL = countCL + 1
                End If
                If IsNumeric(values(i, 5)) Then
                    val5 = CDbl(values(i, 5))
                    If val5 <= -1 Or val5 >= 1 Then: sumCL = sumCL + val5: countCL = countCL + 1
                End If
                If IsNumeric(values(i, 6)) Then
                    val6 = CDbl(values(i, 6))
                    If val6 <= -1 Or val6 >= 1 Then: sumCL = sumCL + val6: countCL = countCL + 1
                End If
                If IsNumeric(values(i, 7)) Then
                    val7 = CDbl(values(i, 7))
                    If val7 <= -1 Or val7 >= 1 Then: sumCL = sumCL + val7: countCL = countCL + 1
                End If
                If IsNumeric(values(i, 8)) Then
                    val8 = CDbl(values(i, 8))
                    If val8 <= -1 Or val8 >= 1 Then: sumCL = sumCL + val8: countCL = countCL + 1
                End If
                If IsNumeric(values(i, 9)) Then
                    val9 = CDbl(values(i, 9))
                    If val9 <= -1 Or val9 >= 1 Then: sumCL = sumCL + val9: countCL = countCL + 1
                End If
                If IsNumeric(values(i, 10)) Then
                    val10 = CDbl(values(i, 10))
                    If val10 <= -1 Or val10 >= 1 Then: sumCL = sumCL + val10: countCL = countCL + 1
                End If
                
                ' Check if qualifies
                If sumCL >= minScoreVal Or sumCL <= minScoreVal * -1 Then ' added Or for minScore * -1 to include <=minScore
                'If countCL >= minScoreVal Then
                    totalResultCount = totalResultCount + 1
                    
                    ' Store result in master array
                    allResults(totalResultCount, 1) = aDate
                    allResults(totalResultCount, 2) = tickers(i, 1)
                    allResults(totalResultCount, 3) = sumCL 'countCL * Sgn(sumCL)
                    allResults(totalResultCount, 4) = company(i, 1)
                    allResults(totalResultCount, 5) = price
                End If
            End If
        End If
    Next i
End Sub

' OPTIMIZATION: Streamlined error handling
Private Sub HandleProcessingError(procedureName As String, errObj As ErrObject)
    Debug.Print "Error in " & procedureName & ": " & errObj.Description
    ResetApplicationSettings
End Sub

' OPTIMIZATION: Improved GetUserInputs with better error handling
Public Function GetUserInputs(ByRef minScore As Variant, ByRef minPrice As Double, ByRef maxPrice As Double, ByRef analysisDate As Date) As Boolean
    On Error GoTo InputError
    
    Dim tempValue As String
    
    ' More efficient input handling
    tempValue = InputBox("Score:", "Minimum Score ", CStr(minScore))
    If tempValue = "" Then GoTo InputError
    minScore = CDbl(tempValue)
    
    tempValue = InputBox("Enter minimum Price:", "Min Value", CStr(minPrice))
    If tempValue = "" Then GoTo InputError
    minPrice = CDbl(tempValue)
    If minPrice <= 0 Then GoTo InputError
    
    tempValue = InputBox("Enter maximum Price:", "Max Value", CStr(maxPrice))
    If tempValue = "" Then GoTo InputError
    maxPrice = CDbl(tempValue)
    If maxPrice <= 0 Then GoTo InputError
    
    tempValue = InputBox("Enter the analysis date (optional, leave blank for today's date):", "Analysis Date", Format(analysisDate, "yyyy-mm-dd"))
    If tempValue <> "" Then
        Dim testDate As Date
        testDate = CDate(tempValue)
        If testDate <= GetPreviousWorkday(Date) Then
            analysisDate = Application.WorkDay(testDate, 0)
        End If
    End If
    
    GetUserInputs = True
    Exit Function
    
InputError:
    MsgBox "Invalid input. Please try again.", vbExclamation
    GetUserInputs = False
End Function

' ***DFBU - ***********
Sub DataFromBackup(analysisDate As Date)
    Dim wsTo As Worksheet, wsFrom As Worksheet, wsDash As Worksheet
    Dim lastFromRow As Long, lastDashRow As Long, lastToRow As Long
    Dim endDate As Date, startDate As Date
    Dim tickerList As Collection
    Dim i As Long, j As Long
    Dim dataArray As Variant, sortedData As Variant
    Dim rowToAdd As Variant
    Dim dictUnique As Object
    Dim isWeeklyAnalysis As Boolean
    
    On Error Resume Next
    
    ' Disable screen updates and calculation for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Set worksheet references
    Set wsTo = ThisWorkbook.Worksheets("Data")
    Set wsDash = ThisWorkbook.Worksheets("DashBoard")
    
    ' Determine analysis type
    isWeeklyAnalysis = (wsDash.Range("H1").Value = "WEEKLY")
    
    ' Get date range based on analysis type - ALWAYS USE BackupAll worksheet
    Set wsFrom = ThisWorkbook.Worksheets("BackupAll")
    
    ' FIXED: Proper date range calculation
    endDate = analysisDate
    If isWeeklyAnalysis Then
        startDate = endDate - (50) ' 50 weeks back
        '' Debug.Print "WEEKLY: Date range " & startDate & " to " & endDate
    Else
        startDate = endDate - 50  ' 50 days back for daily
        '' Debug.Print "DAILY: Date range " & startDate & " to " & endDate
    End If
    
    ' Clear previous report data
    wsTo.Range("A:G").ClearContents
    
    ' Get list of tickers from DashBoard (with validation)
    lastDashRow = wsDash.Cells(wsDash.Rows.count, "A").End(xlUp).row
    Set tickerList = New Collection
    For i = 8 To lastDashRow
        If Not IsEmpty(wsDash.Range("A" & i).Value) And Len(Trim(wsDash.Range("A" & i).Value)) > 0 Then
            On Error Resume Next
            tickerList.Add Trim(wsDash.Range("A" & i).Value)
            On Error GoTo 0
        End If
    Next i
    
    ' ' Debug: Print ticker list count
     'Debug.Print "DfromBU Found " & tickerList.count & " tickers to process"
    
    ' Exit if no tickers found
    If tickerList.count < 1 Then GoTo FinalCleanup
    
    ' Load data from BackupAll into an array (with validation)
    lastFromRow = wsFrom.Cells(wsFrom.Rows.count, "A").End(xlUp).row
    If lastFromRow < 2 Then
        '' Debug.Print "No data found in BackupAll sheet"
        GoTo FinalCleanup ' No data to process
    End If
    
    '' Debug.Print "Loading " & lastFromRow & " rows from BackupAll"
    dataArray = wsFrom.Range("A2:G" & lastFromRow).Value ' Assuming headers in row 1
    
    ' Initialize a dictionary for unique rows
    Set dictUnique = CreateObject("Scripting.Dictionary")
    
    ' Process each ticker and filter data in memory
    Dim currentDate As Date
    Dim dateConditionMet As Boolean
    Dim key As String
    Dim matchCount As Long: matchCount = 0
    Dim ticker As Variant
    For Each ticker In tickerList
        For j = 1 To UBound(dataArray, 1)
            ' Check if dataArray(j,7) is a string and matches the ticker
            If TypeName(dataArray(j, 7)) = "String" And dataArray(j, 7) = ticker Then
                ' Check if dataArray(j,1) is a valid date
                If IsDate(dataArray(j, 1)) Then
                    currentDate = CDate(dataArray(j, 1))
                    
                    ' FIXED: Apply different date filtering logic based on analysis type
                    dateConditionMet = False
                    
                    If isWeeklyAnalysis Then
                        ' Weekly logic - only Mondays in range
                        If currentDate >= startDate And currentDate <= endDate Then ' And weekday(currentDate, vbMonday) = 1 Then
                            dateConditionMet = True
                        End If
                    Else
                        ' FIXED: Daily logic - any weekday in range (exclude weekends)
                        If currentDate >= startDate And currentDate <= endDate Then
                            ' For daily analysis, include all weekdays
                            If weekday(currentDate) <> vbSaturday And weekday(currentDate) <> vbSunday Then
                                dateConditionMet = True
                            End If
                        End If
                    End If
                    
                    If dateConditionMet Then
                        matchCount = matchCount + 1
                        ' Create a key for uniqueness based on columns A, B, E
                        key = CStr(dataArray(j, 1)) & "|" & CStr(dataArray(j, 2)) & "|" & CStr(dataArray(j, 5))
                        If Not dictUnique.Exists(key) Then
                            dictUnique.Add key, Array(dataArray(j, 1), dataArray(j, 2), dataArray(j, 3), _
                                                      dataArray(j, 4), dataArray(j, 5), dataArray(j, 6), dataArray(j, 7))
                        End If
                    End If
                End If
            End If
        Next j
        '' Debug.Print ticker
    Next ticker
    
    ' ' Debug: Print match statistics
    Debug.Print "Data from Backup Found " & matchCount & " matching records, " & dictUnique.count & " unique records"
    
    ' Convert dictionary to array for sorting
    If dictUnique.count = 0 Then
       ' ' Debug.Print "No data matches criteria - exiting"
        GoTo FinalCleanup
    End If
    
    ReDim sortedData(1 To dictUnique.count, 1 To 7)
    i = 1
    Dim dictKey As Variant
    For Each dictKey In dictUnique.Keys
        rowToAdd = dictUnique(dictKey)
        For j = 1 To 7
            sortedData(i, j) = rowToAdd(j - 1)
        Next j
        i = i + 1
    Next dictKey
    
    Call dataHeader(wsTo)
    
    ' Sort the array by Ticker (column 7) and then by Date (column 1)
    Call SortArrayWithExcel(sortedData, 7, 1)
    
    ' Write sorted data to Data sheet
    If UBound(sortedData, 1) > 50 Then
        lastToRow = wsTo.Cells(wsTo.Rows.count, "A").End(xlUp).row + 1
        wsTo.Range("A" & lastToRow).Resize(UBound(sortedData, 1), 7).Value = sortedData
        wsTo.Columns.AutoFit
        
        ' ' Debug: Confirm data written
       'Debug.Print "Written " & UBound(sortedData, 1) & " rows to Data sheet starting at row " & lastToRow
    End If
    
    Call DeleteNARows(wsTo)
    Call CalcATR
    
FinalCleanup:
    ' Re-enable application settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    

End Sub

Function GetNextMonday(currentDate As Date) As Date
    Dim dow As Long
    dow = weekday(currentDate, vbMonday)  ' vbMonday means Monday is treated as day 1.
    
    If dow = 1 Then
        ' If the date is Monday, Add 7 days to get the Next Monday.
        GetNextMonday = currentDate + 7
    Else
        ' Otherwise, Add the number of days since Monday.
        GetNextMonday = currentDate + (dow + 2)
    End If
End Function

Function GetPreviousMonday(currentDate As Date) As Date
    Dim dow As Long
    dow = weekday(currentDate, vbMonday)  ' vbMonday means Monday is treated as day 1.
    
    If dow = 1 Then
        ' If the date is Monday, subtract 7 days to get the previous Monday.
        GetPreviousMonday = currentDate - 7
    Else
        ' Otherwise, subtract the number of days since Monday.
        GetPreviousMonday = currentDate - (dow - 1)
    End If
End Function

Function GetPreviousWorkday(currentDate As Date) As Date
    Dim previousDay As Date
    previousDay = currentDate - 1
    ' Keep going back until we find a workday (not Saturday or Sunday)
    Do While weekday(previousDay) = vbSaturday Or weekday(previousDay) = vbSunday
        previousDay = previousDay - 1
    Loop
    GetPreviousWorkday = previousDay
End Function

Sub dataHeader(ws As Worksheet)
    Dim headers As Variant
    Dim col As Integer
    
    ws.Range("A:Z").ClearContents
    headers = Array("Date", "Open", "High", "Low", "Close", "Volume", "Ticker")
    
    For col = 0 To UBound(headers)
        ws.Cells(1, col + 1).Value = headers(col)
    Next col
End Sub

Sub DeleteNARows(ws As Worksheet)
    Dim lastRow As Long
    Dim deleteRange As Range
    Dim cell As Range

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    If lastRow <= 1 Then Exit Sub

    ' Check every cell for errors, "N/A", "#N/A", or empty values
    For Each cell In ws.Range("A2:G" & lastRow)
        If IsError(cell.Value) Then
            ' Add row to delete list if cell contains any error (e.g., #N/A)
            If deleteRange Is Nothing Then
                Set deleteRange = cell.EntireRow
            Else
                Set deleteRange = Union(deleteRange, cell.EntireRow)
            End If
        Else
            ' Check for text values: "N/A", "#N/A", or empty cells
            If cell.Value = "N/A" Or cell.Value = "#N/A" Or Trim(cell.Value) = "" Then
                If deleteRange Is Nothing Then
                    Set deleteRange = cell.EntireRow
                Else
                    Set deleteRange = Union(deleteRange, cell.EntireRow)
                End If
            End If
        End If
    Next cell

    ' Delete all identified rows in one operation
    If Not deleteRange Is Nothing Then deleteRange.Delete
End Sub

Sub SortArray(ByRef arr As Variant, ByVal sortCol1 As Integer, ByVal sortCol2 As Integer)
    Dim i As Long, j As Long, k As Long
    Dim temp As Variant
    Dim numCols As Long
    
    ' Get number of columns in the array
    numCols = UBound(arr, 2)
    
    ' Bubble sort with two-level sorting
    For i = LBound(arr, 1) To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            ' Compare primary sort column (sortCol1), then secondary (sortCol2)
            Dim needSwap As Boolean
            needSwap = False
            
            ' Primary sort comparison
            If arr(i, sortCol1) > arr(j, sortCol1) Then
                needSwap = True
            ElseIf arr(i, sortCol1) = arr(j, sortCol1) Then
                ' If primary columns are equal, check secondary sort
                If arr(i, sortCol2) > arr(j, sortCol2) Then
                    needSwap = True
                End If
            End If
            
            ' Swap entire rows if needed
            If needSwap Then
                ' Swap each column value in the two rows
                For k = 1 To numCols
                    temp = arr(i, k)
                    arr(i, k) = arr(j, k)
                    arr(j, k) = temp
                Next k
            End If
        Next j
    Next i
End Sub

' Alternative faster version using Excel's built-in sort
Sub SortArrayWithExcel(ByRef arr As Variant, ByVal sortCol1 As Integer, ByVal sortCol2 As Integer)
    Dim tempWs As Worksheet
    Dim tempRange As Range
    Dim lastRow As Long, lastCol As Long
    
    ' Create temporary worksheet
    Set tempWs = ThisWorkbook.Worksheets.Add
    tempWs.Name = "TempSort_" & Format(Now, "hhmmss")
    
    ' Write array to temporary worksheet
    lastRow = UBound(arr, 1)
    lastCol = UBound(arr, 2)
    tempWs.Range("A1").Resize(lastRow, lastCol).Value = arr
    
    ' Sort using Excel's built-in sort
    Set tempRange = tempWs.Range("A1").Resize(lastRow, lastCol)
    tempRange.Sort key1:=tempWs.Columns(sortCol1), Order1:=xlAscending, _
                   Key2:=tempWs.Columns(sortCol2), Order2:=xlAscending, _
                   Header:=xlNo
    
    ' Read sorted data back into array
    arr = tempRange.Value
    
    ' Clean up temporary worksheet
    Application.DisplayAlerts = False
    tempWs.Delete
    Application.DisplayAlerts = True
    
    Set tempWs = Nothing
    Set tempRange = Nothing
End Sub

Sub GroupbyPrice()
       
    'Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim wsTJX As Worksheet
    Dim wsDashboard As Worksheet
    Dim i As Long, k As Long
    Dim lastRow As Long, groupSize As Long
    Dim visibleRows() As Long
    Dim answer As VbMsgBoxResult
    Dim startTime As Double
    Dim groupStart As Long, groupEnd As Long
    Dim firstRow As Long, lastRowGroup As Long
    Dim arr()
    Dim m As Long, n As Long
    Dim minValue As Variant, maxValue As Variant
    Dim analysisDate As Date
    
    startTime = Timer
    gStopMacro = False
    
    On Error GoTo ErrorHandler
        
    ' Set references to the worksheets
    Set wsTJX = ThisWorkbook.Sheets("TJX")
    Set wsDashboard = ThisWorkbook.Sheets("DashBoard")
    
    groupSize = 50
    analysisDate = wsDashboard.Range("H5")
    
    ' Clear any existing filters
    wsTJX.AutoFilterMode = False
    wsDashboard.AutoFilterMode = False
    
    ' Check if there are existing tickers on Dashboard and ask user what to do
    Dim existingTickers As Boolean
    existingTickers = False
    
    ' Check if there are any existing tickers in A8:A57
    For i = 8 To 57
        If wsDashboard.Cells(i, 1).Value <> "" Then
            existingTickers = True
            Exit For
        End If
    Next i
    
    ' If there are existing tickers, ask user what to do with them
    If existingTickers Then
        answer = MsgBox("There are existing tickers on the Dashboard." & vbCrLf & vbCrLf & "Select 'Yes' to analyze these existing tickers." & vbCrLf & vbCrLf & "'No' will clear them and continue with TJX filtering." & vbCrLf & vbCrLf & "'Cancel' to Exit.", vbYesNoCancel + vbQuestion, "Process Existing Tickers?")
        
        Select Case answer
            Case vbYes
                ' User wants to analyze existing tickers
                Call ModDataFromBackup 'Call DataFromBackup(analysisDate)
                GoTo CleanExit
            Case vbNo
                ' User wants to clear existing tickers and continue
                wsDashboard.Range("A8:A57").ClearContents
            Case Else
                ' User wants to cancel
                GoTo CleanExit
        End Select
    End If
    
    ' Get the data range
    lastRow = wsTJX.Cells(wsTJX.Rows.count, "A").End(xlUp).row
    
    With wsDashboard
            .Range("A3:AP3").Copy
            .Range("A8:AP57").PasteSpecial Paste:=xlPasteFormulas
            .Range("A8:AP57").Font.Name = "Calibri"
            .Range("A8:AP57").Font.Size = 8
    End With
    
    
    ' Get user input for filter criteria
    minValue = InputBox("Enter minimum Price:", "Min Value", wsTJX.Range("C1").Value)
    If minValue = "" Then GoTo CleanExit  ' User clicked Cancel
    
    maxValue = InputBox("Enter maximum Price:", "Max Value", wsTJX.Range("E1").Value)
    If maxValue = "" Then GoTo CleanExit  ' User clicked Cancel
    
    ' Validate input
    If Not IsNumeric(minValue) Or Not IsNumeric(maxValue) Then
        MsgBox "Please enter valid numeric values!", vbExclamation
        GoTo CleanExit
    End If
    
    ' Convert to numbers
    minValue = CDbl(minValue)
    maxValue = CDbl(maxValue)
    
    ' Validate range
    If minValue > maxValue Then
        MsgBox "Minimum value cannot be greater than maximum value!", vbExclamation
        GoTo CleanExit
    End If
    
    ' Apply the filter with user-input values
    wsTJX.Range("A3:G" & lastRow).AutoFilter Field:=4, Criteria1:=">=" & minValue, Operator:=xlAnd, Criteria2:="<=" & maxValue
    
    ' Get all visible row numbers
    k = 1
    For i = 3 To lastRow
        If Not wsTJX.Rows(i).Hidden Then
            If k = 1 Then
                ReDim visibleRows(1 To 1)
            Else
                ReDim Preserve visibleRows(1 To k)
            End If
            visibleRows(k) = i
            k = k + 1
        End If
    Next i
    
    If k - 1 = 0 Then
        MsgBox "No data matches the filter criteria!" & vbCrLf & "Filter range: " & minValue & " to " & maxValue, vbExclamation
        GoTo CleanExit
    End If
    
    ' Process visible rows in groups of 50
    groupStart = 1
    Do While groupStart <= k - 1
        groupEnd = groupStart + groupSize - 1
        If groupEnd > k - 1 Then groupEnd = k - 1
        
        ' Create a range for the group
        firstRow = visibleRows(groupStart)
        lastRowGroup = visibleRows(groupEnd)
        Dim tickerGroup As Range
        Set tickerGroup = wsTJX.Range("A" & firstRow & ":A" & lastRowGroup)
        
        ' Copy the visible tickers to Dashboard FIRST (before asking user)
        ReDim arr(1 To groupEnd - groupStart + 1, 1 To 1)
        m = 1
        For n = groupStart To groupEnd
            arr(m, 1) = wsTJX.Range("A" & visibleRows(n)).Value
            m = m + 1
        Next n
        ' Clear previous contents and copy new tickers
        wsDashboard.Range("A8:A57").ClearContents
        wsDashboard.Range("A8").Resize(UBound(arr, 1), 1).Value = arr
        
        ' Ask user for action AFTER displaying the tickers
        answer = MsgBox("Select 'Yes' to analyze onscreen Tickers." & vbCrLf & vbCrLf & "'No' will Skip this Group." & vbCrLf & vbCrLf & "'Cancel' to Exit.", vbYesNoCancel + vbQuestion, "Update, Copy or Cancel?")
        
        Select Case answer
            Case vbYes
                ' User wants to analyze - call the backup function and exit
                Call ModDataFromBackup 'Call DataFromBackup(analysisDate)
                wsTJX.AutoFilterMode = False
                GoTo CleanExit
            Case vbNo
                ' User wants to skip this group - continue to next group
                ' (Tickers are already displayed, just move on)
            Case Else
                ' User wants to cancel
                wsTJX.AutoFilterMode = False
                GoTo CleanExit
        End Select
        
        ' Move to next group
        groupStart = groupEnd + 1
        
        If gStopMacro Then
            MsgBox "Macro stopped by user.", vbInformation
            'ResetApplicationSettings
            GoTo CleanExit
        End If
    Loop
    
CleanExit:
    ' Turn off filter and restore settings
   ' wsTJX.AutoFilterMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    If Not pubNotice Then MsgBox "Filtering completed in " & Round((Timer - startTime) / 60, 2) & "m.", vbExclamation
    wsDashboard.Range("$A$7:$AE$57").AutoFilter  ' .AutoFilter Field:=9, Criteria1:="1"
    
    'ShowGroupForm
    
    Exit Sub

ErrorHandler:
HandleProcessingError "GroupByPrice", Err
    MsgBox "An error occurred in GroupByPrice: " & Err.Description, vbCritical

End Sub

Sub TheOnScreenGroup()
       
    Dim wsTJX As Worksheet
    Dim wsDashboard As Worksheet
    Dim i As Long, k As Long
    Dim lastRow As Long, groupSize As Long
    Dim visibleRows() As Long
    Dim answer As VbMsgBoxResult
    Dim startTime As Double
    Dim groupStart As Long, groupEnd As Long
    Dim firstRow As Long, lastRowGroup As Long
    Dim arr()
    Dim m As Long, n As Long
    Dim minValue As Variant, maxValue As Variant
    Dim analysisDate As Date
    
    startTime = Timer
    gStopMacro = False
    
    On Error GoTo ErrorHandler
        
    ' Set references to the worksheets
    Set wsTJX = ThisWorkbook.Sheets("TJX")
    Set wsDashboard = ThisWorkbook.Sheets("DashBoard")
    
    groupSize = 50
    analysisDate = wsDashboard.Range("H5")
    
    ' Clear any existing filters
    wsTJX.AutoFilterMode = False
    wsDashboard.AutoFilterMode = False
    
    ' Check if there are existing tickers on Dashboard and ask user what to do
    Dim existingTickers As Boolean
    existingTickers = False
    
    ' Check if there are any existing tickers in A8:A57
    For i = 8 To 57
        If wsDashboard.Cells(i, 1).Value <> "" Then
            existingTickers = True
            Exit For
        End If
    Next i
    
    ' If there are existing tickers, ask user what to do with them
    If existingTickers Then
        answer = MsgBox("There are existing tickers on the Dashboard." & vbCrLf & vbCrLf & "Select 'Yes' to analyze these existing tickers." & vbCrLf & vbCrLf & "'No' will clear them and continue with TJX filtering." & vbCrLf & vbCrLf & "'Cancel' to Exit.", vbYesNoCancel + vbQuestion, "Process Existing Tickers?")
        
        Select Case answer
            Case vbYes
                ' User wants to analyze existing tickers
                Call DataFromBackup(analysisDate)
                GoTo CleanExit
            Case vbNo
                ' User wants to clear existing tickers and continue
                wsDashboard.Range("A8:A57").ClearContents
            Case Else
                ' User wants to cancel
                GoTo CleanExit
        End Select
    End If
    
    ' Get the data range
    lastRow = wsTJX.Cells(wsTJX.Rows.count, "A").End(xlUp).row
    
    With wsDashboard
            .Range("A3:AP3").Copy
            .Range("A8:AP57").PasteSpecial Paste:=xlPasteFormulas
            .Range("A8:AP57").Font.Name = "Calibri"
            .Range("A8:AP57").Font.Size = 8
    End With
    
    
    ' Get user input for filter criteria
    minValue = InputBox("Enter minimum Price:", "Min Value", wsTJX.Range("C1").Value)
    If minValue = "" Then GoTo CleanExit  ' User clicked Cancel
    
    maxValue = InputBox("Enter maximum Price:", "Max Value", wsTJX.Range("E1").Value)
    If maxValue = "" Then GoTo CleanExit  ' User clicked Cancel
    
    ' Validate input
    If Not IsNumeric(minValue) Or Not IsNumeric(maxValue) Then
        MsgBox "Please enter valid numeric values!", vbExclamation
        GoTo CleanExit
    End If
    
    ' Convert to numbers
    minValue = CDbl(minValue)
    maxValue = CDbl(maxValue)
    
    ' Validate range
    If minValue > maxValue Then
        MsgBox "Minimum value cannot be greater than maximum value!", vbExclamation
        GoTo CleanExit
    End If
    
    ' Apply the filter with user-input values
    wsTJX.Range("A3:G" & lastRow).AutoFilter Field:=4, Criteria1:=">=" & minValue, Operator:=xlAnd, Criteria2:="<=" & maxValue
    
    ' Get all visible row numbers
    k = 1
    For i = 3 To lastRow
        If Not wsTJX.Rows(i).Hidden Then
            If k = 1 Then
                ReDim visibleRows(1 To 1)
            Else
                ReDim Preserve visibleRows(1 To k)
            End If
            visibleRows(k) = i
            k = k + 1
        End If
    Next i
    
    If k - 1 = 0 Then
        MsgBox "No data matches the filter criteria!" & vbCrLf & "Filter range: " & minValue & " to " & maxValue, vbExclamation
        GoTo CleanExit
    End If
    
    ' Process visible rows in groups of 50
    groupStart = 1
    Do While groupStart <= k - 1
        groupEnd = groupStart + groupSize - 1
        If groupEnd > k - 1 Then groupEnd = k - 1
        
        ' Create a range for the group
        firstRow = visibleRows(groupStart)
        lastRowGroup = visibleRows(groupEnd)
        Dim tickerGroup As Range
        Set tickerGroup = wsTJX.Range("A" & firstRow & ":A" & lastRowGroup)
        
        ' Copy the visible tickers to Dashboard FIRST (before asking user)
        ReDim arr(1 To groupEnd - groupStart + 1, 1 To 1)
        m = 1
        For n = groupStart To groupEnd
            arr(m, 1) = wsTJX.Range("A" & visibleRows(n)).Value
            m = m + 1
        Next n
        ' Clear previous contents and copy new tickers
        wsDashboard.Range("A8:A57").ClearContents
        wsDashboard.Range("A8").Resize(UBound(arr, 1), 1).Value = arr
        
        ' Ask user for action AFTER displaying the tickers
        answer = MsgBox("Select 'Yes' to analyze onscreen Tickers." & vbCrLf & vbCrLf & "'No' will Skip this Group." & vbCrLf & vbCrLf & "'Cancel' to Exit.", vbYesNoCancel + vbQuestion, "Update, Copy or Cancel?")
        
        Select Case answer
            Case vbYes
                ' User wants to analyze - call the backup function and exit
                Call CheckTickersOnScreen
            Case vbNo
                ' User wants to skip this group - continue to next group
                ' (Tickers are already displayed, just move on)
            Case Else
                ' User wants to cancel
                wsTJX.AutoFilterMode = False
                GoTo CleanExit
        End Select
        
        ' Move to next group
        groupStart = groupEnd + 1
        
        If gStopMacro Then
            MsgBox "Macro stopped by user.", vbInformation
            'ResetApplicationSettings
            GoTo CleanExit
        End If
    Loop
    
CleanExit:
    ' Turn off filter and restore settings
   ' wsTJX.AutoFilterMode = False
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    If Not pubNotice Then MsgBox "On-Screen Tickers completed in " & Round((Timer - startTime) / 60, 2) & "m.", vbExclamation
   ' wsDashboard.Range("$A$7:$AE$57").AutoFilter  ' .AutoFilter Field:=9, Criteria1:="1"
    
    'Call theReporter
    
    ShowGroupForm
    
    Exit Sub

ErrorHandler:
HandleProcessingError "GroupByPrice", Err
    MsgBox "An error occurred in GroupByPrice: " & Err.Description, vbCritical

End Sub

Sub CheckTickersOnScreen()

    Dim wsTJX As Worksheet, wsDash As Worksheet, wsRptLog As Worksheet, wsRptHist As Worksheet, wsTLog As Worksheet
    Dim lastRow As Long, tickerCount As Long, i As Long
    Dim groupSize As Long
    Dim priceThreshold As Double, minpriceThreshold As Double
    Dim analysisDate As Date, minScore As Variant
    Dim filterArray() As Variant
    Dim startTime As Double
    Dim totalIterations As Long
    Dim currentIteration As Long
    Dim testDate As Date
    Dim tstPeriod As Variant
    
    startTime = Timer
        
    ' OPTIMIZATION: Disable all unnecessary features
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' Set&Prep worksheets
    With ThisWorkbook
        Set wsTJX = .Sheets("TJX")
        Set wsDash = .Sheets("DashBoard")
        Set wsRptLog = .Sheets("ReportLog")
        Set wsRptHist = .Sheets("ReportHistory")
        Set wsTLog = .Sheets("TRADE LOG")
    End With
    
    Call ClearAllFilters
    Call CreateWorksheets
    
    ' Define constants
    groupSize = 50
    
    ' Get parameters
    minScore = wsDash.Range("W5").Value
    minpriceThreshold = wsTJX.Range("C1").Value
    priceThreshold = wsTJX.Range("E1").Value
    analysisDate = wsDash.Range("H5").Value
   
    'Debug.Print "Analysis Date: " & analysisDate & " minScore: " & minScore & ", minPrice: " & minpriceThreshold & ", maxPrice: " & priceThreshold

    ' OPTIMIZATION: Clear larger range in one operation
    wsRptHist.Range("A4:M1000").ClearContents
    
    ' OPTIMIZATION: Minimize dashboard operations
    With wsDash
        .Range("B3:AP3").Copy
        .Range("B8:AP57").PasteSpecial Paste:=xlPasteFormulas
    End With
    Application.CutCopyMode = False
    
    ' OPTIMIZATION: Read all ticker data at once
    lastRow = wsDash.Cells(wsDash.Rows.count, "A").End(xlUp).row
    tickerCount = lastRow - 7
    
    ' Read all tickers into array in one operation
    Dim tickerData As Variant
    tickerData = wsDash.Range("A8:A" & lastRow).Value
    
    ReDim filterArray(1 To tickerCount)
    For i = 1 To tickerCount
        filterArray(i) = tickerData(i, 1)
    Next i

    ' OPTIMIZATION: Pre-allocate result collection array
    Dim allResults() As Variant
    Dim totalResultCount As Long
    ReDim allResults(1 To tickerCount, 1 To 5) ' Maximum possible size
    totalResultCount = 0

    ' Process in batches
    Dim batchStart As Long: batchStart = 1
    Do While batchStart <= tickerCount
        currentIteration = currentIteration + 1
        
        Dim batchEnd As Long
        batchEnd = WorksheetFunction.min(batchStart + groupSize - 1, tickerCount)
        Dim batchSize As Long
        batchSize = batchEnd - batchStart + 1
        
        ' OPTIMIZATION: Clear only what we need
        'wsDash.Range("A8:A" & (7 + batchSize)).ClearContents
        
        ' OPTIMIZATION: Load batch using range operations instead of loop
        Dim batchRange As Range
        Set batchRange = wsDash.Range("A8:A" & (7 + batchSize))
        
        Dim batchData() As Variant
        ReDim batchData(1 To batchSize, 1 To 1)
        
        For i = 1 To batchSize
            batchData(i, 1) = filterArray(batchStart + i - 1)
        Next i
        
        batchRange.Value = batchData

        'Debug.Print "Processing batch " & currentIteration & "/" & totalIterations & " (" & batchStart & " to " & batchEnd & ")"

        ' Fetch historical data for current batch
        Call DataFromBackup(analysisDate)
        
        ' OPTIMIZATION: Calculate only once per batch
        Application.Calculation = xlCalculationAutomatic
        DoEvents ' Allow calculation to complete
        Application.Calculation = xlCalculationManual
        
        ' Process the batch and collect results
        Call ProcessTickersUltraFast(wsDash, analysisDate, minScore, minpriceThreshold, priceThreshold, batchSize, allResults, totalResultCount)
        
        batchStart = batchEnd + 1

    Loop

    ' OPTIMIZATION: Write all results at once at the end
    If totalResultCount > 0 Then
        Dim j As Integer
        Dim finalResults() As Variant
        ReDim finalResults(1 To totalResultCount, 1 To 5)
        
        For i = 1 To totalResultCount
            For j = 1 To 5
                finalResults(i, j) = allResults(i, j)
            Next j
        Next i
        
        wsRptHist.Range("A4").Resize(totalResultCount, 5).Value = finalResults
        'Debug.Print "Total qualifying tickers: " & totalResultCount
    End If

    Application.Calculation = xlCalculationAutomatic

    Call SetupTradeLog(wsTLog, tickerCount)
   
    tstPeriod = Array(7, 14, 21, 28) ', 35, 42, 49, 56, 63, 70, 77, 84) ', 91, 98, 105, 112)
    For i = 0 To UBound(tstPeriod)
        If tstPeriod(i) + wsDash.Range("H5").Value > wsDash.Range("S5").Value Then Exit For
        wsTLog.Range("N2").Value = tstPeriod(i)
        perfTest = True
        Application.Calculate
        Call UpdateTradePerformance

    Next i
    
    Call AnalyzeTrades
    
    Sheets("Charts").Select
        
CleanExit:
    
    DisplayCompletionMessage (startTime)
 
    Exit Sub

ErrorHandler:
    HandleProcessingError "QuickTick", Err
    Resume Next
    GoTo CleanExit
    
End Sub



