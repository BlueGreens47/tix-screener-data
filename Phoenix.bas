Attribute VB_Name = "Phoenix"
Option Explicit
' ===== MAIN EXECUTION SUB =====
Sub TradeMaster()
    Dim startTime As Double
    startTime = Timer
    
    Call MasterDataFromBackup
    Call UpdateSystemWithATR  ' ATR zones + risk management (ATRCalculation.bas)
    Call CalculateIndicators  ' Fixed: was CalculateEnhancedIndicators (wrong column layout)
    
    'Call GenerateCompleteTradingSignals_Main
    'Call MasterWeeklyTradingStrategyAndSignals
    Call cweGenerateTradingSignals
    
    Sheets("cweSignals").Range("D:H").NumberFormat = "0.00_ ;[Red]-0.00 "
    Sheets("mainTradingSignals").Range("D:E,G:G,I:P").NumberFormat = "0.00_ ;[Red]-0.00 "
    Sheets("WeeklySignals").Range("E:G,K:T").NumberFormat = "0.00_ ;[Red]-0.00 "
 
    Application.EnableEvents = True
    Application.CALCULATION = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    msgbox "TradeMaster done in: " & Format((Timer - startTime) / 60, "0.00") & " minutes"
End Sub

Sub MasterDataFromBackup()
    Dim wsDashboard As Worksheet, wsTJX As Worksheet
    Dim wsBackupAll As Worksheet, wsData As Worksheet
    Dim maxPrice As Double, minPrice As Double, endDate As Date
    Dim lastRowTJX As Long, lastRowBackup As Long, lastRowData As Long
    Dim i As Long, j As Long, k As Long, outputRow As Long
    Dim ticker As String, tickerPrice As Double
    Dim tjxData As Variant, backupData As Variant
    Dim outputData() As Variant
    Dim tickerRecords() As Variant
    Dim recordCount As Long
    Dim frequency As String
    Dim validDates As Object
    Dim dateRange As Range, cell As Range
    Dim currentDate As Date
    
    Application.ScreenUpdating = False
    Application.CALCULATION = xlCalculationManual
    
    ' Set worksheet references
    Set wsDashboard = ThisWorkbook.Sheets("Dashboard")
    Set wsTJX = ThisWorkbook.Sheets("TJX")
    Set wsBackupAll = ThisWorkbook.Sheets("BackupAll")
    Set wsData = ThisWorkbook.Sheets("Data")
    wsData.Range("A2:V" & wsData.Rows.count).ClearContents
    
    ' Get filter criteria from Dashboard
    frequency = wsDashboard.Range("H1").value
    endDate = wsDashboard.Range("H5").value
    maxPrice = wsDashboard.Range("Y5").value
    minPrice = wsDashboard.Range("Y6").value
    minScore = wsDashboard.Range("W5").value
    
    Dim skipPrompt As Boolean
    skipPrompt = pubNotice Or perfTest
    
    If Not skipPrompt Then
        If Not ConfirmProcessing() Then
            If Not GetUserInputs(minScore, minPrice, maxPrice, endDate) Then
                Exit Sub
            End If
        End If
    End If
    
    ' Load valid dates based on frequency selection
    Set validDates = CreateObject("Scripting.Dictionary")
    
    If UCase(frequency) = "WEEKLY" Then
        Set dateRange = wsDashboard.Range("Weekly")
    ElseIf UCase(frequency) = "DAILY" Then
        Set dateRange = wsDashboard.Range("Daily")
    Else
        msgbox "Please select either 'DAILY' or 'WEEKLY' in cell H1", vbExclamation
        Application.ScreenUpdating = True
        Application.CALCULATION = xlCalculationAutomatic
        Exit Sub
    End If
    
    ' Build dictionary of valid dates for fast lookup
    For Each cell In dateRange
        If IsDate(cell.value) Then
            If cell.value <= endDate Then
                validDates(CLng(cell.value)) = True
            End If
        End If
    Next cell
    
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
                    currentDate = backupData(j, 1) ' Column 1 is Date
                    If currentDate <= endDate Then
                        ' For WEEKLY, only include Mondays; for DAILY, include all dates
                        Dim includeRecord As Boolean
                        includeRecord = False
                        
                        If UCase(frequency) = "WEEKLY" Then
                            ' Check if it's a Monday (Weekday = 2)
                            If weekday(currentDate, vbSunday) = 2 Then 'ThisWorkbook.Sheets("DashBoard").Range("L4").value Then
                                includeRecord = True
                            End If
                        Else ' DAILY
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
                    End If
                End If
            Next j
            
            ' Take the last 50 records (most recent)
            If recordCount > 0 Then
                Dim startIdx As Long
                startIdx = Application.WorksheetFunction.max(1, recordCount - 36)
                
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
        wsData.Range("A" & 2).Resize(outputIndex, 7).value = outputData
    End If
    
    Application.ScreenUpdating = True
    Application.CALCULATION = xlCalculationAutomatic
    
    'msgbox "UltraDataFromBackup complete! " & outputIndex & " records appended to the Data sheet using " & frequency & " frequency.", vbInformation
    
End Sub

Sub MasterWeeklyTradingStrategyAndSignals()
    Dim startTime As Double
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.CALCULATION = xlCalculationManual
    Application.EnableEvents = False
    
    Dim ws As Worksheet
    Set ws = Worksheets("Data")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Pre-fetch ALL data into memory arrays
    Dim allData As Variant
    allData = ws.Range("A2:V" & lastRow).value
    
    ' Create output array
    Dim signals() As Variant
    ReDim signals(1 To lastRow, 1 To 21)
    Dim signalCount As Long
    
    ' Process data from arrays
    Dim i As Long
    For i = 30 To lastRow
        If HasCompleteWeeklyData_Array(allData, i) Then
            Dim signal As String
            signal = GenerateWeeklySignal_Array(allData, i)
            
            If signal <> "HOLD" Then
                signalCount = signalCount + 1
                PopulateSignalArray_Array signals, signalCount, allData, i, signal
            End If
        End If
    Next i
    
    ' Output results
    If signalCount > 0 Then
        ' Create a new properly sized array
        Dim finalSignals() As Variant
        ReDim finalSignals(1 To signalCount, 1 To 21)
        
        ' Copy data to the properly sized array
        Dim copyI As Long, copyJ As Long
        For copyI = 1 To signalCount
            For copyJ = 1 To 21
                finalSignals(copyI, copyJ) = signals(copyI, copyJ)
            Next copyJ
        Next copyI
        
        CreateAndOutputSignalsSheet finalSignals, signalCount
    Else
        msgbox "No trading signals generated by GenWe."
    End If
    
    Application.EnableEvents = True
    Application.CALCULATION = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    'msgbox "GenerateWeeklyTradingStrategyAndSignals " & signalCount & " signals in " & Format(Timer - startTime, "0.000") & " seconds"
End Sub

' ===== HELPER FUNCTIONS =====
Function HasCompleteWeeklyData_Array(allData As Variant, rowNum As Long) As Boolean
    ' One-liner version - less error prone
    On Error GoTo ErrorHandler
    HasCompleteWeeklyData_Array = (allData(rowNum, 1) <> "" And allData(rowNum, 5) <> "" And _
                                  allData(rowNum, 7) <> "" And allData(rowNum, 10) <> "" And _
                                  allData(rowNum, 11) <> "" And allData(rowNum, 12) <> "" And _
                                  allData(rowNum, 14) <> "")
    Exit Function
ErrorHandler:
    HasCompleteWeeklyData_Array = False
End Function

Function GenerateWeeklySignal_Array(allData As Variant, rowNum As Long) As String
    ' Generate signal using array data only
    Dim rsi As Double, macd As Double, macdSignal As Double
    Dim priceVsMA As Double, compositeScore As Double, volumeSpike As Double
    Dim atrPercent As Double, ibs As Double
    
    ibs = ssNz(allData(rowNum, 8), 50)
    compositeScore = ssNz(allData(rowNum, 9), 0)
    rsi = ssNz(allData(rowNum, 10), 50)
    macd = ssNz(allData(rowNum, 11), 0)
    macdSignal = ssNz(allData(rowNum, 12), 0)
    priceVsMA = ssNz(allData(rowNum, 13), 0)
    atrPercent = ssNz(allData(rowNum, 15), 0)
    volumeSpike = ssNz(allData(rowNum, 16), 1)
    
    Dim score As Integer
    score = CalculateSignalScore_Fast(rsi, macd, macdSignal, priceVsMA, compositeScore, volumeSpike, atrPercent, ibs)
    
    If score >= 4 Then
        GenerateWeeklySignal_Array = "STRONG BUY"
    ElseIf score >= 2 Then
        GenerateWeeklySignal_Array = "BUY"
    ElseIf score <= -4 Then
        GenerateWeeklySignal_Array = "STRONG SELL"
    ElseIf score <= -2 Then
        GenerateWeeklySignal_Array = "SELL"
    Else
        GenerateWeeklySignal_Array = "HOLD"
    End If
End Function

Function CalculateSignalScore_Fast(rsi As Double, macd As Double, macdSignal As Double, _
                            priceVsMA As Double, compositeScore As Double, _
                            volumeSpike As Double, atrPercent As Double, ibs As Double) As Integer
    Dim score As Integer
    score = 0
    
    ' RSI scoring
    If rsi < 35 Then score = score + 2
    If rsi < 45 Then score = score + 1
    If rsi > 65 Then score = score - 2
    If rsi > 55 Then score = score - 1
    
    ' MACD scoring
    If macd > macdSignal Then
        If macd > 0 Then score = score + 2 Else score = score + 1
    Else
        If macd < 0 Then score = score - 2 Else score = score - 1
    End If
    
    ' Other indicators
    If priceVsMA > 2 Then score = score + 1
    If priceVsMA < -2 Then score = score - 1
    If compositeScore > 1 Then score = score + 1
    If compositeScore < -1 Then score = score - 1
    If volumeSpike > 1.2 Then
        If priceVsMA > 0 Then score = score + 1 Else score = score - 1
    End If
    If ibs < 30 Then score = score + 1
    If ibs > 70 Then score = score - 1
    If atrPercent > 8 Then score = score * 0.5
    
    CalculateSignalScore_Fast = score
End Function

Sub PopulateSignalArray_Array(signals() As Variant, signalCount As Long, allData As Variant, rowNum As Long, signal As String)
    ' Populate output array using array data only
    Dim entryPrice As Double, atr As Double
    entryPrice = allData(rowNum, 5)
    atr = allData(rowNum, 14)
    
    Dim stopLoss As Double, target As Double, positionSize As Double
    Dim riskPercent As Double, rewardRisk As Double, confidence As Integer
    Dim weeklyRange As Double, volatility As Double
    
    weeklyRange = ((allData(rowNum, 3) - allData(rowNum, 4)) / allData(rowNum, 4)) * 100
    CalculateWeeklyRiskManagement signal, entryPrice, atr, weeklyRange, stopLoss, target, positionSize, riskPercent, rewardRisk
    confidence = CalculateSignalConfidence_Array(allData, rowNum, signal)
    volatility = allData(rowNum, 15)
    
    signals(signalCount, 1) = allData(rowNum, 7)  ' Ticker
    signals(signalCount, 2) = allData(rowNum, 1)  ' Date
    signals(signalCount, 3) = signal
    signals(signalCount, 4) = GetSignalStrength(signal)
    signals(signalCount, 5) = entryPrice
    signals(signalCount, 6) = stopLoss
    signals(signalCount, 7) = target
    signals(signalCount, 8) = positionSize
    signals(signalCount, 9) = riskPercent
    signals(signalCount, 10) = rewardRisk
    signals(signalCount, 11) = allData(rowNum, 10) ' RSI
    signals(signalCount, 12) = allData(rowNum, 11) ' MACD
    signals(signalCount, 13) = allData(rowNum, 12) ' MACD Signal
    signals(signalCount, 14) = allData(rowNum, 13) ' Price vs MA
    signals(signalCount, 15) = allData(rowNum, 9)  ' Composite Score
    signals(signalCount, 16) = allData(rowNum, 16) ' Volume Spike
    signals(signalCount, 17) = allData(rowNum, 15) ' ATR %
    signals(signalCount, 18) = allData(rowNum, 8)  ' IBS
    signals(signalCount, 19) = weeklyRange
    signals(signalCount, 20) = volatility
    signals(signalCount, 21) = confidence
End Sub

Sub CalculateWeeklyRiskManagement(signal As String, entryPrice As Double, atr As Double, _
                                weeklyRange As Double, ByRef stopLoss As Double, _
                                ByRef target As Double, ByRef positionSize As Double, _
                                ByRef riskPercent As Double, ByRef rewardRisk As Double)
    Dim baseRisk As Double, atrMultiplier As Double
    
    ' Weekly trading parameters
    baseRisk = 0.02 ' 2% risk per trade
    atrMultiplier = 2 ' Use 2x ATR for stop loss
    
    If signal Like "*BUY*" Then
        ' BUY signal risk management
        stopLoss = entryPrice - (atr * atrMultiplier)
        stopLoss = Round(stopLoss, 2)
        target = entryPrice + (3 * atr) ' 3:1 reward/risk ratio
        target = Round(target, 2)
    ElseIf signal Like "*SELL*" Then
        ' SELL signal risk management
        stopLoss = entryPrice + (atr * atrMultiplier)
        stopLoss = Round(stopLoss, 2)
        target = entryPrice - (3 * atr) ' 3:1 reward/risk ratio
        target = Round(target, 2)
    Else
        stopLoss = 0
        target = 0
    End If
    
    ' Calculate position size based on risk
    Dim riskPerShare As Double
    riskPerShare = Abs(entryPrice - stopLoss)
    
    If riskPerShare > 0 Then
        positionSize = baseRisk / (riskPerShare / entryPrice)
        positionSize = Round(positionSize, 0)
    Else
        positionSize = 0
    End If
    
    riskPercent = baseRisk * 100
    rewardRisk = 3 ' Fixed 3:1 ratio for weekly trades
End Sub

Function CalculateSignalConfidence_Array(allData As Variant, rowNum As Long, signal As String) As Integer
    Dim confidence As Integer
    confidence = 3
    
    Dim rsi As Double, macd As Double, macdSignal As Double
    Dim volumeSpike As Double, compositeScore As Double
    
    rsi = allData(rowNum, 10)
    macd = allData(rowNum, 11)
    macdSignal = allData(rowNum, 12)
    volumeSpike = allData(rowNum, 16)
    compositeScore = allData(rowNum, 9)
    
    If signal = "STRONG BUY" Or signal = "STRONG SELL" Then confidence = confidence + 1
    If volumeSpike > 1.5 Then confidence = confidence + 1
    If (rsi < 35 And macd > macdSignal) Or (rsi > 65 And macd < macdSignal) Then confidence = confidence + 1
    
    If confidence > 5 Then confidence = 5
    If confidence < 1 Then confidence = 1
    
    CalculateSignalConfidence_Array = confidence
End Function

Function GetSignalStrength(signal As String) As String
    Select Case signal
        Case "STRONG BUY", "STRONG SELL"
            GetSignalStrength = "STRONG"
        Case "BUY", "SELL"
            GetSignalStrength = "MODERATE"
        Case Else
            GetSignalStrength = "WEAK"
    End Select
End Function

Function ssNz(value As Variant, Optional defaultVal As Variant = 0) As Variant
    If IsEmpty(value) Or value = "" Or IsNull(value) Then
        ssNz = defaultVal
    Else
        ssNz = value
    End If
End Function

Sub PopulateSignalArray_Single(signalData() As Variant, allData As Variant, rowNum As Long, signal As String)
    ' Populate a single signal data array
    Dim entryPrice As Double, atr As Double
    entryPrice = allData(rowNum, 5)
    atr = allData(rowNum, 14)
    
    Dim stopLoss As Double, target As Double, positionSize As Double
    Dim riskPercent As Double, rewardRisk As Double, confidence As Integer
    Dim weeklyRange As Double, volatility As Double
    
    weeklyRange = ((allData(rowNum, 3) - allData(rowNum, 4)) / allData(rowNum, 4)) * 100
    CalculateWeeklyRiskManagement signal, entryPrice, atr, weeklyRange, stopLoss, target, positionSize, riskPercent, rewardRisk
    confidence = CalculateSignalConfidence_Array(allData, rowNum, signal)
    volatility = allData(rowNum, 15)
    
    signalData(1) = allData(rowNum, 7)  ' Ticker
    signalData(2) = allData(rowNum, 1)  ' Date
    signalData(3) = signal
    signalData(4) = GetSignalStrength(signal)
    signalData(5) = entryPrice
    signalData(6) = stopLoss
    signalData(7) = target
    signalData(8) = positionSize
    signalData(9) = riskPercent
    signalData(10) = rewardRisk
    signalData(11) = allData(rowNum, 10) ' RSI
    signalData(12) = allData(rowNum, 11) ' MACD
    signalData(13) = allData(rowNum, 12) ' MACD Signal
    signalData(14) = allData(rowNum, 13) ' Price vs MA
    signalData(15) = allData(rowNum, 9)  ' Composite Score
    signalData(16) = allData(rowNum, 16) ' Volume Spike
    signalData(17) = allData(rowNum, 15) ' ATR %
    signalData(18) = allData(rowNum, 8)  ' IBS
    signalData(19) = weeklyRange
    signalData(20) = volatility
    signalData(21) = confidence
End Sub

Sub CreateAndOutputSignalsSheet(signals() As Variant, signalCount As Long)
    Dim wsSignals As Worksheet
    
    ' Delete existing sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("WeeklySignals").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create new sheet
    Set wsSignals = Worksheets.Add
    wsSignals.Name = "WeeklySignals"
    
    ' Add headers
    Dim headers() As Variant
    headers = Array("Ticker", "Date", "Signal", "Strength", "Entry Price", _
        "Stop Loss", "Target", "Position Size", "Risk %", "Reward/Risk", "RSI", "MACD", _
        "MACD Signal", "Price vs MA", "Composite Score", "Volume Spike", "ATR %", "IBS", _
        "Weekly Range", "Volatility", "Confidence")
    
    ' Write headers
    wsSignals.Range("A1").Resize(1, 21).value = headers
    
    ' Write all signals in ONE bulk operation
    If signalCount > 0 Then
        wsSignals.Range("A2").Resize(signalCount, 21).value = signals
    End If
    
    ' Apply formatting
    Call FastFormatWeeklySignals(wsSignals, signalCount)
End Sub

' FastFormatWeeklySignals is defined in WeeklySignals.bas — do not duplicate here.
