Attribute VB_Name = "WeeklySignals"
Option Explicit

' ULTRA-FAST SIGNAL GENERATION VERSION - PRE-FETCH ALL DATA
Sub GenerateWeeklyTradingSignals_UltraFast()
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
    allData = ws.Range("A1:Q" & lastRow).value
    
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
                Call PopulateSignalArray_Array(signals, signalCount, allData, i, signal)
            End If
        End If
    Next i
    
    ' Output results
    If signalCount > 0 Then
        ReDim Preserve signals(1 To signalCount, 1 To 21)
        Call CreateAndOutputSignalsSheet(signals, signalCount)
    End If
    
    Application.EnableEvents = True
    Application.CALCULATION = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    msgbox "GenWe...UltraFAST: " & signalCount & " signals in " & Format(Timer - startTime, "0.000") & " seconds"
End Sub

Function GenerateWeeklySignal_Array(allData As Variant, rowNum As Long) As String
    ' Generate signal using array data only
    Dim rsi As Double, macd As Double, macdSignal As Double
    Dim priceVsMA As Double, compositeScore As Double, volumeSpike As Double
    Dim atrPercent As Double, ibs As Double
    
    ibs = Nz(allData(rowNum, 8), 50)
    compositeScore = Nz(allData(rowNum, 9), 0)
    rsi = Nz(allData(rowNum, 10), 50)
    macd = Nz(allData(rowNum, 11), 0)
    macdSignal = Nz(allData(rowNum, 12), 0)
    priceVsMA = Nz(allData(rowNum, 13), 0)
    atrPercent = Nz(allData(rowNum, 15), 0)
    volumeSpike = Nz(allData(rowNum, 16), 1)
    
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
    Call CalculateWeeklyRiskManagement(signal, entryPrice, atr, weeklyRange, stopLoss, target, positionSize, riskPercent, rewardRisk)
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


' CORRECTED VERSION - Option 1 (Recommended)
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

Function origGenerateWeeklySignal_Array(allData As Variant, rowNum As Long) As String
    ' Generate signal using array data only (no sheet access)
    Dim rsi As Double, macd As Double, macdSignal As Double
    Dim priceVsMA As Double, compositeScore As Double, volumeSpike As Double
    Dim atrPercent As Double, ibs As Double
    
    ibs = ulwkNz(allData(rowNum, 8), 50)
    compositeScore = ulwkNz(allData(rowNum, 9), 0)
    rsi = ulwkNz(allData(rowNum, 10), 50)
    macd = ulwkNz(allData(rowNum, 11), 0)
    macdSignal = ulwkNz(allData(rowNum, 12), 0)
    priceVsMA = ulwkNz(allData(rowNum, 13), 0)
    atrPercent = ulwkNz(allData(rowNum, 15), 0)
    volumeSpike = ulwkNz(allData(rowNum, 16), 1)
    
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

Sub origPopulateSignalArray_Array(signals() As Variant, signalCount As Long, allData As Variant, rowNum As Long, signal As String)
    ' Populate output array using array data only
    Dim entryPrice As Double, atr As Double
    entryPrice = allData(rowNum, 5)
    atr = allData(rowNum, 14)
    
    Dim stopLoss As Double, target As Double, positionSize As Double
    Dim riskPercent As Double, rewardRisk As Double, confidence As Integer
    Dim weeklyRange As Double, volatility As Double
    
    weeklyRange = ((allData(rowNum, 3) - allData(rowNum, 4)) / allData(rowNum, 4)) * 100
    Call CalculateWeeklyRiskManagement(signal, entryPrice, atr, weeklyRange, stopLoss, target, positionSize, riskPercent, rewardRisk)
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

Function origCalculateSignalConfidence_Array(allData As Variant, rowNum As Long, signal As String) As Integer
    Dim confidence As Integer: confidence = 3
    
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

' HIGH-SPEED SIGNAL GENERATION WITH BULK OUTPUT
Sub GenerateWeeklyTradingSignals_Optimized()
    Dim ws As Worksheet, wsSignals As Worksheet
    Dim lastRow As Long, i As Long, signalCount As Long
    Dim startTime As Double
    startTime = Timer
    
    Application.ScreenUpdating = False
    Application.CALCULATION = xlCalculationManual
    Application.EnableEvents = False
    
    Set ws = Worksheets("Data")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Create output array in memory
    Dim signals() As Variant
    ReDim signals(1 To 10000, 1 To 21) ' Pre-allocate large array
    
    signalCount = 0
    
    ' Process data and populate array
    For i = 30 To lastRow
        If HasCompleteWeeklyData(ws, i) Then
            Dim ticker As String, currentDate As Date, entryPrice As Double
            Dim signal As String, atr As Double, atrPercent As Double
            Dim weeklyRange As Double, volatility As Double
            
            ticker = ws.Cells(i, 7).value
            currentDate = ws.Cells(i, 1).value
            entryPrice = ws.Cells(i, 5).value
            atr = ws.Cells(i, 14).value
            atrPercent = ws.Cells(i, 15).value
            
            ' Generate trading signal
            signal = GenerateWeeklySignal(ws, i)
            
            ' Only process BUY/SELL signals
            If signal <> "HOLD" Then
                signalCount = signalCount + 1
                
                Dim stopLoss As Double, target As Double, positionSize As Double
                Dim riskPercent As Double, rewardRisk As Double, confidence As Integer
                
                ' Calculate risk management
                Call CalculateWeeklyRiskManagement(signal, entryPrice, atr, weeklyRange, _
                                                  stopLoss, target, positionSize, riskPercent, rewardRisk)
                
                confidence = CalculateSignalConfidence(ws, i, signal)
                weeklyRange = CalculateWeeklyRange(ws, i)
                volatility = CalculateVolatility(ws, i)
                
                ' Populate array (much faster than writing to cells)
                signals(signalCount, 1) = ticker
                signals(signalCount, 2) = currentDate
                signals(signalCount, 3) = signal
                signals(signalCount, 4) = GetSignalStrength(signal)
                signals(signalCount, 5) = entryPrice
                signals(signalCount, 6) = stopLoss
                signals(signalCount, 7) = target
                signals(signalCount, 8) = positionSize
                signals(signalCount, 9) = riskPercent
                signals(signalCount, 10) = rewardRisk
                signals(signalCount, 11) = ws.Cells(i, 10).value
                signals(signalCount, 12) = ws.Cells(i, 11).value
                signals(signalCount, 13) = ws.Cells(i, 12).value
                signals(signalCount, 14) = ws.Cells(i, 13).value
                signals(signalCount, 15) = ws.Cells(i, 9).value
                signals(signalCount, 16) = ws.Cells(i, 16).value
                signals(signalCount, 17) = atrPercent
                signals(signalCount, 18) = ws.Cells(i, 8).value
                signals(signalCount, 19) = weeklyRange
                signals(signalCount, 20) = volatility
                signals(signalCount, 21) = confidence
            End If
        End If
    Next i
    
    ' Create and output to signals sheet in ONE operation
    If signalCount > 0 Then
        Call CreateAndOutputSignalsSheet(signals, signalCount)
    End If
    
    Application.EnableEvents = True
    Application.CALCULATION = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    msgbox "Weekly trading signals generated: " & signalCount & vbCrLf & _
           "Processing time: " & Format(Timer - startTime, "0.00") & " seconds"
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
    
    ' Write headers in one operation
    wsSignals.Range("A1").Resize(1, 21).value = headers
    
    ' Write all signals in ONE bulk operation (fastest method)
    If signalCount > 0 Then
        wsSignals.Range("A2").Resize(signalCount, 21).value = signals
    End If
    
    ' Apply formatting
    Call FastFormatWeeklySignals(wsSignals, signalCount)
End Sub

Sub FastFormatWeeklySignals(ws As Worksheet, signalCount As Long)
    If signalCount < 1 Then Exit Sub
    
    With ws
        ' Auto-fit columns first (fastest when done once)
        .Columns.AutoFit
        
        ' Apply number formats in bulk ranges
        With .Range("E2:E" & signalCount + 1) ' Entry Price
            .NumberFormat = "0.00"
        End With
        With .Range("F2:G" & signalCount + 1) ' Stop Loss & Target
            .NumberFormat = "0.00"
        End With
        With .Range("K2:K" & signalCount + 1) ' RSI
            .NumberFormat = "0.0"
        End With
        With .Range("L2:M" & signalCount + 1) ' MACD
            .NumberFormat = "0.0000"
        End With
        With .Range("N2:O" & signalCount + 1) ' Price vs MA & Composite
            .NumberFormat = "0.00"
        End With
        With .Range("P2:T" & signalCount + 1) ' Volume, ATR%, IBS, Range, Vol
            .NumberFormat = "0.00"
        End With
        
        ' Fast conditional formatting for signals
        Call ApplyFastConditionalFormatting(ws, signalCount)
        
        ' Add borders to entire range at once
        With .Range("A1:U" & signalCount + 1).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        ' AutoFilter and freeze panes
        .Range("A1:U1").AutoFilter
        .Rows(2).Select
        ActiveWindow.FreezePanes = True
        .Range("A1").Select
    End With
End Sub

Sub ApplyFastConditionalFormatting(ws As Worksheet, signalCount As Long)
    If signalCount < 1 Then Exit Sub
    
    ' Remove any existing conditional formatting
    ws.Cells.FormatConditions.Delete
    
    ' Add conditional formatting for BUY signals
    With ws.Range("C2:C" & signalCount + 1).FormatConditions.Add(Type:=xlTextString, String:="*BUY*", TextOperator:=xlContains)
        .Interior.Color = RGB(146, 208, 80)
        .StopIfTrue = False
    End With
    
    ' Add conditional formatting for STRONG BUY signals
    With ws.Range("C2:C" & signalCount + 1).FormatConditions.Add(Type:=xlTextString, String:="STRONG BUY", TextOperator:=xlContains)
        .Interior.Color = RGB(0, 176, 80)
        .Font.Color = RGB(255, 255, 255)
        .StopIfTrue = False
    End With
    
    ' Add conditional formatting for SELL signals
    With ws.Range("C2:C" & signalCount + 1).FormatConditions.Add(Type:=xlTextString, String:="*SELL*", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 102, 102)
        .StopIfTrue = False
    End With
    
    ' Add conditional formatting for STRONG SELL signals
    With ws.Range("C2:C" & signalCount + 1).FormatConditions.Add(Type:=xlTextString, String:="STRONG SELL", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .StopIfTrue = False
    End With
End Sub
' UPDATED ORIGINAL FUNCTION
' OPTIMIZED VERSION OF HASCOMPLETEWEEKLydata

Function HasCompleteWeeklyData(ws As Worksheet, rowNum As Long) As Boolean
    On Error GoTo ErrorHandler
    
    ' Check each required column individually
    If ws.Cells(rowNum, 1).value = "" Then Exit Function  ' Date
    If ws.Cells(rowNum, 5).value = "" Then Exit Function  ' Close Price
    If ws.Cells(rowNum, 7).value = "" Then Exit Function  ' Ticker
    If ws.Cells(rowNum, 10).value = "" Then Exit Function ' RSI
    If ws.Cells(rowNum, 11).value = "" Then Exit Function ' MACD
    If ws.Cells(rowNum, 12).value = "" Then Exit Function ' MACD Signal
    If ws.Cells(rowNum, 14).value = "" Then Exit Function ' ATR
    
    HasCompleteWeeklyData = True
    Exit Function
    
ErrorHandler:
    HasCompleteWeeklyData = False
End Function

' OPTIMIZED SIGNAL GENERATION
Function GenerateWeeklySignal_Fast(ws As Worksheet, currentRow As Long) As String
    ' Get all indicator values in one operation
    Dim indicators As Variant
    indicators = ws.Range(ws.Cells(currentRow, 8), ws.Cells(currentRow, 16)).value
    
    Dim rsi As Double, macd As Double, macdSignal As Double
    Dim priceVsMA As Double, compositeScore As Double, volumeSpike As Double
    Dim atrPercent As Double, ibs As Double
    
    ibs = ulwkNz(indicators(1, 1), 50)
    compositeScore = ulwkNz(indicators(1, 2), 0)
    rsi = ulwkNz(indicators(1, 3), 50)
    macd = ulwkNz(indicators(1, 4), 0)
    macdSignal = ulwkNz(indicators(1, 5), 0)
    priceVsMA = ulwkNz(indicators(1, 6), 0)
    atrPercent = ulwkNz(indicators(1, 7), 0)
    volumeSpike = ulwkNz(indicators(1, 9), 1)
    
    Dim score As Integer
    score = CalculateSignalScore_Fast(rsi, macd, macdSignal, priceVsMA, compositeScore, volumeSpike, atrPercent, ibs)
    
    If score >= 4 Then
        GenerateWeeklySignal_Fast = "STRONG BUY"
    ElseIf score >= 2 Then
        GenerateWeeklySignal_Fast = "BUY"
    ElseIf score <= -4 Then
        GenerateWeeklySignal_Fast = "STRONG SELL"
    ElseIf score <= -2 Then
        GenerateWeeklySignal_Fast = "SELL"
    Else
        GenerateWeeklySignal_Fast = "HOLD"
    End If
End Function

Function origCalculateSignalScore_Fast(rsi As Double, macd As Double, macdSignal As Double, _
                            priceVsMA As Double, compositeScore As Double, _
                            volumeSpike As Double, atrPercent As Double, ibs As Double) As Integer
    Dim score As Integer
    score = 0
    
    ' Fast RSI scoring
    If rsi < 35 Then score = score + 2
    If rsi < 45 Then score = score + 1
    If rsi > 65 Then score = score - 2
    If rsi > 55 Then score = score - 1
    
    ' Fast MACD scoring
    If macd > macdSignal Then
        If macd > 0 Then score = score + 2 Else score = score + 1
    Else
        If macd < 0 Then score = score - 2 Else score = score - 1
    End If
    
    ' Fast other indicators
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
Function ulwkNz(value As Variant, Optional defaultVal As Variant = 0) As Variant
    If IsEmpty(value) Or value = "" Or IsNull(value) Then
        ulwkNz = defaultVal
    Else
        ulwkNz = value
    End If
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

' ===== MISSING FUNCTION WRAPPERS =====
' GenerateWeeklySignal: alias for GenerateWeeklySignal_Fast (same signature)
Function GenerateWeeklySignal(ws As Worksheet, rowNum As Long) As String
    GenerateWeeklySignal = GenerateWeeklySignal_Fast(ws, rowNum)
End Function

' CalculateSignalConfidence: returns 1-5 confidence score based on indicator alignment
Function CalculateSignalConfidence(ws As Worksheet, rowNum As Long, signal As String) As Integer
    Dim score As Integer: score = 0
    Dim rsi As Double, macd As Double, macdSig As Double
    Dim priceVsMA As Double, compScore As Double, vol As Double, atrPct As Double
    On Error Resume Next
    rsi       = CDbl(ws.Cells(rowNum, 10).Value)
    macd      = CDbl(ws.Cells(rowNum, 11).Value)
    macdSig   = CDbl(ws.Cells(rowNum, 12).Value)
    priceVsMA = CDbl(ws.Cells(rowNum, 13).Value)
    compScore = CDbl(ws.Cells(rowNum, 9).Value)
    vol       = CDbl(ws.Cells(rowNum, 16).Value)
    atrPct    = CDbl(ws.Cells(rowNum, 15).Value)
    On Error GoTo 0

    If InStr(signal, "BUY") > 0 Then
        If rsi < 45 Then score = score + 1
        If macd > macdSig Then score = score + 1
        If priceVsMA > 0 Then score = score + 1
        If compScore > 0 Then score = score + 1
        If vol > 1.2 Then score = score + 1
    Else
        If rsi > 55 Then score = score + 1
        If macd < macdSig Then score = score + 1
        If priceVsMA < 0 Then score = score + 1
        If compScore < 0 Then score = score + 1
        If vol > 1.2 Then score = score + 1
    End If
    If atrPct > 8 Then score = score - 1  ' Penalise extreme volatility
    CalculateSignalConfidence = WorksheetFunction.Max(1, WorksheetFunction.Min(5, score))
End Function

' CalculateWeeklyRange: high-low range as % of close over last 5 rows for same ticker
Function CalculateWeeklyRange(ws As Worksheet, rowNum As Long) As Double
    On Error GoTo Fail
    Dim ticker As String: ticker = CStr(ws.Cells(rowNum, 7).Value)
    Dim closePrice As Double: closePrice = CDbl(ws.Cells(rowNum, 5).Value)
    If closePrice = 0 Then GoTo Fail

    Dim hi As Double, lo As Double, found As Long
    hi = 0: lo = 1E+20: found = 0
    Dim i As Long
    For i = rowNum To WorksheetFunction.Max(2, rowNum - 20) Step -1
        If CStr(ws.Cells(i, 7).Value) = ticker Then
            Dim h As Double, l As Double
            h = CDbl(ws.Cells(i, 3).Value)  ' High col
            l = CDbl(ws.Cells(i, 4).Value)  ' Low col
            If h > hi Then hi = h
            If l < lo Then lo = l
            found = found + 1
            If found >= 5 Then Exit For
        End If
    Next i
    If hi > 0 And lo > 0 Then
        CalculateWeeklyRange = (hi - lo) / closePrice * 100
    End If
    Exit Function
Fail:
    CalculateWeeklyRange = 0
End Function

' CalculateVolatility: ATR% as a volatility proxy (already calculated in col 15)
Function CalculateVolatility(ws As Worksheet, rowNum As Long) As Double
    On Error Resume Next
    CalculateVolatility = CDbl(ws.Cells(rowNum, 15).Value)  ' ATR% column
    On Error GoTo 0
End Function


