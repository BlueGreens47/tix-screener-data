Attribute VB_Name = "CompleteSignalsGeneration"
Option Explicit

'******** 5. TRADING SIGNAL GENERATION ********
Sub GenerateCompleteTradingSignals_Main()
    Dim wsData As Worksheet, wsDash As Worksheet, wsSignals As Worksheet
    Dim lastRow As Long, signalCount As Long
    
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsDash = ThisWorkbook.Sheets("DashBoard")
    Set wsSignals = GetOrCreateSheet("mainTradingSignals")
    
    Call SetupTradingSignalsSheet(wsSignals)
    
    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).row
    If lastRow < 50 Then
        Debug.Print "Insufficient data for signal generation"
        Exit Sub
    End If
    
    Dim tickers() As String
    tickers = GetUniqueTickersUniversal(wsData, lastRow)
    
    Dim tickerCount As Long
    If IsArrayEmpty(tickers) Then
        tickerCount = 0
    Else
        tickerCount = UBound(tickers) - LBound(tickers) + 1
    End If
    
    If tickerCount > 0 Then
        signalCount = ProcessTradingSignalsArray(wsData, wsDash, wsSignals, tickers, tickerCount, lastRow)
        'msgbox "Successfully generated " & signalCount & " trading signals", vbInformation
    Else
        msgbox "No qualifying tickers found for signal generation", vbExclamation
    End If
End Sub

Function GetUniqueTickersUniversal(ws As Worksheet, lastRow As Long) As String()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To lastRow
        On Error Resume Next
        Dim cellValue As Variant
        cellValue = ws.Cells(i, 7).value
        
        If Not IsError(cellValue) And Not IsEmpty(cellValue) And cellValue <> "" Then
            Dim ticker As String
            ticker = Trim(CStr(cellValue))
            If ticker <> "" And Not dict.Exists(ticker) Then
                dict.Add ticker, 1
            End If
        End If
        On Error GoTo 0
    Next i
    
    If dict.count > 0 Then
        Dim result() As String
        ReDim result(1 To dict.count)
        Dim key As Variant, idx As Long
        idx = 1
        For Each key In dict.Keys
            result(idx) = CStr(key)
            idx = idx + 1
        Next key
        GetUniqueTickersUniversal = result
    Else
        GetUniqueTickersUniversal = Split("", ",")
    End If
End Function

Function IsArrayEmpty(arr As Variant) As Boolean
    On Error GoTo ErrorHandler
    If Not IsArray(arr) Then
        IsArrayEmpty = True
        Exit Function
    End If
    
    Dim test As Long
    test = UBound(arr)
    IsArrayEmpty = (UBound(arr) < LBound(arr))
    Exit Function
    
ErrorHandler:
    IsArrayEmpty = True
End Function

' Helper function to check if row has sufficient data
Function HasSufficientData(ws As Worksheet, rowNum As Long) As Boolean
    ' Check if critical data columns have values
    If IsEmpty(ws.Cells(rowNum, 9).value) Or IsEmpty(ws.Cells(rowNum, 10).value) Or _
       IsEmpty(ws.Cells(rowNum, 11).value) Or IsEmpty(ws.Cells(rowNum, 12).value) Then
        HasSufficientData = False
    Else
        HasSufficientData = True
    End If
End Function
'*** 5.TRADE SIGNAL GENERATION END***

' ARRAY-BASED SIGNAL PROCESSING
Function ProcessTradingSignalsArray(wsData As Worksheet, wsDash As Worksheet, wsSignals As Worksheet, tickers() As String, tickerCount As Long, lastRow As Long) As Long
    Dim i As Long, signalCount As Long
    Dim signalData() As Variant
    Dim maxSignals As Long
    
    maxSignals = tickerCount
    ReDim signalData(1 To maxSignals, 1 To 17)
    signalCount = 0
    
    ' Process each ticker
    For i = 1 To tickerCount
        Dim ticker As String
        ticker = tickers(i)
        
        ' Get signal data for this ticker
        Dim signalResult As Variant
        signalResult = GetTickerSignalData(wsData, wsDash, ticker, lastRow)
        
        ' If we have a valid signal (not HOLD), add to results
        If IsArray(signalResult) Then
            signalCount = signalCount + 1
            Dim j As Long
            For j = 1 To 17
                signalData(signalCount, j) = signalResult(j - 1)
            Next j
        End If
    Next i
    
    ' Output signals if we have any
    If signalCount > 0 Then
        Call OutputTradingSignalsArray(wsSignals, signalData, signalCount)
        Call ApplyTradingSignalsFormattingArray(wsSignals, signalCount)
        Call GenerateTradingSummaryArray(wsSignals, signalData, signalCount)
    End If
    
    ProcessTradingSignalsArray = signalCount
End Function

' GET TICKER SIGNAL DATA - WITH EXACT COLUMN MAPPING
Function modGetTickerSignalData(ws As Worksheet, wsDash As Worksheet, ticker As String, lastRow As Long) As Variant
    Dim i As Long
    Dim signalArray(0 To 16) As Variant
    
    ' Initialize with empty values
    For i = 0 To 16
        signalArray(i) = ""
    Next i
    
    ' Find the most recent COMPLETE data for this ticker
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 7).value = ticker Then
            ' Skip rows without sufficient data
            If Not HasSufficientData(ws, i) Then
                GoTo ContinueLoop
            End If
            
            ' ... [rest of your existing calculation code remains the same]
            ' Get indicator values
            Dim entryPrice As Double, compositeScore As Double, rsi As Double
            Dim macd As Double, macdSignal As Double, priceVsMA As Double
            Dim atr As Double, atrPercent As Double, volumeSpike As Double
            
            entryPrice = Nz(ws.Cells(i, 5).value, 0)
            compositeScore = Nz(ws.Cells(i, 9).value, 0)
            rsi = Nz(ws.Cells(i, 10).value, 0)
            macd = Nz(ws.Cells(i, 11).value, 0)
            macdSignal = Nz(ws.Cells(i, 12).value, 0)
            priceVsMA = Nz(ws.Cells(i, 13).value, 0)
            atr = Nz(ws.Cells(i, 14).value, 0)
            atrPercent = Nz(ws.Cells(i, 15).value, 0)
            volumeSpike = Nz(ws.Cells(i, 13).value, 1)
            
            ' Calculate buy/sell scores
            Dim buyScore As Double, sellScore As Double
            buyScore = CalculateBuyScoreSimple(rsi, macd, macdSignal, priceVsMA, compositeScore, volumeSpike, atrPercent)
            sellScore = CalculateSellScoreSimple(rsi, macd, macdSignal, priceVsMA, compositeScore, volumeSpike, atrPercent)
            
            ' Get threshold parameters from dashboard
            Dim minScore As Double, maxScore As Double
            minScore = Nz(wsDash.Range("W5").value, 2)
            maxScore = Nz(wsDash.Range("W6").value, -2)
            
            ' DETERMINE SIGNAL - WITH CORRECTED THRESHOLD LOGIC
            Dim signalType As String, signalStrength As String
            Dim stopLoss As Double, positionSize As Double, riskPerShare As Double, riskRewardRatio As Double
            
            ' FIXED: Compare buyScore against absolute value of sellScore
            If buyScore >= minScore And buyScore > Abs(sellScore) Then
                signalType = "BUY"
                signalStrength = GetSignalStrengthSimple(buyScore, minScore)
                Call CalculateBuyRiskManagementSimple(entryPrice, atr, atrPercent, stopLoss, positionSize, riskPerShare, riskRewardRatio)
            ' FIXED: Compare absolute sellScore against buyScore
            ElseIf Abs(sellScore) >= Abs(maxScore) And Abs(sellScore) > buyScore Then
                signalType = "SELL"
                signalStrength = GetSignalStrengthSimple(Abs(sellScore), Abs(maxScore))
                Call CalculateSellRiskManagementSimple(entryPrice, atr, atrPercent, stopLoss, positionSize, riskPerShare, riskRewardRatio)
            Else
                ' HOLD signal - skip this row and continue
                GoTo ContinueLoop
            End If
            
            ' Populate signal array
            signalArray(0) = ticker
            signalArray(1) = signalType
            signalArray(2) = signalStrength
            signalArray(3) = Round(entryPrice, 2)
            signalArray(4) = Round(stopLoss, 2)
            signalArray(5) = positionSize
            signalArray(6) = Round(riskPerShare, 2)
            signalArray(7) = Round(riskRewardRatio, 2)
            signalArray(8) = Round(compositeScore, 2)
            signalArray(9) = Round(rsi, 1)
            signalArray(10) = Round(macd, 4)
            signalArray(11) = Round(macdSignal, 4)
            signalArray(12) = Round(priceVsMA, 2)
            signalArray(13) = Round(atr, 4)
            signalArray(14) = Round(atrPercent, 2)
            signalArray(15) = Round(volumeSpike, 2)
            signalArray(16) = Date
            
            GetTickerSignalData = signalArray
            Exit Function
        End If
        
ContinueLoop:
    Next i
    
    ' No signal found
    GetTickerSignalData = False
End Function

' GET TICKER SIGNAL DATA - WITH EXACT COLUMN MAPPING
Private Function GetTickerSignalData(ws As Worksheet, wsDash As Worksheet, ticker As String, lastRow As Long) As Variant  ' Private: canonical version in CompleteTrading.bas
    Dim i As Long
    Dim signalArray(0 To 16) As Variant
    Dim dataCheck As Boolean
    
    ' Initialize with empty values
    For i = 0 To 16
        signalArray(i) = ""
    Next i
    
    ' Find the most recent data for this ticker
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 7).value = ticker Then  ' Column G: Ticker
            ' Check if this row has sufficient data for calculations
            dataCheck = HasSufficientData(ws, i)
            
            If Not dataCheck Then
                ' Skip this row and continue to next iteration
                GoTo ContinueLoop
            End If
            
            ' Get indicator values - WITH EXACT COLUMN MAPPING
            Dim entryPrice As Double, compositeScore As Double, rsi As Double
            Dim macd As Double, macdSignal As Double, priceVsMA As Double
            Dim atr As Double, atrPercent As Double, volumeSpike As Double
            
            ' EXACT COLUMN MAPPING BASED ON YOUR HEADERS:
            entryPrice = Nz(ws.Cells(i, 5).value, 0)        ' E: Close
           
            rsi = Nz(ws.Cells(i, 8).value, 0)               ' H: RSI
            macd = Nz(ws.Cells(i, 9).value, 0)              ' I: MACD
            macdSignal = Nz(ws.Cells(i, 10).value, 0)       ' J: MACD_Signal
            priceVsMA = Nz(ws.Cells(i, 11).value, 0)        ' K: PriceVsMA50
            volumeSpike = Nz(ws.Cells(i, 13).value, 1)      ' M: Volume_Spike
            compositeScore = Nz(ws.Cells(i, 14).value, 0)   ' N: Composite_Score
            atr = Nz(ws.Cells(i, 16).value, 0)              ' P: ATR
            atrPercent = Nz(ws.Cells(i, 17).value, 0)       ' Q: ATR %
            
            
            ' Calculate buy/sell scores
            Dim buyScore As Double, sellScore As Double
            buyScore = CalculateBuyScoreSimple(rsi, macd, macdSignal, priceVsMA, compositeScore, volumeSpike, atrPercent)
            sellScore = CalculateSellScoreSimple(rsi, macd, macdSignal, priceVsMA, compositeScore, volumeSpike, atrPercent)
            
            ' Get threshold parameters from dashboard
            Dim minScore As Double, maxScore As Double
            minScore = Nz(wsDash.Range("Y5").value, 2)
            maxScore = Nz(wsDash.Range("Y6").value, -2)
            
            ' DETERMINE SIGNAL - WITH CORRECTED THRESHOLD LOGIC
            Dim signalType As String, signalStrength As String
            Dim stopLoss As Double, positionSize As Double, riskPerShare As Double, riskRewardRatio As Double
            
            ' FIXED: Compare buyScore against absolute value of sellScore
            If buyScore >= minScore And buyScore > Abs(sellScore) Then
                signalType = "BUY"
                signalStrength = GetSignalStrengthSimple(buyScore, minScore)
                Call CalculateBuyRiskManagementSimple(entryPrice, atr, atrPercent, stopLoss, positionSize, riskPerShare, riskRewardRatio)
            ' FIXED: Compare absolute sellScore against buyScore
            ElseIf Abs(sellScore) >= Abs(maxScore) And Abs(sellScore) > buyScore Then
                signalType = "SELL"
                signalStrength = GetSignalStrengthSimple(Abs(sellScore), Abs(maxScore))
                Call CalculateSellRiskManagementSimple(entryPrice, atr, atrPercent, stopLoss, positionSize, riskPerShare, riskRewardRatio)
            Else
                ' HOLD signal - skip this row and continue
                GoTo ContinueLoop
            End If
            
            ' Populate signal array
            signalArray(0) = ticker
            signalArray(1) = signalType
            signalArray(2) = signalStrength
            signalArray(3) = Round(entryPrice, 2)
            signalArray(4) = Round(stopLoss, 2)
            signalArray(5) = positionSize
            signalArray(6) = Round(riskPerShare, 2)
            signalArray(7) = Round(riskRewardRatio, 2)
            signalArray(8) = Round(compositeScore, 2)
            signalArray(9) = Round(rsi, 1)
            signalArray(10) = Round(macd, 4)
            signalArray(11) = Round(macdSignal, 4)
            signalArray(12) = Round(priceVsMA, 2)
            signalArray(13) = Round(atr, 4)
            signalArray(14) = Round(atrPercent, 2)
            signalArray(15) = Round(volumeSpike, 2)
            signalArray(16) = wsDash.Range("H5").value
            
            GetTickerSignalData = signalArray
            Exit Function
        End If
        
ContinueLoop:
    Next i
    
    ' No signal found
    GetTickerSignalData = False
End Function

' CALCULATION SETTINGS
Function CalculateBuyScoreSimple(rsi As Double, macd As Double, macdSignal As Double, priceVsMA As Double, compositeScore As Double, volumeSpike As Double, atrPercent As Double) As Double
    Dim score As Double
    
    ' RSI - bullish when oversold
    If rsi < 30 Then
        score = score + ((30 - rsi) / 30) * 25
    ElseIf rsi < 45 Then
        score = score + ((45 - rsi) / 15) * 15
    End If
    
    ' MACD - bullish when above signal line
    If macd > macdSignal Then
        score = score + (Abs(macd - macdSignal) * 20)
    End If
    
    ' Price vs MA - bullish when above MA
    If priceVsMA > 0 Then
        score = score + (priceVsMA * 2)
    ElseIf priceVsMA > -2 Then
        score = score + 5
    End If
    
    ' Composite Score
    If compositeScore > 0 Then
        score = score + (compositeScore * 15)
    End If
    
    ' Volume confirmation
    If volumeSpike > 1.2 Then
        score = score + 10
    ElseIf volumeSpike > 1 Then
        score = score + 5
    End If
    
    ' ATR - lower volatility preferred for buys
    If atrPercent < 3 Then
        score = score + 5
    End If
    
    CalculateBuyScoreSimple = score
End Function

Function CalculateSellScoreSimple(rsi As Double, macd As Double, macdSignal As Double, priceVsMA As Double, compositeScore As Double, volumeSpike As Double, atrPercent As Double) As Double
    Dim score As Double
    
    ' RSI - bearish when overbought
    If rsi > 70 Then
        score = score - ((rsi - 70) / 30) * 25
    ElseIf rsi > 55 Then
        score = score - ((rsi - 55) / 15) * 15
    End If
    
    ' MACD - bearish when below signal line
    If macd < macdSignal Then
        score = score - (Abs(macd - macdSignal) * 20)
    End If
    
    ' Price vs MA - bearish when below MA
    If priceVsMA < 0 Then
        score = score + (priceVsMA * 2)
    ElseIf priceVsMA < 2 Then
        score = score - 5
    End If
    
    ' Composite Score
    If compositeScore < 0 Then
        score = score + (compositeScore * 15)
    End If
    
    ' Volume confirmation
    If volumeSpike > 1.2 Then
        score = score - 10
    ElseIf volumeSpike > 1 Then
        score = score - 5
    End If
    
    ' ATR - higher volatility adds to sell signal strength
    If atrPercent > 3 Then
        score = score - 5
    End If
    
    CalculateSellScoreSimple = score
End Function

Function GetSignalStrengthSimple(score As Double, threshold As Double) As String
    Dim strengthRatio As Double
    strengthRatio = score / threshold
    
    If strengthRatio >= 2 Then
        GetSignalStrengthSimple = "STRONG"
    ElseIf strengthRatio >= 1.5 Then
        GetSignalStrengthSimple = "MEDIUM"
    Else
        GetSignalStrengthSimple = "WEAK"
    End If
End Function


Sub CalculateBuyRiskManagementSimple(entryPrice As Double, atr As Double, atrPercent As Double, ByRef stopLoss As Double, ByRef positionSize As Double, ByRef riskPerShare As Double, ByRef riskRewardRatio As Double)
    stopLoss = entryPrice - (2 * atr)
    riskPerShare = entryPrice - stopLoss
    
    ' Position Size based on volatility
    If atrPercent < 1.5 Then
        positionSize = 8
    ElseIf atrPercent < 3 Then
        positionSize = 6
    ElseIf atrPercent < 5 Then
        positionSize = 4
    Else
        positionSize = 2
    End If
    
    ' Risk/Reward Ratio
    Dim targetPrice As Double
    targetPrice = entryPrice + (4 * atr)
    riskRewardRatio = (targetPrice - entryPrice) / riskPerShare
End Sub

Sub CalculateSellRiskManagementSimple(entryPrice As Double, atr As Double, atrPercent As Double, ByRef stopLoss As Double, ByRef positionSize As Double, ByRef riskPerShare As Double, ByRef riskRewardRatio As Double)
    stopLoss = entryPrice + (2 * atr)
    riskPerShare = stopLoss - entryPrice
    
    ' More conservative position sizing for short sales
    If atrPercent < 1.5 Then
        positionSize = 6
    ElseIf atrPercent < 3 Then
        positionSize = 4
    ElseIf atrPercent < 5 Then
        positionSize = 3
    Else
        positionSize = 1
    End If
    
    ' Risk/Reward Ratio for short sales
    Dim targetPrice As Double
    targetPrice = entryPrice - (4 * atr)
    riskRewardRatio = (entryPrice - targetPrice) / riskPerShare
End Sub

' OUTPUT AND FORMATTING FUNCTIONS
' OUTPUT AND FORMATTING FUNCTIONS - OPTIMIZED
Sub OutputTradingSignalsArray(ws As Worksheet, signalData() As Variant, signalCount As Long)
    Dim outputRange As Range
    Dim outputArray() As Variant
    Dim i As Long
    
    If signalCount = 0 Then Exit Sub
    
    ' Resize output array to match signal count
    ReDim outputArray(1 To signalCount, 1 To 17)
    
    ' Fill output array in memory (much faster than writing to cells)
    For i = 1 To signalCount
        outputArray(i, 1) = signalData(i, 1)   ' Ticker
        outputArray(i, 2) = signalData(i, 2)   ' Signal
        outputArray(i, 3) = signalData(i, 3)   ' Strength
        outputArray(i, 4) = signalData(i, 4)   ' Entry Price
        outputArray(i, 5) = signalData(i, 5)   ' Stop Loss
        outputArray(i, 6) = signalData(i, 6)   ' Position Size
        outputArray(i, 7) = signalData(i, 7)   ' Risk/Share
        outputArray(i, 8) = signalData(i, 8)   ' R/R Ratio
        outputArray(i, 9) = signalData(i, 9)   ' Composite Score
        outputArray(i, 10) = signalData(i, 10) ' RSI
        outputArray(i, 11) = signalData(i, 11) ' MACD
        outputArray(i, 12) = signalData(i, 12) ' MACD Signal
        outputArray(i, 13) = signalData(i, 13) ' Price vs MA
        outputArray(i, 14) = signalData(i, 14) ' ATR
        outputArray(i, 15) = signalData(i, 15) ' ATR %
        outputArray(i, 16) = signalData(i, 16) ' Volume Spike
        outputArray(i, 17) = signalData(i, 17) ' Timestamp
    Next i
    
    ' Write entire array to worksheet in one operation
    Set outputRange = ws.Range("A4").Resize(signalCount, 17)
    outputRange.value = outputArray
    
    ws.Columns.AutoFit
End Sub

Sub origOutputTradingSignalsArray(ws As Worksheet, signalData() As Variant, signalCount As Long)
    Dim i As Long, outputRow As Long
    outputRow = 4 ' Start after headers
    
    For i = 1 To signalCount
        ws.Cells(outputRow, 1).value = signalData(i, 1)  ' Ticker
        ws.Cells(outputRow, 2).value = signalData(i, 2)  ' Signal
        ws.Cells(outputRow, 3).value = signalData(i, 3)  ' Strength
        ws.Cells(outputRow, 4).value = signalData(i, 4)  ' Entry Price
        ws.Cells(outputRow, 5).value = signalData(i, 5)  ' Stop Loss
        ws.Cells(outputRow, 6).value = signalData(i, 6)  ' Position Size
        ws.Cells(outputRow, 7).value = signalData(i, 7)  ' Risk/Share
        ws.Cells(outputRow, 8).value = signalData(i, 8)  ' R/R Ratio
        ws.Cells(outputRow, 9).value = signalData(i, 9)  ' Composite Score
        ws.Cells(outputRow, 10).value = signalData(i, 10) ' RSI
        
        ws.Cells(outputRow, 11).value = signalData(i, 11) ' MACD
        ws.Cells(outputRow, 12).value = signalData(i, 12) ' MACD Signal
        ws.Cells(outputRow, 13).value = signalData(i, 13) ' Price vs MA
        ws.Cells(outputRow, 14).value = signalData(i, 14) ' ATR
        ws.Cells(outputRow, 15).value = signalData(i, 15) ' ATR %
        ws.Cells(outputRow, 16).value = signalData(i, 16) ' Volume Spike
        ws.Cells(outputRow, 17).value = signalData(i, 17) ' Timestamp
        
        outputRow = outputRow + 1
    Next i
    
    ws.Columns.AutoFit
End Sub

Sub ApplyTradingSignalsFormattingArray(ws As Worksheet, signalCount As Long)
    Dim i As Long
    
    For i = 4 To signalCount + 3
        ' Signal type coloring
        With ws.Cells(i, 2)
            If .value = "BUY" Then
                .Interior.Color = RGB(198, 239, 206)
                .Font.Color = RGB(0, 128, 0)
            ElseIf .value = "SELL" Then
                .Interior.Color = RGB(255, 199, 206)
                .Font.Color = RGB(255, 0, 0)
            End If
            .Font.Bold = True
        End With
        
        ' Signal strength formatting
        With ws.Cells(i, 3)
            If .value = "STRONG" Then
                .Font.Bold = True
                .Font.Color = RGB(0, 100, 0)
            ElseIf .value = "WEAK" Then
                .Font.Color = RGB(128, 128, 128)
            End If
        End With
        
        ' RSI coloring
        With ws.Cells(i, 10)
            If .value < 30 Then
                .Interior.Color = RGB(198, 239, 206)
            ElseIf .value > 70 Then
                .Interior.Color = RGB(255, 199, 206)
            End If
        End With
        
        ' Risk/Reward ratio coloring
        With ws.Cells(i, 8)
            If .value >= 2 Then
                .Interior.Color = RGB(198, 239, 206)
            ElseIf .value < 1 Then
                .Interior.Color = RGB(255, 199, 206)
            End If
        End With
    Next i
End Sub

Sub GenerateTradingSummaryArray(ws As Worksheet, signalData() As Variant, signalCount As Long)
    Dim buyCount As Long, sellCount As Long
    Dim strongCount As Long, mediumCount As Long, weakCount As Long
    Dim i As Long, summaryRow As Long
    
    summaryRow = signalCount + 6
    
    ' Count signals
    For i = 1 To signalCount
        If signalData(i, 2) = "BUY" Then buyCount = buyCount + 1
        If signalData(i, 2) = "SELL" Then sellCount = sellCount + 1
        
        Select Case signalData(i, 3)
            Case "STRONG": strongCount = strongCount + 1
            Case "MEDIUM": mediumCount = mediumCount + 1
            Case "WEAK": weakCount = weakCount + 1
        End Select
    Next i
    
    With ws
        .Cells(summaryRow, 1).value = "TRADING SIGNALS SUMMARY"
        .Cells(summaryRow, 1).Font.Bold = True
        .Cells(summaryRow, 1).Font.Size = 12
        summaryRow = summaryRow + 1
        
        .Cells(summaryRow, 1).value = "Total Signals:"
        .Cells(summaryRow, 2).value = signalCount
        summaryRow = summaryRow + 1
        
        .Cells(summaryRow, 1).value = "Buy Signals:"
        .Cells(summaryRow, 2).value = buyCount
        .Cells(summaryRow, 2).Interior.Color = RGB(198, 239, 206)
        summaryRow = summaryRow + 1
        
        .Cells(summaryRow, 1).value = "Sell Signals:"
        .Cells(summaryRow, 2).value = sellCount
        .Cells(summaryRow, 2).Interior.Color = RGB(255, 199, 206)
        summaryRow = summaryRow + 1
        
        .Cells(summaryRow, 1).value = "Strong Signals:"
        .Cells(summaryRow, 2).value = strongCount
        summaryRow = summaryRow + 1
        
        .Cells(summaryRow, 1).value = "Medium Signals:"
        .Cells(summaryRow, 2).value = mediumCount
        summaryRow = summaryRow + 1
        
        .Cells(summaryRow, 1).value = "Weak Signals:"
        .Cells(summaryRow, 2).value = weakCount
        summaryRow = summaryRow + 1
        
        ' Risk management guidelines
        summaryRow = summaryRow + 1
        .Cells(summaryRow, 1).value = "RISK MANAGEMENT GUIDELINES"
        .Cells(summaryRow, 1).Font.Bold = True
        summaryRow = summaryRow + 1
        
        .Cells(summaryRow, 1).value = "� Position Size: Based on volatility (2-8%)"
        summaryRow = summaryRow + 1
        .Cells(summaryRow, 1).value = "� Stop Loss: 2 x ATR from entry"
        summaryRow = summaryRow + 1
        .Cells(summaryRow, 1).value = "� Target: 4 x ATR from entry (2:1 R/R)"
        summaryRow = summaryRow + 1
        .Cells(summaryRow, 1).value = "� Max Portfolio Risk: 1-2% per trade"
    End With
End Sub


' ADD THIS FUNCTION TO YOUR CODE
Function CollectionToArray(col As Collection) As String()
    Dim arr() As String
    If col.count > 0 Then
        ReDim arr(1 To col.count)
        Dim i As Long
        For i = 1 To col.count
            arr(i) = col(i)
        Next i
    Else
        ReDim arr(1 To 0) ' Empty array
    End If
    CollectionToArray = arr
End Function

' KEEP YOUR EXISTING HELPER FUNCTIONS:
' - CalculateWeightedScore
' - CalculateSignalQuality
' - HasVolumeConfirmation
' - IsFalsePositive
' - GetMarketRegime
' - GetRegimeAdjustedThreshold
' - LogPerformanceMetrics
' - DisplayCompletionMessage
' - HandleProcessingError
