Attribute VB_Name = "CompleteTrading"
Option Explicit
' ENHANCED FILTERANDREPORT WITH COMPLETE SIGNAL GENERATION
Sub FilterAndReport_Enhanced()
    Dim wsTJX As Worksheet, wsDash As Worksheet, wsRptLog As Worksheet, wsRptHist As Worksheet, wsTLog As Worksheet
    Dim lastRowTJX As Long, tickerCount As Long, i As Long
    Dim groupSize As Long
    Dim priceThreshold As Double, minpriceThreshold As Double
    Dim analysisDate As Date, minScore As Variant
    Dim filterArray() As Variant
    Dim startTime As Double
    Dim totalIterations As Long
    Dim currentIteration As Long
    
    DoEvents
    If gStopMacro Then
        MsgBox "...E-Stopped!", vbInformation
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    startTime = Timer
    
    ' OPTIMIZATION: Disable all unnecessary features
    Application.ScreenUpdating = False
    Application.CALCULATION = xlCalculationManual
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
    
    ' Define constants
    groupSize = 50
    
    ' Get parameters
    minScore = wsDash.Range("W5").value
    minpriceThreshold = wsTJX.Range("C1").value
    priceThreshold = wsTJX.Range("E1").value
    analysisDate = wsDash.Range("H5").value

    Dim skipPrompt As Boolean
    skipPrompt = pubNotice Or perfTest
    
    If Not skipPrompt Then
        If Not ConfirmProcessing() Then
            If Not GetUserInputs(minScore, minpriceThreshold, priceThreshold, analysisDate) Then
                Exit Sub
            End If
        End If
    End If

    ' Clear ranges
    wsRptHist.Range("A4:G1000").ClearContents
    wsTLog.Range("B4:AF53").ClearContents
    
    ' Setup dashboard
    With wsDash
        .Range("A8:A57").ClearContents
        .Range("H5").value = analysisDate
        .Range("W5").value = minScore
        .Range("Y5").value = CStr(priceThreshold) & " Max Price"
        .Range("A3:AQ3").Copy
        .Range("A8:AQ57").PasteSpecial Paste:=xlPasteFormulas
    End With
    Application.CutCopyMode = False
    
    ' Read all ticker data
    lastRowTJX = wsTJX.Cells(wsTJX.Rows.count, "A").End(xlUp).row
    tickerCount = lastRowTJX - 2
    
    Dim tickerData As Variant
    tickerData = wsTJX.Range("A3:A" & lastRowTJX).value
    
    ReDim filterArray(1 To tickerCount)
    For i = 1 To tickerCount
        filterArray(i) = tickerData(i, 1)
    Next i

    totalIterations = Application.WorksheetFunction.Ceiling(tickerCount / groupSize, 1)
    currentIteration = 0

    ' Pre-allocate result arrays
    Dim allResults() As Variant
    Dim totalResultCount As Long
    ReDim allResults(1 To tickerCount, 1 To 7)
    totalResultCount = 0

    ' Collection for ALL qualifying tickers
    Dim allQualifyingTickers As Collection
    Set allQualifyingTickers = New Collection

    ' Process in batches
    Dim batchStart As Long: batchStart = 1
    Do While batchStart <= tickerCount
        currentIteration = currentIteration + 1
        
        ' Update status
        If currentIteration Mod 10 = 0 Or currentIteration = totalIterations Then
            Application.StatusBar = "Processing... " & currentIteration & " of " & totalIterations & " (" & Format((currentIteration / totalIterations) * 100, "0") & "%)"
        End If

        Dim batchEnd As Long
        batchEnd = WorksheetFunction.min(batchStart + groupSize - 1, tickerCount)
        Dim batchSize As Long
        batchSize = batchEnd - batchStart + 1
        
        ' Load batch
        wsDash.Range("A8:A" & (7 + batchSize)).ClearContents
        
        Dim batchRange As Range
        Set batchRange = wsDash.Range("A8:A" & (7 + batchSize))
        
        Dim batchData() As Variant
        ReDim batchData(1 To batchSize, 1 To 1)
        
        For i = 1 To batchSize
            batchData(i, 1) = filterArray(batchStart + i - 1)
        Next i
        
        batchRange.value = batchData

        ' Process batch data
        Call DataFromBackup(analysisDate)
        ' Note: CalculateRSISignals removed — it was a dev/test tool targeting "RSITest"
        ' sheet with InputBox prompts. RSI is calculated by CalculateEnhancedIndicators below.

        ' Calculate and process
        Application.CALCULATION = xlCalculationAutomatic
        DoEvents
        Application.CALCULATION = xlCalculationManual
        
        ' Enhanced processing with ticker collection
        Call ProcessTickersUltraFast_WithCollection(wsDash, analysisDate, minScore, minpriceThreshold, priceThreshold, batchSize, allResults, totalResultCount, allQualifyingTickers)
        
        batchStart = batchEnd + 1

        If currentIteration Mod 20 = 0 Then
            DoEvents
            If gStopMacro Then
                MsgBox "Macro stopped by user.", vbInformation
                GoTo CleanExit
            End If
        End If
    Loop

    ' Write results
    If totalResultCount > 0 Then
        Dim j As Integer
        Dim finalResults() As Variant
        ReDim finalResults(1 To totalResultCount, 1 To 7)
        
        For i = 1 To totalResultCount
            For j = 1 To 7
                finalResults(i, j) = allResults(i, j)
            Next j
        Next i
        
        With wsRptHist
            .Range("A4").Resize(totalResultCount, 7).value = finalResults
            .Range("A3").value = "Date"
            .Range("B3").value = "Ticker"
            .Range("C3").value = "Weighted Score"
            .Range("D3").value = "Company"
            .Range("E3").value = "Price"
            .Range("F3").value = "Signal Quality"
            .Range("G3").value = "Market Regime"
            .Columns("A:G").AutoFit
        End With
    End If

    Application.CALCULATION = xlCalculationAutomatic
    
    ' Generate trading signals for ALL qualifying tickers
    If allQualifyingTickers.count > 0 Then
        ' Create array from collection
        Dim qualifyingTickers() As String
        ReDim qualifyingTickers(1 To allQualifyingTickers.count)
        Dim k As Long
        For k = 1 To allQualifyingTickers.count
            qualifyingTickers(k) = allQualifyingTickers(k)
        Next k
        
        ' Load ALL qualifying tickers and calculate indicators
        Call LoadAllQualifyingTickersData(wsDash, qualifyingTickers, analysisDate)
        Call CalculateEnhancedIndicators
        Call UpdateSystemWithATR_Complete  ' ? ATR CALCULATION
        
        ' Generate final trading signals
        Call GenerateCompleteTradingSignals_Integrated
    End If
    
    Call theReporter

CleanExit:
    If Not skipPrompt Then
        Call DisplayCompletionMessage(startTime)
        wsDash.Range("A8:A57").Font.Size = 10
    End If
    
    Application.StatusBar = "Completed: " & totalResultCount & " qualified tickers from " & tickerCount & " total"
    Exit Sub

ErrorHandler:
    HandleProcessingError "FilterAndReport_Enhanced", Err
    Resume Next
    GoTo CleanExit
End Sub
' HELPER: Check if array is empty
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
' INTEGRATED TRADING SIGNAL GENERATION
Sub GenerateCompleteTradingSignals_Integrated()
    Dim wsData As Worksheet, wsDash As Worksheet, wsSignals As Worksheet
    Dim lastRow As Long, signalCount As Long
    Dim startTime As Double
    
    startTime = Timer
    Application.ScreenUpdating = False
    Application.CALCULATION = xlCalculationManual
    
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsDash = ThisWorkbook.Sheets("DashBoard")
    Set wsSignals = GetOrCreateSheet("TradingSignals")
    
    ' Clear and setup signals sheet
    Call SetupTradingSignalsSheet(wsSignals)
    
    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).row
    If lastRow < 3 Then  ' Fixed: was 50 (blocked runs with few tickers); need at least header + 2 data rows
        Debug.Print "No data rows for signal generation (lastRow=" & lastRow & ")"
        GoTo Cleanup
    End If

    ' Get unique tickers from current Data sheet
    Dim tickers() As String
    Dim tickerCount As Long
    tickerCount = GetUniqueTickersFromData(wsData, tickers, lastRow)
    
    ' Generate comprehensive trading signals
    signalCount = ProcessTradingSignalsArray(wsData, wsDash, wsSignals, tickers, tickerCount, lastRow)
    
Cleanup:
    Application.CALCULATION = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If signalCount > 0 Then
        Debug.Print "Trading signals generated: " & signalCount & " actionable signals"
    End If
End Sub

' UPDATED: ProcessTickersUltraFast that collects qualifying tickers
Sub ProcessTickersUltraFast_WithCollection(wsDash As Worksheet, aDate As Date, minScore As Variant, minpriceThreshold As Double, priceThreshold As Double, batchSize As Long, ByRef allResults() As Variant, ByRef totalResultCount As Long, ByRef qualifyingTickers As Collection)
    Dim i As Long, j As Long
    Dim minScoreVal As Double
    Dim minCompScore As Double
    Dim tickCountry As String
    Dim maxMarketCap As Double, minMarketCap As Double
    
    minScoreVal = CDbl(minScore)
    minCompScore = wsDash.Range("R5").value
    tickCountry = wsDash.Range("S5").value
    maxMarketCap = wsDash.Range("C5").value
    minMarketCap = wsDash.Range("C6").value
    
    ' OPTIMIZATION: Read ALL required data in SINGLE operation
    Dim batchData As Variant
    batchData = wsDash.Range("A8:AQ" & (7 + batchSize)).value
    
    ' Column indices for maintainability
    Const COL_TICKER As Long = 1
    Const COL_COMPANY As Long = 2
    Const COL_MKTCAP As Long = 3
    Const COL_VOLUME As Long = 4         ' D
    Const COL_SCORE_START As Long = 7    ' G
    Const COL_COMPSCORE As Long = 18     ' R
    Const COL_COUNTRY As Long = 19       ' S
    Const COL_PRICE As Long = 25         ' Y
    
    
    ' RESEARCH-BACKED INDICATOR WEIGHTS
    Dim indicatorWeights(1 To 10) As Double
    indicatorWeights(1) = 1.6   ' RSI
    indicatorWeights(2) = 1.4   ' MACD
    indicatorWeights(3) = 1.3   ' Volume
    indicatorWeights(4) = 1.1   ' ATR
    indicatorWeights(5) = 1#    ' Price action
    indicatorWeights(6) = 0#    ' Stochastic
    indicatorWeights(7) = 0#    ' Williams %R
    indicatorWeights(8) = 0#    ' CCI
    indicatorWeights(9) = 0.3   ' OBV
    indicatorWeights(10) = 0.6  ' ADX
    
    ' Get market regime
    Dim marketRegime As String
    marketRegime = GetMarketRegime(wsDash)
    
    ' Adjust threshold based on market regime
    Dim adjustedMinScore As Double
    adjustedMinScore = GetRegimeAdjustedThreshold(marketRegime, minScoreVal)
    
    ' Process each ticker with enhanced logic
    For i = 1 To batchSize
        If Not IsEmpty(batchData(i, COL_TICKER)) And IsNumeric(batchData(i, COL_PRICE)) Then
            Dim price As Double, cScore As Double, mCap As Double
            Dim tickO As String
            
            price = CDbl(batchData(i, COL_PRICE))
            cScore = CDbl(batchData(i, COL_COMPSCORE))
            tickO = CStr(batchData(i, COL_COUNTRY))
            mCap = CDbl(batchData(i, COL_MKTCAP))
            
            ' Quick filters with early exits
            If price < minpriceThreshold Or price > priceThreshold Then GoTo NextTicker
            If cScore < minCompScore Then GoTo NextTicker
            
            ' Calculate weighted score with signal quality
            Dim weightedScore As Double
            Dim signalQuality As Double
            Dim volumeConfirmation As Boolean
            
            weightedScore = CalculateWeightedScore(batchData, i, COL_SCORE_START, indicatorWeights)
            signalQuality = CalculateSignalQuality(batchData, i, COL_SCORE_START, indicatorWeights)
            volumeConfirmation = HasVolumeConfirmation(batchData, i, COL_VOLUME, weightedScore)
            
            ' Check for false positives
            If IsFalsePositive(CStr(batchData(i, COL_TICKER)), weightedScore, aDate) Then GoTo NextTicker
            
            ' Multi-factor qualification — quality floor lowered to 0.4; volume confirmation
            ' is preferred but not mandatory when signal quality is high (>= 0.7)
            Dim qualifies As Boolean
            qualifies = (Abs(weightedScore) >= adjustedMinScore) And _
                        (signalQuality >= 0.4) And _
                        (volumeConfirmation Or signalQuality >= 0.7)
            If qualifies Then
               
                totalResultCount = totalResultCount + 1
                
                ' Ensure array is large enough
                If totalResultCount > UBound(allResults, 1) Then
                    ReDim Preserve allResults(1 To totalResultCount + 100, 1 To 7)
                End If
                
                ' Store enhanced results
                allResults(totalResultCount, 1) = aDate
                allResults(totalResultCount, 2) = batchData(i, COL_TICKER)
                allResults(totalResultCount, 3) = weightedScore
                allResults(totalResultCount, 4) = batchData(i, COL_COMPANY)
                allResults(totalResultCount, 5) = price
                allResults(totalResultCount, 6) = signalQuality
                allResults(totalResultCount, 7) = marketRegime
                
                ' NEW: Add to qualifying tickers collection for signal generation
                On Error Resume Next ' Ignore duplicates
                qualifyingTickers.Add batchData(i, COL_TICKER)
                On Error GoTo 0
            End If
        End If
        
NextTicker:
    Next i
End Sub

' NEW: Generate trading signals for ALL qualifying tickers
Sub GenerateTradingSignalsForAllTickers(qualifyingTickers As Collection, analysisDate As Date)
    Dim wsData As Worksheet, wsDash As Worksheet, wsSignals As Worksheet
    Dim i As Long, signalCount As Long
    Dim startTime As Double
    
    startTime = Timer
    Application.ScreenUpdating = False
    Application.CALCULATION = xlCalculationManual
    
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsDash = ThisWorkbook.Sheets("DashBoard")
    Set wsSignals = GetOrCreateSheet("TradingSignals")
    
    ' Clear and setup signals sheet
    Call SetupTradingSignalsSheet(wsSignals)
    
    ' Process ALL qualifying tickers in one batch
    Dim tickers() As String
    ReDim tickers(1 To qualifyingTickers.count)
    
    For i = 1 To qualifyingTickers.count
        tickers(i) = qualifyingTickers(i)
    Next i
    
    ' Load ALL qualifying tickers data at once
    Call LoadAllQualifyingTickersData(wsDash, tickers, analysisDate)
    
    ' Calculate indicators for all tickers
    Call CalculateEnhancedIndicators
    Call CalculateATRWithSignals_Optimized
    
    ' Generate signals for all tickers
    signalCount = ProcessTradingSignalsArray(wsData, wsDash, wsSignals, tickers, qualifyingTickers.count, wsData.Cells(wsData.Rows.count, "A").End(xlUp).row)
    
    Application.CALCULATION = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Debug.Print "Trading signals generated for " & signalCount & " tickers in " & Format(Timer - startTime, "0.00") & " seconds"
End Sub

' ENHANCED: Load data for all qualifying tickers at once
Sub LoadAllQualifyingTickersData(wsDash As Worksheet, tickers() As String, analysisDate As Date)
    Dim i As Long
    Dim lastRow As Long
    
    ' Clear dashboard
    wsDash.Range("A8:A57").ClearContents
    
    ' Load all qualifying tickers
    lastRow = 7 + UBound(tickers, 1)
    If lastRow > 57 Then lastRow = 57 ' Limit to dashboard capacity
    
    For i = 1 To UBound(tickers, 1)
        If 7 + i <= 57 Then ' Stay within dashboard bounds
            wsDash.Cells(7 + i, 1).value = tickers(i)
        End If
    Next i
    
    ' Fetch historical data for all qualifying tickers
    Call DataFromBackup(analysisDate)
    
    ' Calculate indicators (CalculateRSISignals removed — dev-only tool, RSI handled by CalculateEnhancedIndicators)
    Application.CALCULATION = xlCalculationAutomatic
    DoEvents
    Application.CALCULATION = xlCalculationManual
End Sub

' COMPLETE TRADING SIGNAL GENERATION WITH RISK MANAGEMENT
Sub GenerateCompleteTradingSignals()
    ' Ensure all calculations are up to date
    Call CalculateEnhancedIndicators
    Call CalculateATRWithSignals_Optimized
    Call GenerateTradingSignalsWithRiskManagement
End Sub

' MAIN SIGNAL GENERATION FUNCTION
Sub GenerateTradingSignalsWithRiskManagement()
    Dim wsData As Worksheet, wsDash As Worksheet, wsSignals As Worksheet
    Dim lastRow As Long, i As Long, signalCount As Long
    Dim startTime As Double
    
    startTime = Timer
    Application.ScreenUpdating = False
    Application.CALCULATION = xlCalculationManual
    
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsDash = ThisWorkbook.Sheets("DashBoard")
    Set wsSignals = GetOrCreateSheet("TradingSignals")
    
    ' Clear and setup signals sheet
    Call SetupTradingSignalsSheet(wsSignals)
    
    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).row
    If lastRow < 3 Then  ' Fixed: was 50
        MsgBox "No data rows found for signal generation", vbExclamation
        GoTo Cleanup
    End If

    ' Get unique tickers
    Dim tickers() As String
    Dim tickerCount As Long
    tickerCount = GetUniqueTickersFromData(wsData, tickers, lastRow)

    ' Process signals using array approach
    signalCount = ProcessTradingSignalsArray(wsData, wsDash, wsSignals, tickers, tickerCount, lastRow)
    
Cleanup:
    Application.CALCULATION = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Trading signals generated: " & signalCount & " signals found" & vbCrLf & _
           "Processing time: " & Format(Timer - startTime, "0.00") & " seconds", _
           vbInformation, "Signal Generation Complete"
End Sub

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

' GET TICKER SIGNAL DATA
Function GetTickerSignalData(ws As Worksheet, wsDash As Worksheet, ticker As String, lastRow As Long) As Variant
    Dim i As Long
    Dim signalArray(0 To 16) As Variant
    
    ' Initialize with empty values
    For i = 0 To 16
        signalArray(i) = ""
    Next i
    
    ' Find the most recent data for this ticker
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 7).value = ticker Then
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
            volumeSpike = Nz(ws.Cells(i, 16).value, 1)  ' Fixed: col 16 = VolumeSpike (was col 13 = PriceVsMA)
            ' Calculate buy/sell scores
            Dim buyScore As Double, sellScore As Double
            buyScore = CalculateBuyScoreSimple(rsi, macd, macdSignal, priceVsMA, compositeScore, volumeSpike, atrPercent)
            sellScore = CalculateSellScoreSimple(rsi, macd, macdSignal, priceVsMA, compositeScore, volumeSpike, atrPercent)
            
            ' Get threshold parameters from dashboard
            Dim minScore As Double, maxScore As Double
            minScore = Nz(wsDash.Range("W5").value, 2)
            maxScore = Nz(wsDash.Range("X5").value, -2)
            
            ' Determine signal
            Dim signalType As String, signalStrength As String
            Dim stopLoss As Double, positionSize As Double, riskPerShare As Double, riskRewardRatio As Double
            
            If buyScore >= minScore And buyScore > sellScore Then
                signalType = "BUY"
                signalStrength = GetSignalStrengthSimple(buyScore, minScore)
                Call CalculateBuyRiskManagementSimple(entryPrice, atr, atrPercent, stopLoss, positionSize, riskPerShare, riskRewardRatio)
            ElseIf sellScore <= maxScore And sellScore < buyScore Then
                signalType = "SELL"
                signalStrength = GetSignalStrengthSimple(Abs(sellScore), Abs(maxScore))
                Call CalculateSellRiskManagementSimple(entryPrice, atr, atrPercent, stopLoss, positionSize, riskPerShare, riskRewardRatio)
            Else
                ' HOLD signal - exit without returning data
                Exit Function
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
    Next i
    
    ' No signal found
    GetTickerSignalData = False
End Function

' SIMPLIFIED CALCULATION FUNCTIONS
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
    
    ' Price vs MA - bullish only when above MA (removed erroneous +5 for slightly-below-MA)
    If priceVsMA > 0 Then
        score = score + (priceVsMA * 2)
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
Sub OutputTradingSignalsArray(ws As Worksheet, signalData() As Variant, signalCount As Long)
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

' HELPER FUNCTIONS
Sub SetupTradingSignalsSheet(ws As Worksheet)
    ws.Cells.Clear
    
    With ws
        ' Main headers
        .Range("A1").value = "Trading Signals - " & Format(Date, "yyyy-mm-dd")
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        
        ' Column headers
        Dim headers(1 To 17) As String
        headers(1) = "Ticker"
        headers(2) = "Signal"
        headers(3) = "Strength"
        headers(4) = "Entry Price"
        headers(5) = "Stop Loss"
        headers(6) = "Position Size %"
        headers(7) = "Risk/Share"
        headers(8) = "R/R Ratio"
        headers(9) = "Composite Score"
        headers(10) = "RSI"
        headers(11) = "MACD"
        headers(12) = "MACD Signal"
        headers(13) = "Price vs MA"
        headers(14) = "ATR"
        headers(15) = "ATR %"
        headers(16) = "Volume Spike"
        headers(17) = "Timestamp"
        
        .Range("A3:Q3").value = headers
        .Range("A3:Q3").Font.Bold = True
        .Range("A3:Q3").Interior.Color = RGB(200, 200, 200)
        .Range("A3:Q3").HorizontalAlignment = xlCenter
    End With
End Sub

Function GetOrCreateSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Sheets(sheetName)
    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Sheets.Add
        GetOrCreateSheet.Name = sheetName
    End If
    On Error GoTo 0
End Function

Function GetUniqueTickersFromData(ws As Worksheet, ByRef tickers() As String, lastRow As Long) As Long
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 2 To lastRow
        If Not dict.Exists(ws.Cells(i, 7).value) Then
            dict.Add ws.Cells(i, 7).value, 1
        End If
    Next i
    
    GetUniqueTickersFromData = dict.count
    ReDim tickers(1 To dict.count)
    
    Dim key As Variant, idx As Long
    idx = 1
    For Each key In dict.Keys
        tickers(idx) = key
        idx = idx + 1
    Next key
End Function

Function Nz(value As Variant, defaultVal As Double) As Double
    If IsEmpty(value) Or IsError(value) Or value = "" Then
        Nz = defaultVal
    Else
        Nz = CDbl(value)
    End If
End Function

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

