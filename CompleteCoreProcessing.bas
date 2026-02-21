Attribute VB_Name = "CompleteCoreProcessing"
Option Explicit

'******** 2. COMPLETE CORE PROCESSING FUNCTIONS ********
Private Sub ProcessTickersUltraFast_WithCollection(wsDash As Worksheet, aDate As Date, minScore As Variant, minpriceThreshold As Double, priceThreshold As Double, batchSize As Long, ByRef allResults() As Variant, ByRef totalResultCount As Long, ByRef qualifyingTickers As Collection)  ' Private: canonical version in CompleteTrading.bas
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
    
    Dim batchData As Variant
    batchData = wsDash.Range("A8:AQ" & (7 + batchSize)).value
    
    Const COL_TICKER As Long = 1
    Const COL_COMPANY As Long = 2
    Const COL_MKTCAP As Long = 3
    Const COL_VOLUME As Long = 4
    Const COL_SCORE_START As Long = 7
    Const COL_COMPSCORE As Long = 18
    Const COL_COUNTRY As Long = 19
    Const COL_PRICE As Long = 25
    
    
    Dim indicatorWeights(1 To 10) As Double
    indicatorWeights(1) = 1.6
    indicatorWeights(2) = 1.4
    indicatorWeights(3) = 1.3
    indicatorWeights(4) = 1.1
    indicatorWeights(5) = 1#
    indicatorWeights(6) = 0#
    indicatorWeights(7) = 0#
    indicatorWeights(8) = 0#
    indicatorWeights(9) = 0.3
    indicatorWeights(10) = 0.6
    
    Dim marketRegime As String
    marketRegime = GetMarketRegime(wsDash)
    
    Dim adjustedMinScore As Double
    adjustedMinScore = GetRegimeAdjustedThreshold(marketRegime, minScoreVal)
    
    
    For i = 1 To batchSize
        If Not IsEmpty(batchData(i, COL_TICKER)) And IsNumeric(batchData(i, COL_PRICE)) Then
            Dim price As Double, cScore As Double, mCap As Double
            Dim tickO As String, tempTick As String
            
            price = CDbl(batchData(i, COL_PRICE))
            cScore = CDbl(batchData(i, COL_COMPSCORE))
            tickO = CStr(batchData(i, COL_COUNTRY))
            mCap = CDbl(batchData(i, COL_MKTCAP))
            tempTick = CStr(batchData(i, COL_TICKER))
            
            If price < minpriceThreshold Or price > priceThreshold Then GoTo NextTicker
            If cScore < minCompScore Then GoTo NextTicker
            
            Dim weightedScore As Double
            Dim signalQuality As Double
            Dim volumeConfirmation As Boolean
            
            weightedScore = CalculateWeightedScore(batchData, i, COL_SCORE_START, indicatorWeights)
            signalQuality = CalculateSignalQuality(batchData, i, COL_SCORE_START, indicatorWeights)
            volumeConfirmation = HasVolumeConfirmation(batchData, i, COL_VOLUME, weightedScore)
            
             
            If IsFalsePositive(tempTick, weightedScore, aDate) Then GoTo NextTicker
            
            If Abs(weightedScore) >= adjustedMinScore And _
               signalQuality >= 0.6 And _
               volumeConfirmation Then
               
                totalResultCount = totalResultCount + 1
                
                If totalResultCount > UBound(allResults, 1) Then
                    ReDim Preserve allResults(1 To totalResultCount + 100, 1 To 7)
                End If
                
                allResults(totalResultCount, 1) = aDate
                allResults(totalResultCount, 2) = batchData(i, COL_TICKER)
                allResults(totalResultCount, 3) = weightedScore
                allResults(totalResultCount, 4) = batchData(i, COL_COMPANY)
                allResults(totalResultCount, 5) = price
                allResults(totalResultCount, 6) = signalQuality
                allResults(totalResultCount, 7) = marketRegime
                
                On Error Resume Next
                qualifyingTickers.Add batchData(i, COL_TICKER)
                On Error GoTo 0
            End If
        End If
        
NextTicker:
'Debug.Print qualifyingTickers.count
    Next i
    
End Sub

Sub LoadAllQualifyingTickersData(wsDash As Worksheet, tickers() As String, analysisDate As Date)
    Dim i As Long
    Dim lastRow As Long
    
    wsDash.Range("A8:A57").ClearContents
    
    lastRow = 7 + UBound(tickers, 1)
    If lastRow > 57 Then lastRow = 57
    
    For i = 1 To UBound(tickers, 1)
        If 7 + i <= 57 Then
            wsDash.Cells(7 + i, 1).value = tickers(i)
        End If
    Next i
    
    'Call DataFromBackup(analysisDate)
    'Call CalculateRSISignals
    Application.CALCULATION = xlCalculationAutomatic
    DoEvents
    Application.CALCULATION = xlCalculationManual
End Sub
'***2. CORE PROCESSING END ***
  
