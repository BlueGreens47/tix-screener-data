Attribute VB_Name = "EnhancedDS"
Option Explicit

'Focus: quality, context-awareness, and confirmation
'Recommended Action Plan:
'Phase 1: Reduce to 5 core indicators + volume
'Phase 2: Implement market regime detection
'Phase 3: Add signal quality scoring
'Phase 4: Implement false positive filtering

'Recommended Core Set:

Type EssentialIndicators
    RSI As Double           ' Momentum
    MACD As Double          ' Trend + Momentum
    ATR As Double           ' Volatility
    VolumeProfile As Double ' Volume confirmation
    PriceAction As Double   ' Support/Resistance
End Type

'''Accuracy Improvement Strategies'''
'1. Feature Engineering

' ENHANCED SCORING LOGIC
Function CalculateEnhancedScore(values As Variant, price As Double, volume As Variant) As Double
    Dim score As Double
    Dim momentumScore As Double, volatilityScore As Double, volumeScore As Double
    
    ' 1. WEIGHTED MOMENTUM (not all indicators are equal)
    momentumScore = CalculateWeightedMomentum(values)
    
    ' 2. VOLATILITY-ADJUSTED SIGNALS
    volatilityScore = CalculateVolatilityAdjustedScore(values, price)
    
    ' 3. VOLUME CONFIRMATION (critical for accuracy)
    volumeScore = CalculateVolumeConfirmation(volume)
    
    ' 4. COMPOSITE SCORE WITH MARKET REGIME ADJUSTMENT
    score = momentumScore * 0.6 + volatilityScore * 0.3 + volumeScore * 0.1
    
    CalculateEnhancedScore = score
End Function

Function CalculateWeightedMomentum(values As Variant) As Double
    ' Give more weight to recent signals and reliable indicators
    Dim weights(1 To 10) As Double
    weights(1) = 1.2  ' Recent RSI
    weights(2) = 1.1  ' MACD
    weights(3) = 0.9  ' Older signals
    weights(4) = 0.9
    weights(5) = 0.8
    weights(6) = 0.8
    weights(7) = 0.7
    weights(8) = 0.7
    weights(9) = 0.6
    weights(10) = 0.6
    
    Dim weightedSum As Double
    Dim i As Long
    For i = 1 To 10
        If IsNumeric(values(i)) Then
            Dim val As Double
            val = CDbl(values(i))
            If Abs(val) >= 1 Then
                weightedSum = weightedSum + weights(i) * val
            End If
        End If
    Next i
    
    CalculateWeightedMomentum = weightedSum
End Function

'2. Market Regime Detection
Function GetMarketRegime() As String
    ' Determine if market is trending, ranging, volatile
    Dim volatility As Double, trendStrength As Double
    
    ' Calculate market regime from SPY or similar
    volatility = CalculateMarketVolatility
    trendStrength = CalculateTrendStrength
    
    If volatility > 0.25 Then
        GetMarketRegime = "HIGH_VOLATILITY"
    ElseIf trendStrength > 0.7 Then
        GetMarketRegime = "STRONG_TREND"
    ElseIf trendStrength < 0.3 Then
        GetMarketRegime = "RANGING"
    Else
        GetMarketRegime = "NORMAL"
    End If
End Function

Function GetRegimeAdjustedThreshold(regime As String, baseThreshold As Double) As Double
    ' Adjust thresholds based on market conditions
    Select Case regime
        Case "HIGH_VOLATILITY"
            GetRegimeAdjustedThreshold = baseThreshold * 1.3
        Case "STRONG_TREND"
            GetRegimeAdjustedThreshold = baseThreshold * 0.8
        Case "RANGING"
            GetRegimeAdjustedThreshold = baseThreshold * 1.1
        Case Else
            GetRegimeAdjustedThreshold = baseThreshold
    End Select
End Function

'3. Signal Quality Metrics
Function CalculateSignalQuality(values As Variant) As Double
    ' Measure signal consistency and strength
    Dim signalCount As Long, consistentSigns As Long
    Dim i As Long
    Dim firstSign As Integer
    
    signalCount = 0
    consistentSigns = 0
    
    For i = 1 To 10
        If IsNumeric(values(i)) Then
            Dim val As Double
            val = CDbl(values(i))
            If Abs(val) >= 1 Then
                signalCount = signalCount + 1
                If signalCount = 1 Then
                    firstSign = Sgn(val)
                ElseIf Sgn(val) = firstSign Then
                    consistentSigns = consistentSigns + 1
                End If
            End If
        End If
    Next i
    
    If signalCount > 0 Then
        ' Quality score: consistency ratio * signal strength
        CalculateSignalQuality = (consistentSigns / signalCount) * (signalCount / 10)
    Else
        CalculateSignalQuality = 0
    End If
End Function

'''Indicator Rationalization Strategy'''

'1. Correlation Analysis
Sub AnalyzeIndicatorCorrelation()
    ' Run this periodically to identify redundant indicators
    Dim correlationMatrix(1 To 10, 1 To 10) As Double
    Dim i As Long, j As Long
    
    For i = 1 To 10
        For j = 1 To 10
            correlationMatrix(i, j) = CalculateCorrelation( _
                GetIndicatorHistory(i), GetIndicatorHistory(j))
        Next j
    Next i
    
    ' Remove indicators with correlation > 0.8
    ' NOTE: RemoveRedundantIndicators has no active implementation — correlation pruning
    ' is handled manually via indicator weights in CompleteTrading.bas (RSI 1.6x, MACD 1.4x etc.)
    ' Call RemoveRedundantIndicators(correlationMatrix, 0.8)
End Sub

'2. Feature Importance Ranking
Function RankIndicatorImportance() As Collection
    ' Backtest each indicator individually
    Dim importanceScores As New Collection
    Dim i As Long
    
    For i = 1 To 10
        Dim performance As Double
        performance = BacktestSingleIndicator(i)
        importanceScores.Add performance, CStr(i)
    Next i
    
    ' Sort by performance and keep top 5
    Set RankIndicatorImportance = SortCollectionByValue(importanceScores)
End Function

'3. Simplified Smarter Processing
Sub ProcessTickersSmart(wsDash As Worksheet, config As FilterConfig, ByRef results() As Variant, ByRef resultCount As Long)
    Dim batchData As Variant
    batchData = wsDash.Range("A8:AQ" & (7 + config.batchSize)).value
    
    For i = 1 To config.batchSize
        If IsQualifyingTicker(batchData, i, config) Then
            Dim essentialSignals As EssentialIndicators
            essentialSignals = ExtractEssentialSignals(batchData, i)
            
            Dim compositeScore As Double
            compositeScore = CalculateCompositeScore(essentialSignals, config)
            
            Dim signalQuality As Double
            signalQuality = CalculateSignalQuality(batchData, i)
            
            ' Only accept high-quality signals
            If compositeScore >= config.minScore And signalQuality >= 0.7 Then
                StoreQualifiedTicker results, resultCount, batchData, i, compositeScore
            End If
        End If
    Next i
End Sub

'''Key Accuracy Improvements'''

'1. Add False Positive Filtering
Function IsFalsePositive(ticker As String, signal As Double, historicalData As Variant) As Boolean
    ' Filter out common false signal patterns
    Dim recentPerformance As Double
    recentPerformance = GetRecentPerformance(ticker, 5) ' 5 days
    
    ' Avoid buying into strong downtrends or selling into strong uptrends
    If signal > 0 And recentPerformance < -0.08 Then  ' 8% drop
        IsFalsePositive = True
        Exit Function
    End If
    
    If signal < 0 And recentPerformance > 0.08 Then   ' 8% gain
        IsFalsePositive = True
        Exit Function
    End If
    
    IsFalsePositive = False
End Function

'2. Volume Confirmation
Function HasVolumeConfirmation(ticker As String, signalType As String) As Boolean
    Dim volumeRatio As Double
    volumeRatio = CalculateVolumeRatio(ticker) ' Current vs average volume
    
    If signalType = "BREAKOUT" And volumeRatio < 1.2 Then
        HasVolumeConfirmation = False
    ElseIf signalType = "REVERSAL" And volumeRatio < 0.8 Then
        HasVolumeConfirmation = False
    Else
        HasVolumeConfirmation = True
    End If
End Function

