Attribute VB_Name = "ATRCalculation"
Option Explicit

' COMPREHENSIVE ATR CALCULATION WITH SIGNAL GENERATION
Sub CalculateATRWithSignals()
    Dim wsData As Worksheet, wsDash As Worksheet
    Dim lastRow As Long, i As Long
    Dim ticker As String, currentTicker As String
    Dim atrPeriod As Long
    
    Set wsData = ThisWorkbook.Sheets("ATR")
    Set wsDash = ThisWorkbook.Sheets("DashBoard")
    atrPeriod = 14 ' Standard ATR period
    
    Application.ScreenUpdating = False
    Application.CALCULATION = xlCalculationManual
    
    lastRow = Application.WorksheetFunction.CountIf(wsData.Range("A2:A3200"), ">0")
    'lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).row
    If lastRow < atrPeriod + 1 Then Exit Sub
    
    ' Clear previous ATR calculations
    If wsData.Cells(1, 8).value = "True Range" Then
        wsData.Range("H:O").ClearContents
    End If
    
    ' Setup ATR headers
    Call SetupATRHeaders(wsData)
    
    ' Process each ticker
    currentTicker = wsData.Range("G2").value
    Dim startRow As Long: startRow = 7
    Dim endRow As Long: endRow = 7
    
    For i = 7 To lastRow
        If wsData.Cells(i, 7).value <> currentTicker Or i = lastRow Then
            If i = lastRow Then endRow = i Else endRow = i - 1
            
            If endRow - startRow >= atrPeriod Then
                Call CalculateTickerATR(wsData, startRow, endRow, atrPeriod)
            End If
            
            startRow = i
            currentTicker = wsData.Cells(i, 7).value
        End If
    Next i
    
    ' Generate ATR-based signals
    Call GenerateATRSignals(wsData, wsDash)
    
    Application.CALCULATION = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "ATR calculations and signals completed successfully", vbInformation
End Sub

Sub SetupATRHeaders(ws As Worksheet)
    With ws
        .Range("O1").value = "True Range"
        .Range("P1").value = "ATR"
        .Range("Q1").value = "ATR %"
        .Range("R1").value = "ATR Ratio"
        .Range("S1").value = "Volatility Zone"
        .Range("T1").value = "ATR Signal"
        .Range("U1").value = "Stop Loss Level"
        .Range("V1").value = "Position Size %"
        
        ' Format headers
        With .Range("O1:V1")
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
            .HorizontalAlignment = xlCenter
        End With
    End With
End Sub

Sub CalculateTickerATR(ws As Worksheet, startRow As Long, endRow As Long, period As Long)
    Dim i As Long, dataCount As Long
    Dim highPrices() As Double, lowPrices() As Double, closePrices() As Double
    Dim trueRanges() As Double, atrValues() As Double
    
    dataCount = endRow - startRow + 1
    ReDim highPrices(1 To dataCount)
    ReDim lowPrices(1 To dataCount)
    ReDim closePrices(1 To dataCount)
    ReDim trueRanges(1 To dataCount)
    ReDim atrValues(1 To dataCount)
    
    ' Load price data
    For i = 1 To dataCount
        highPrices(i) = ws.Cells(startRow + i - 1, 3).value ' High
        lowPrices(i) = ws.Cells(startRow + i - 1, 4).value  ' Low
        closePrices(i) = ws.Cells(startRow + i - 1, 5).value ' Close
    Next i
    
    ' Calculate True Range
    For i = 2 To dataCount
        Dim tr1 As Double, tr2 As Double, tr3 As Double
        tr1 = highPrices(i) - lowPrices(i) ' Current High - Current Low
        tr2 = Abs(highPrices(i) - closePrices(i - 1)) ' |Current High - Previous Close|
        tr3 = Abs(lowPrices(i) - closePrices(i - 1))  ' |Current Low - Previous Close|
        
        trueRanges(i) = Application.WorksheetFunction.max(tr1, tr2, tr3)
        ws.Cells(startRow + i - 1, 15).value = trueRanges(i) ' Column O
    Next i
    
    ' Calculate ATR (using Wilder's smoothing method)
    Dim atrSum As Double
    atrSum = 0
    
    ' First ATR is simple average of first 'period' TR values
    For i = 2 To period + 1
        atrSum = atrSum + trueRanges(i)
    Next i
    
    atrValues(period + 1) = atrSum / period
    ws.Cells(startRow + period, 16).value = atrValues(period + 1) ' Column P
    
    ' Subsequent ATR values using Wilder's smoothing
    For i = period + 2 To dataCount
        atrValues(i) = (atrValues(i - 1) * (period - 1) + trueRanges(i)) / period
        ws.Cells(startRow + i - 1, 16).value = atrValues(i) ' Column P
        
        ' Calculate ATR as percentage of price
        If closePrices(i) > 0 Then
            ws.Cells(startRow + i - 1, 17).value = (atrValues(i) / closePrices(i)) * 100 ' Column Q
        End If
        
        ' Calculate ATR Ratio (current ATR vs historical average)
        Call CalculateATRRatio(ws, startRow, i, atrValues, dataCount, period)
        
        ' Determine volatility zone and generate signals
        Call GenerateATRZoneAndSignal(ws, startRow, i, atrValues(i), closePrices(i))
        
        ' Calculate stop loss and position size
        Call CalculateRiskManagement(ws, startRow, i, atrValues(i), closePrices(i))
    Next i
End Sub

Sub CalculateATRRatio(ws As Worksheet, startRow As Long, currentIndex As Long, atrValues() As Double, dataCount As Long, period As Long)
    Dim i As Long
    Dim atrSum As Double, atrCount As Long
    Dim avgATR As Double, atrRatio As Double
    
    ' Calculate average ATR over the entire period (or last 50 periods)
    atrSum = 0
    atrCount = 0
    Dim lookback As Long: lookback = Application.WorksheetFunction.min(50, currentIndex - period)
    
    For i = currentIndex - lookback + 1 To currentIndex
        If atrValues(i) > 0 Then
            atrSum = atrSum + atrValues(i)
            atrCount = atrCount + 1
        End If
    Next i
    
    If atrCount > 0 Then
        avgATR = atrSum / atrCount
        atrRatio = atrValues(currentIndex) / avgATR
        ws.Cells(startRow + currentIndex - 1, 11).value = atrRatio ' Column K
    End If
End Sub

Sub GenerateATRZoneAndSignal(ws As Worksheet, startRow As Long, currentIndex As Long, currentATR As Double, currentPrice As Double)
    Dim atrPercent As Double, atrRatio As Double
    Dim volatilityZone As String, atrSignal As String
    
    atrPercent = ws.Cells(startRow + currentIndex - 1, 10).value ' ATR %
    atrRatio = Nz(ws.Cells(startRow + currentIndex - 1, 11).value, 1) ' ATR Ratio
    
    ' Determine Volatility Zone
    If atrPercent < 1.5 Then
        volatilityZone = "LOW VOL"
    ElseIf atrPercent >= 1.5 And atrPercent < 3 Then
        volatilityZone = "NORMAL VOL"
    ElseIf atrPercent >= 3 And atrPercent < 5 Then
        volatilityZone = "HIGH VOL"
    Else
        volatilityZone = "EXTREME VOL"
    End If
    
    ' Generate ATR Signal based on volatility regime
    If atrRatio > 1.5 Then
        ' High volatility - potential breakout
        atrSignal = "VOLATILITY_SPIKE"
    ElseIf atrRatio < 0.7 Then
        ' Low volatility - potential consolidation
        atrSignal = "VOLATILITY_CONTRACTION"
    ElseIf atrPercent > 4 Then
        ' Extreme volatility - caution
        atrSignal = "EXTREME_VOLATILITY"
    Else
        atrSignal = "NORMAL_VOLATILITY"
    End If
    
    ws.Cells(startRow + currentIndex - 1, 12).value = volatilityZone ' Column L
    ws.Cells(startRow + currentIndex - 1, 13).value = atrSignal ' Column M
End Sub

Sub CalculateRiskManagement(ws As Worksheet, startRow As Long, currentIndex As Long, currentATR As Double, currentPrice As Double)
    Dim stopLossLevel As Double, positionSizePercent As Double
    Dim atrPercent As Double
    
    atrPercent = (currentATR / currentPrice) * 100
    
    ' Calculate stop loss levels (2 * ATR for swing trading)
    stopLossLevel = currentPrice - (2 * currentATR)
    ws.Cells(startRow + currentIndex - 1, 14).value = stopLossLevel ' Column U
    
    ' Calculate position size based on volatility (inverse relationship)
    ' Higher volatility = smaller position size
    If atrPercent < 2 Then
        positionSizePercent = 8 ' Higher allocation in low volatility
    ElseIf atrPercent >= 2 And atrPercent < 3 Then
        positionSizePercent = 6 ' Medium allocation
    ElseIf atrPercent >= 3 And atrPercent < 5 Then
        positionSizePercent = 4 ' Lower allocation in high volatility
    Else
        positionSizePercent = 2 ' Minimal allocation in extreme volatility
    End If
    
    ws.Cells(startRow + currentIndex - 1, 15).value = positionSizePercent ' Column V
End Sub

Sub GenerateATRSignals(wsData As Worksheet, wsDash As Worksheet)
    Dim lastRow As Long, i As Long
    Dim currentTicker As String, previousTicker As String
    Dim atrSignals As Collection
    
    Set atrSignals = New Collection
    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).row
    
    ' Get the most recent ATR data for each ticker
    previousTicker = ""
    For i = lastRow To 2 Step -1
        currentTicker = wsData.Cells(i, 7).value
        
        If currentTicker <> previousTicker Then
            Dim atrSignal As atrSignal
            Set atrSignal = New atrSignal
            
            With atrSignal
                .ticker = currentTicker
                .Price = wsData.Cells(i, 5).value
                .ATR = Nz(wsData.Cells(i, 16).value, 0)
                .atrPercent = Nz(wsData.Cells(i, 17).value, 0)
                .atrRatio = Nz(wsData.Cells(i, 18).value, 1)
                .volatilityZone = wsData.Cells(i, 19).value
                .RawSignal = wsData.Cells(i, 20).value
                .StopLoss = Nz(wsData.Cells(i, 21).value, 0)
                .PositionSize = Nz(wsData.Cells(i, 22).value, 0)
                
                ' Generate trading signal
                Call .GenerateTradingSignal
            End With
            
            atrSignals.Add atrSignal
            previousTicker = currentTicker
        End If
    Next i
    
    ' Output ATR signals to dashboard or dedicated sheet
    Call OutputATRSignals(atrSignals, wsDash)
End Sub

' ATR SIGNAL CLASS MODULE
' (Create a new Class Module named "ATRSignal")
'
' Copy this code into the ATRSignal class module:

'Private pTicker As String
'Private pPrice As Double
'Private pATR As Double
'Private pATRPercent As Double
'Private pATRRatio As Double
'Private pVolatilityZone As String
'Private pRawSignal As String
'Private pTradingSignal As String
'Private pStopLoss As Double
'Private pPositionSize As Double
'
'Public Property Get Ticker() As String
'    Ticker = pTicker
'End Property
'Public Property Let Ticker(value As String)
'    pTicker = value
'End Property
'
'Public Property Get Price() As Double
'    Price = pPrice
'End Property
'Public Property Let Price(value As Double)
'    pPrice = value
'End Property
'
'Public Property Get ATR() As Double
'    ATR = pATR
'End Property
'Public Property Let ATR(value As Double)
'    pATR = value
'End Property
'
'Public Property Get ATRPercent() As Double
'    ATRPercent = pATRPercent
'End Property
'Public Property Let ATRPercent(value As Double)
'    pATRPercent = value
'End Property
'
'Public Property Get ATRRatio() As Double
'    ATRRatio = pATRRatio
'End Property
'Public Property Let ATRRatio(value As Double)
'    pATRRatio = value
'End Property
'
'Public Property Get VolatilityZone() As String
'    VolatilityZone = pVolatilityZone
'End Property
'Public Property Let VolatilityZone(value As String)
'    pVolatilityZone = value
'End Property
'
'Public Property Get RawSignal() As String
'    RawSignal = pRawSignal
'End Property
'Public Property Let RawSignal(value As String)
'    pRawSignal = value
'End Property
'
'Public Property Get TradingSignal() As String
'    TradingSignal = pTradingSignal
'End Property
'
'Public Property Get StopLoss() As Double
'    StopLoss = pStopLoss
'End Property
'Public Property Let StopLoss(value As Double)
'    pStopLoss = value
'End Property
'
'Public Property Get PositionSize() As Double
'    PositionSize = pPositionSize
'End Property
'Public Property Let PositionSize(value As Double)
'    pPositionSize = value
'End Property
'
'Public Sub GenerateTradingSignal()
'    ' ATR-based trading signals
'    If pATRPercent > 5 Then
'        pTradingSignal = "AVOID - Extreme Volatility"
'    ElseIf pRawSignal = "VOLATILITY_SPIKE" And pATRRatio > 1.8 Then
'        pTradingSignal = "BREAKOUT_WATCH"
'    ElseIf pRawSignal = "VOLATILITY_CONTRACTION" And pATRPercent < 1.5 Then
'        pTradingSignal = "CONSOLIDATION - Wait for Breakout"
'    ElseIf pATRPercent < 2 And pATRRatio < 1.2 Then
'        pTradingSignal = "LOW_VOL - Good for Swing"
'    ElseIf pATRPercent >= 2 And pATRPercent <= 3.5 Then
'        pTradingSignal = "NORMAL_VOL - Standard Trading"
'    Else
'        pTradingSignal = "MONITOR - Neutral Volatility"
'    End If
'End Sub

Sub OutputATRSignals(atrSignals As Collection, wsDash As Worksheet)
    Dim outputSheet As Worksheet
    Dim i As Long, outputRow As Long
    
    ' Create or clear ATR signals sheet
    On Error Resume Next
    Set outputSheet = ThisWorkbook.Sheets("ATR Signals")
    If outputSheet Is Nothing Then
        Set outputSheet = ThisWorkbook.Sheets.Add
        outputSheet.Name = "ATR Signals"
    End If
    On Error GoTo 0
    
    outputSheet.Cells.Clear
    outputRow = 1
    
    ' Setup headers
    With outputSheet
        .Cells(outputRow, 1).value = "Ticker"
        .Cells(outputRow, 2).value = "Price"
        .Cells(outputRow, 3).value = "ATR"
        .Cells(outputRow, 4).value = "ATR %"
        .Cells(outputRow, 5).value = "Volatility Zone"
        .Cells(outputRow, 6).value = "Trading Signal"
        .Cells(outputRow, 7).value = "Stop Loss"
        .Cells(outputRow, 8).value = "Position Size %"
        .Cells(outputRow, 9).value = "Risk per Share"
        
        .Range("A1:I1").Font.Bold = True
        .Range("A1:I1").Interior.Color = RGB(200, 200, 200)
    End With
    
    outputRow = 2
    
    ' Output signals
    For i = 1 To atrSignals.count
        Dim signal As atrSignal
        Set signal = atrSignals(i)
        
        With outputSheet
            .Cells(outputRow, 1).value = signal.ticker
            .Cells(outputRow, 2).value = signal.Price
            .Cells(outputRow, 3).value = Round(signal.ATR, 4)
            .Cells(outputRow, 4).value = Round(signal.atrPercent, 2)
            .Cells(outputRow, 5).value = signal.volatilityZone
            .Cells(outputRow, 6).value = signal.TradingSignal
            .Cells(outputRow, 7).value = Round(signal.StopLoss, 2)
            .Cells(outputRow, 8).value = signal.PositionSize
            .Cells(outputRow, 9).value = Round(signal.Price - signal.StopLoss, 2)
        End With
        
        ' Apply conditional formatting based on volatility zone
        With outputSheet.Range("E" & outputRow)
            Select Case signal.volatilityZone
                Case "LOW VOL"
                    .Interior.Color = RGB(198, 239, 206) ' Green
                Case "HIGH VOL"
                    .Interior.Color = RGB(255, 235, 156) ' Yellow
                Case "EXTREME VOL"
                    .Interior.Color = RGB(255, 199, 206) ' Red
            End Select
        End With
        
        ' Apply conditional formatting based on trading signal
        With outputSheet.Range("F" & outputRow)
            If InStr(signal.TradingSignal, "AVOID") > 0 Then
                .Font.Color = RGB(255, 0, 0)
                .Font.Bold = True
            ElseIf InStr(signal.TradingSignal, "BREAKOUT") > 0 Then
                .Font.Color = RGB(0, 128, 0)
                .Font.Bold = True
            ElseIf InStr(signal.TradingSignal, "LOW_VOL") > 0 Then
                .Font.Color = RGB(0, 0, 255)
            End If
        End With
        
        outputRow = outputRow + 1
    Next i
    
    ' Auto-fit columns
    outputSheet.Columns.AutoFit
    
    ' Add summary statistics
    Call AddATRSummary(outputSheet, outputRow, atrSignals)
End Sub

Sub AddATRSummary(ws As Worksheet, startRow As Long, signals As Collection)
    Dim i As Long
    Dim lowVolCount As Long, normalVolCount As Long, highVolCount As Long, extremeVolCount As Long
    
    For i = 1 To signals.count
        Dim signal As atrSignal
        Set signal = signals(i)
        
        Select Case signal.volatilityZone
            Case "LOW VOL": lowVolCount = lowVolCount + 1
            Case "NORMAL VOL": normalVolCount = normalVolCount + 1
            Case "HIGH VOL": highVolCount = highVolCount + 1
            Case "EXTREME VOL": extremeVolCount = extremeVolCount + 1
        End Select
    Next i
    
    With ws
        .Cells(startRow + 2, 1).value = "ATR SIGNALS SUMMARY"
        .Cells(startRow + 2, 1).Font.Bold = True
        .Cells(startRow + 2, 1).Font.Size = 12
        
        .Cells(startRow + 3, 1).value = "Low Volatility:"
        .Cells(startRow + 3, 2).value = lowVolCount
        .Cells(startRow + 3, 2).Interior.Color = RGB(198, 239, 206)
        
        .Cells(startRow + 4, 1).value = "Normal Volatility:"
        .Cells(startRow + 4, 2).value = normalVolCount
        
        .Cells(startRow + 5, 1).value = "High Volatility:"
        .Cells(startRow + 5, 2).value = highVolCount
        .Cells(startRow + 5, 2).Interior.Color = RGB(255, 235, 156)
        
        .Cells(startRow + 6, 1).value = "Extreme Volatility:"
        .Cells(startRow + 6, 2).value = extremeVolCount
        .Cells(startRow + 6, 2).Interior.Color = RGB(255, 199, 206)
        
        .Cells(startRow + 7, 1).value = "Total Securities:"
        .Cells(startRow + 7, 2).value = signals.count
        .Cells(startRow + 7, 2).Font.Bold = True
    End With
End Sub

' HELPER FUNCTION
Function Nz(value As Variant, defaultVal As Double) As Double
    If IsEmpty(value) Or IsError(value) Or value = "" Then
        Nz = defaultVal
    Else
        Nz = CDbl(value)
    End If
End Function

' INTEGRATION WITH MAIN SYSTEM
Sub UpdateSystemWithATR()
    ' Call this from your main FilterAndReport procedure
    Call CalculateATRWithSignals
End Sub
