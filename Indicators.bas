Attribute VB_Name = "Indicators"
Option Explicit
Private Const COL_EMA8 As String = "F"
Private Const COL_EMA21 As String = "G"
Private Const COL_RSI As String = "H"

Private Sub CleanupTempColumns()
    On Error Resume Next
    ws.Range("AB:AG").Clear
    On Error GoTo 0
End Sub

Sub prepData()

Dim lastRow As Long, period As Integer
Dim endDate As Date
Dim startDate As Date
Dim visibleRange As Range
Dim ticker As String
Dim wsFrom As Worksheet, wsTo As Worksheet, wsDash As Worksheet
Dim startTime As Double
Dim answer As VbMsgBoxResult
    
    startTime = Timer ' Start timer to calculate duration
    gStopMacro = False  ' Reset at the start of the macro
     
    Set wsDash = ThisWorkbook.Sheets("DashBoard")
    Set wsFrom = ThisWorkbook.Sheets("BackupAll")
    
    ticker = wsDash.Range("AF8")
    startDate = Format(GetPreviousWorkday(wsDash.Range("AAC2")), "yyyy-mm-dd")
    endDate = Format(GetPreviousWorkday(Date), "yyyy-mm-dd")
    
    ' Create sheet for the ticker if it doesn't exist
    On Error Resume Next
    Set wsTo = ThisWorkbook.Sheets(ticker)
    If wsTo Is Nothing Then
        Set wsTo = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("TechnicalAnalysis"))
        wsTo.Name = ticker
    End If
    On Error GoTo 0
    
    ' Clear existing data
    wsTo.Cells.Clear
            
    wsFrom.AutoFilterMode = False
    lastRow = wsFrom.Cells(wsFrom.Rows.count, "A").End(xlUp).row
          
    Set visibleRange = wsFrom.Range("A1:G" & lastRow)
    visibleRange.AutoFilter Field:=1, Criteria1:=">=" & startDate, Operator:=xlAnd, Criteria2:="<=" & endDate
    visibleRange.AutoFilter Field:=7, Criteria1:=ticker
    visibleRange.Copy wsTo.Range("A1")
    
    Application.CutCopyMode = False
    wsFrom.AutoFilterMode = False
      
    'Application.ScreenUpdating = True
    
End Sub

Sub CalculateIndicators()
    Dim wsTA As Worksheet
    Dim ws As Worksheet
    Dim selectedSheetName As String
         
     If MsgBox("Calculate Indicators ~ 20 minutes? ", vbYesNo) = vbNo Then End
     
    ' Create sheet for the ticker if it doesn't exist
     selectedSheetName = ThisWorkbook.Worksheets("DashBoard").Range("AF8").value
     
     On Error Resume Next
    Set ws = ThisWorkbook.Sheets(selectedSheetName)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("TechnicalAnalysis"))
        ws.Name = ticker
        prepData
    End If
    On Error GoTo 0
    Set ws = ThisWorkbook.Sheets(selectedSheetName)
    
    ws.Range("B:E").NumberFormat = "#,##0.00"
    
    gStopMacro = False  ' Reset at the start of the macro
    If gStopMacro Then
            MsgBox "Macro stopped by user.", vbInformation
            End
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
   
    Set wsTA = ThisWorkbook.Worksheets("TechnicalAnalysis")
    wsTA.Cells.Clear
    
    ' Copy date and price data
    ws.Range("A1:E" & lastRow).Copy wsTA.Range("A1")
    Application.CutCopyMode = False
    
    ' Calculate EMAs
    Call CalculateEMA(wsTA, lastRow, 8, "F")  ' 8-day EMA
    Call CalculateEMA(wsTA, lastRow, 21, "G") ' 21-day EMA
  
    ' Calculate RSI
    Call CalculateRSI(wsTA, lastRow, 14, "H")
    
    ' Calculate MACD
    Call CalculateMACD(wsTA, lastRow, "I", "J", "K")
    
    ' Calculate Bollinger Bands
    Call CalculateBollingerBands(wsTA, lastRow, 20, "L", "M", "N")
      
    ' Calculate SMI
    Call Calculate_SMI(wsTA, 2)
    
    ' Calculate Bressert DSS
    Call Calculate_Bressert_DSS(wsTA, 2)
    
    ' Calculate Volume spike, Elder-Ray, ATR and Keltner
    Call CalculateComplementaryIndicators
    
    wsTA.Range("B:BB").NumberFormat = "#,##0.00_ ;[Red]-#,##0.00 "
    
End Sub

Sub CalculateComplementaryIndicators()
    Dim wsTA As Worksheet
    Dim ws As Worksheet
    Dim selectedSheetName As String
    Dim dataRange As Range
    Dim lastRow As Long
    
    'Set ws = ThisWorkbook.ActiveSheet
    Set wsTA = ThisWorkbook.Worksheets("TechnicalAnalysis")
    
    selectedSheetName = ThisWorkbook.Worksheets("DashBoard").Range("AF8").value ' Ensure the selected sheet exists
    Set ws = ThisWorkbook.Sheets(selectedSheetName)
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    Set dataRange = ws.Range("A2:F" & lastRow)
    
    ' Calculate ATR
    Call CalculateATR(wsTA, lastRow, 14, "W")
    
    ' Calculate Elder-Ray
    Call CalculateElderRay(wsTA, lastRow, 13, "X", "Y")
    
    ' Calculate Volume Profile
        'Call CalculateVolumeProfile(wsTA, lastRow, "Z")
    Call CalculateVolumeProfile(Worksheets("TechnicalAnalysis"), lastRow, 26)
    
    ' Calculate Keltner Channels
    Call CalculateKeltnerChannels(wsTA, lastRow, 20, "AA", "AB", "AC")
    
    ' Calculate Directional Movement Index
    Call toCalculateDMI(dataRange, 14)
End Sub


Sub CalculateEMA(ws As Worksheet, lastRow As Long, period As Integer, col As String)
    Dim i As Long
    Dim multiplier As Double
    multiplier = 2 / (period + 1)
    
    ' First EMA is SMA
    ws.Range(col & "1").value = period & "-EMA"
    ws.Range(col & period).value = Application.Average(ws.Range("E2:E" & period + 1))
    
    ' Calculate subsequent EMAs
    For i = period + 1 To lastRow
        ws.Range(col & i).value = (ws.Range("E" & i).value - ws.Range(col & (i - 1)).value) _
            * multiplier + ws.Range(col & (i - 1)).value
    Next i
End Sub

Sub CalculateRSI(ws As Worksheet, lastRow As Long, period As Integer, col As String)
    Dim i As Long
    ws.Range(col & "1").value = "RSI"
    
    ' Calculate price changes
    For i = 3 To lastRow
        Dim change As Double
        change = ws.Range("E" & i).value - ws.Range("E" & (i - 1)).value
        ws.Range("AB" & i).value = IIf(change > 0, change, 0)
        ws.Range("AC" & i).value = IIf(change < 0, -change, 0)
    Next i
    
    ' Calculate first RSI
    Dim avgGain As Double, avgLoss As Double
    avgGain = CDbl(Application.Average(ws.Range("AB3:AB" & (period + 2))))
    avgLoss = CDbl(Application.Average(ws.Range("AC3:AC" & (period + 2))))
    
    ws.Range(col & (period + 2)).value = 100 - (100 / (1 + (avgGain / avgLoss)))
    
    ' Calculate subsequent RSIs
    For i = period + 3 To lastRow
        avgGain = (ws.Range("AB" & i).value + (period - 1) * avgGain) / period
        avgLoss = (ws.Range("AC" & i).value + (period - 1) * avgLoss) / period
        ws.Range(col & i).value = 100 - (100 / (1 + (avgGain / avgLoss)))
    Next i
    
    ' Clean up temporary columns
    ws.Range("AB:AC").Clear
End Sub

Sub CalculateMACD(ws As Worksheet, lastRow As Long, macdCol As String, signalCol As String, histCol As String)
    ' Calculate 12 and 26 day EMAs
    Call CalculateEMA(ws, lastRow, 12, "AD")
    Call CalculateEMA(ws, lastRow, 26, "AE")
    
    ' Calculate MACD Line
    ws.Range(macdCol & "1").value = "MACD"
    ws.Range(signalCol & "1").value = "Signal"
    ws.Range(histCol & "1").value = "Histogram"
    
    Dim i As Long
    For i = 27 To lastRow
        ws.Range(macdCol & i).value = ws.Range("AD" & i).value - ws.Range("AE" & i).value  ' 12-EMA minus 26-EMA
    Next i
    
    ' Calculate 9-day EMA of MACD for Signal line
    Dim multiplier As Double
    multiplier = 2 / (9 + 1)
    
    ws.Range(signalCol & "35").value = Application.Average(ws.Range(macdCol & "27:" & macdCol & "35"))
    
    For i = 36 To lastRow
        ws.Range(signalCol & i).value = (ws.Range(macdCol & i).value - ws.Range(signalCol & (i - 1)).value) _
            * multiplier + ws.Range(signalCol & (i - 1)).value
            
        ' Calculate Histogram
        ws.Range(histCol & i).value = ws.Range(macdCol & i).value - ws.Range(signalCol & i).value
    Next i
    
    ' Clean up temporary columns
    ws.Range("AD:AE").Clear
End Sub

Sub CalculateBollingerBands(ws As Worksheet, lastRow As Long, period As Integer, upperCol As String, middleCol As String, lowerCol As String)
    ws.Range(upperCol & "1").value = "Upper BB"
    ws.Range(middleCol & "1").value = "Middle BB"
    ws.Range(lowerCol & "1").value = "Lower BB"
    
    Dim i As Long
    For i = period To lastRow
        Dim rng As Range
        Set rng = ws.Range("E" & (i - period + 1) & ":E" & i)
        
        Dim sma As Double
        sma = Application.Average(rng)
        
        Dim stdDev As Double
        stdDev = Application.StDev(rng)
        
        ws.Range(middleCol & i).value = sma
        ws.Range(upperCol & i).value = sma + (2 * stdDev)
        ws.Range(lowerCol & i).value = sma - (2 * stdDev)
    Next i
End Sub

Public Function SMI_EMA(ByVal values As Range, ByVal period As Long) As Double
    Dim multiplier As Double
    Dim i As Long
    Dim ema As Double
    Dim firstSum As Double
    
    multiplier = 2 / (period + 1)
    
    ' Calculate first SMA
    firstSum = 0
    For i = 1 To period
        firstSum = firstSum + CDbl(values.Cells(i, 1).value)
    Next i
    ema = firstSum / period
    
    ' Calculate EMA
    For i = period + 1 To values.Rows.count
        ema = (CDbl(values.Cells(i, 1).value) - ema) * multiplier + ema
    Next i
    
    SMI_EMA = ema
End Function

Public Function Calculate_SMI(ByRef ws As Worksheet, ByVal startRow As Long, _
    Optional ByVal period As Long = 10, _
    Optional ByVal smoothK As Long = 3, _
    Optional ByVal smoothD As Long = 3, _
    Optional ByVal overbought As Double = 60, _
    Optional ByVal oversold As Double = -60) As Boolean
    
    Dim lastRow As Long
    Dim i As Long
    Dim highest As Double, lowest As Double
    Dim CM As Double, HL As Double
    Dim DS As Double, DHL As Double
    Dim prevSMI As Double, currentSMI As Double
    
    ' Add headers
    ws.Cells(1, 15).value = "SMI_DS"
    ws.Cells(1, 16).value = "SMI_DHL"
    ws.Cells(1, 17).value = "SMI"
    ws.Cells(1, 18).value = "SMI_Signal"
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    For i = startRow + period To lastRow
        ' Calculate Distance from Median
        highest = CDbl(Application.max(ws.Range(ws.Cells(i - period + 1, "C"), ws.Cells(i, "C"))))
        lowest = CDbl(Application.min(ws.Range(ws.Cells(i - period + 1, "D"), ws.Cells(i, "D"))))
        CM = CDbl(ws.Cells(i, "E").value) - (highest + lowest) / 2
        HL = highest - lowest
        
        ' Store intermediate values
        ws.Cells(i, "O").value = CM
        ws.Cells(i, "P").value = HL
        
        ' Calculate smoothed values using EMA
        If i >= startRow + period + smoothK Then
            DS = SMI_EMA(ws.Range(ws.Cells(i - smoothK + 1, "O"), ws.Cells(i, "O")), smoothK)
            DHL = SMI_EMA(ws.Range(ws.Cells(i - smoothK + 1, "P"), ws.Cells(i, "P")), smoothK)
            
            ' Calculate final SMI
                If DHL <> 0 Then
                    currentSMI = 100 * (DS / (DHL / 2))
                    ws.Cells(i, "Q").value = currentSMI
                    
                    ' Calculate signals
                        If i > startRow + period + smoothK Then
                            prevSMI = ws.Cells(i - 1, "Q").value
                            
                            ' Signal logic
                                If currentSMI > overbought And prevSMI <= overbought Then
                                    ws.Cells(i, "R").value = "Sell"
                                ElseIf currentSMI < oversold And prevSMI >= oversold Then
                                    ws.Cells(i, "R").value = "Buy"
                                ElseIf (currentSMI > 0 And prevSMI <= 0) Or (currentSMI < 0 And prevSMI >= 0) Then
                                    ws.Cells(i, "R").value = ""
                                Else
                                    ws.Cells(i, "R").value = ""
                            End If
                    End If
            Else
                ws.Cells(i, "Q").value = 0
                ws.Cells(i, "R").value = ""
            End If
        End If
    Next i
    
    Calculate_SMI = True
End Function

Public Function Calculate_Bressert_DSS(ByRef ws As Worksheet, ByVal startRow As Long, _
    Optional ByVal period As Long = 13, _
    Optional ByVal smoothPeriod As Long = 8, _
    Optional ByVal overbought As Double = 80, _
    Optional ByVal oversold As Double = 20) As Boolean
    
    Dim lastRow As Long
    Dim i As Long
    Dim stochK As Double
    Dim highest As Double, lowest As Double
    Dim prevprevDSS As Double, prevDSS As Double, currentDSS As Double
    
    ' Add headers
    ws.Cells(1, 19).value = "Stoch_K"
    ws.Cells(1, 20).value = "First_Smooth"
    ws.Cells(1, 21).value = "DSS"
    ws.Cells(1, 22).value = "DSS_Signal"
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    For i = startRow + period To lastRow
        ' Calculate Stochastic K
        highest = CDbl(Application.max(ws.Range(ws.Cells(i - period + 1, "C"), ws.Cells(i, "C"))))
        lowest = CDbl(Application.min(ws.Range(ws.Cells(i - period + 1, "D"), ws.Cells(i, "D"))))
        
        If highest <> lowest Then
            stochK = 100 * (CDbl(ws.Cells(i, "E").value) - lowest) / (highest - lowest)
        Else
            stochK = 50
        End If
        
        ws.Cells(i, "S").value = stochK
        
        ' Calculate double smoothed stochastic using EMA
        If i >= startRow + period + smoothPeriod Then
            ' First smoothing
            Dim firstSmooth As Double
            firstSmooth = SMI_EMA(ws.Range(ws.Cells(i - smoothPeriod + 1, "S"), ws.Cells(i, "S")), smoothPeriod)
            
            ' Store first smoothing result for second smoothing
            ws.Cells(i, "T").value = firstSmooth
            
            ' Second smoothing
            If i >= startRow + period + (2 * smoothPeriod) Then
                currentDSS = SMI_EMA(ws.Range(ws.Cells(i - smoothPeriod + 1, "T"), ws.Cells(i, "T")), smoothPeriod)
                ws.Cells(i, "U").value = currentDSS
                
                ' Calculate signals
                If i > startRow + period + (2 * smoothPeriod) Then
                    prevDSS = ws.Cells(i - 1, "S").value
                    prevprevDSS = ws.Cells(i - 2, "S").value
                    
                    ' Signal logic
                    If currentDSS > overbought And prevDSS <= overbought And prevprevDSS <= overbought Then
                        ws.Cells(i, "V").value = "Sell"
                    ElseIf currentDSS < oversold And prevDSS >= oversold And prevprevDSS >= oversold Then
                        ws.Cells(i, "V").value = "Buy"
                    ElseIf (currentDSS > 50 And prevDSS <= 50) Or (currentDSS < 50 And prevDSS >= 50) Then
                        ws.Cells(i, "V").value = ""
                    Else
                        ws.Cells(i, "V").value = ""
                    End If
                End If
            End If
        End If
    Next i
    
    Calculate_Bressert_DSS = True
End Function

Sub otherAddTechnicalOverlay(ws As Worksheet, startCol As Long, endCol As Long, lastRow As Long)
    ' Add signal legend
    ws.Cells(8, 9).value = "Technical Signals:"
    ws.Cells(9, 9).value = "? EMA Cross (8/21)"
    ws.Cells(10, 9).value = "? RSI Extremes (30/70)"
    ws.Cells(11, 9).value = "? MACD Cross"
    ws.Cells(12, 9).value = "Yellow highlight = Signal"
End Sub

Sub FormatChart(ws As Worksheet, lastCol As Long, startDate As String, endDate As String)
    ' Format existing P&F chart
    With ws.Range(ws.Cells(1, 11), ws.Cells(100, lastCol))
        .Font.Name = "Consolas"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ColumnWidth = 2
        .RowHeight = 15
    End With
    
    ' Add date information
    ws.Cells(13, 9).value = "Date Range:"
    ws.Cells(14, 9).value = Format(startDate, "mm/dd/yyyy") & " - " & Format(endDate, "mm/dd/yyyy")
End Sub

Sub CalculateATR(ws As Worksheet, lastRow As Long, period As Integer, col As String)
    ws.Cells(1, col).value = "ATR"
    
    Dim i As Long
    For i = 2 To lastRow
        ' Calculate True Range
        Dim TR As Double
        Dim highLow As Double
        Dim highClose As Double
        Dim lowClose As Double
        
        highLow = Abs(ws.Cells(i, "C").value - ws.Cells(i, "D").value)
        If i > 2 Then
            highClose = Abs(ws.Cells(i, "C").value - ws.Cells(i - 1, "E").value)
            lowClose = Abs(ws.Cells(i, "D").value - ws.Cells(i - 1, "E").value)
            TR = WorksheetFunction.max(highLow, highClose, lowClose)
        Else
            TR = highLow
        End If
        
        ' Calculate ATR
        If i = 2 Then
            ws.Cells(i, col).value = TR
        Else
            ws.Cells(i, col).value = ((period - 1) * ws.Cells(i - 1, col).value + TR) / period
        End If
    Next i
End Sub

Sub CalculateElderRay(ws As Worksheet, lastRow As Long, period As Integer, bullCol As String, bearCol As String)
    ws.Cells(1, bullCol).value = "Bull Power"
    ws.Cells(1, bearCol).value = "Bear Power"
    
    ' Calculate 13-day EMA first
    Call CalculateEMA(ws, lastRow, period, "AH")  ' Temporary column
    
    Dim i As Long
    For i = period + 1 To lastRow
        ' Bull Power = High - EMA
        ws.Cells(i, bullCol).value = ws.Cells(i, "C").value - ws.Cells(i, "AH").value
        
        ' Bear Power = Low - EMA
        ws.Cells(i, bearCol).value = ws.Cells(i, "D").value - ws.Cells(i, "AH").value
    Next i
    
    ' Clean up temporary column
    ws.Range("AH:AH").Clear
End Sub

Sub CalculateVolumeProfile(ws As Worksheet, lastRow As Long, col As Long)
    
    'ws.Range("Z1:Z" & lastrow).ClearContents
    ws.Cells(1, col).value = "Volume Profile"
  
    Dim i As Long
    Dim priceRange As Double
    Dim numLevels As Long
    Dim wsData As Worksheet
    Dim selectedSheetName As String
    
     ' Assuming volume is in column F
    
    selectedSheetName = ThisWorkbook.Worksheets("DashBoard").Range("AF8").value
    
    Set wsData = ThisWorkbook.Sheets(selectedSheetName)
    
    numLevels = 10  ' Number of price levels to analyze
    
    ' Find price range
    Dim highestPrice As Double
    Dim lowestPrice As Double
    highestPrice = WorksheetFunction.max(ws.Range("C2:C" & lastRow))
    lowestPrice = WorksheetFunction.min(ws.Range("D2:D" & lastRow))
    priceRange = highestPrice - lowestPrice
    
    ' Calculate price levels
    Dim levelSize As Double
    levelSize = priceRange / numLevels
    
    ' Initialize arrays for volume at each level
    Dim volumeLevels() As Long
    ReDim volumeLevels(1 To numLevels)
    
    ' Calculate volume distribution
    For i = 2 To lastRow
        Dim price As Double
        Dim level As Long
        price = ws.Cells(i, "E").value  ' Using close price
        
        ' Prevent division by zero and handle edge cases
        If levelSize > 0 Then
            level = Int((price - lowestPrice) / levelSize) + 1
            
            ' Ensure level is within array bounds
            If level < 1 Then level = 1
            If level > numLevels Then level = numLevels
            
            volumeLevels(level) = volumeLevels(level) + CLng(wsData.Cells(i, "F").value)
        End If
    Next i
    
    ' Find POC (Point of Control) - level with highest volume
    Dim maxVolume As Long
    Dim pocLevel As Long
    maxVolume = 0
    For i = 1 To numLevels
        If volumeLevels(i) > maxVolume Then
            maxVolume = volumeLevels(i)
            pocLevel = i
        End If
    Next i
    
    ' Mark significant volume levels
    For i = 2 To lastRow
        Dim currentLevel As Long
        price = ws.Cells(i, "E").value
        
        If levelSize > 0 Then
            currentLevel = Int((price - lowestPrice) / levelSize) + 1
            
            ' Ensure currentLevel is within bounds
            If currentLevel < 1 Then currentLevel = 1
            If currentLevel > numLevels Then currentLevel = numLevels
            
            ' Mark POC and high volume areas
            If currentLevel = pocLevel Then
                ws.Cells(i, col).value = "POC"
            ElseIf volumeLevels(currentLevel) > maxVolume * 0.7 Then
                ws.Cells(i, col).value = "HVN"  ' High Volume Node
            ElseIf volumeLevels(currentLevel) < maxVolume * 0.3 Then
                ws.Cells(i, col).value = "LVN"  ' Low Volume Node
            End If
        End If
    Next i
End Sub

Sub CalculateKeltnerChannels(ws As Worksheet, lastRow As Long, period As Integer, upperCol As String, middleCol As String, lowerCol As String)
    ws.Cells(1, upperCol).value = "Keltner Upper"
    ws.Cells(1, middleCol).value = "Keltner Middle"
    ws.Cells(1, lowerCol).value = "Keltner Lower"
    
    ' Calculate 20-day EMA for middle line
    Call CalculateEMA(ws, lastRow, period, middleCol)
    
    ' Calculate ATR if not already calculated
    'Call CalculateATR(ws, lastRow, period, "AI")
    
    Dim i As Long
    For i = period + 1 To lastRow
        ' Upper = EMA + (2 * ATR)
        ws.Cells(i, upperCol).value = ws.Cells(i, middleCol).value + (2 * ws.Cells(i, "W").value)
        
        ' Lower = EMA - (2 * ATR)
        ws.Cells(i, lowerCol).value = ws.Cells(i, middleCol).value - (2 * ws.Cells(i, "W").value)
    Next i
    
    ' Clean up temporary ATR column
   ' ws.Range("AI:AI").Clear
End Sub

Function IsHighProbabilitySignal(row As Long) As Boolean
    Dim wsTA As Worksheet
    Set wsTA = ThisWorkbook.Worksheets("TechnicalAnalysis")
    
    ' Get all indicator values
    Dim ATR As Double
    Dim bullPower As Double
    Dim bearPower As Double
    Dim VolumeProfile As String
    Dim keltnerUpper As Double
    Dim keltnerLower As Double
    Dim price As Double
    
    ATR = wsTA.Cells(row, "W").value
    bullPower = wsTA.Cells(row, "X").value
    bearPower = wsTA.Cells(row, "Y").value
    VolumeProfile = wsTA.Cells(row, "Z").value
    keltnerUpper = wsTA.Cells(row, "AA").value
    keltnerLower = wsTA.Cells(row, "AC").value
    price = wsTA.Cells(row, "E").value
    
    ' Check if it's a regular signal first
    If Not IsSignal(row) Then
        IsHighProbabilitySignal = False
        Exit Function
    End If
    
    ' Additional confirmation criteria
    Dim volumeConfirmation As Boolean
    Dim trendConfirmation As Boolean
    Dim volatilityConfirmation As Boolean
    
    ' Volume Profile confirmation
    volumeConfirmation = (VolumeProfile = "POC" Or VolumeProfile = "HVN")
    
    ' Trend confirmation using Elder-Ray
    trendConfirmation = (bullPower > 0 And bearPower < 0) Or (bullPower < 0 And bearPower > 0)
    
    ' Volatility confirmation using ATR and Keltner Channels
    Dim avgATR As Double
    avgATR = CDbl(Application.Average(wsTA.Range("W" & row - 5 & ":W" & row)))
    volatilityConfirmation = (ATR > avgATR) And _
                            (price > keltnerLower And price < keltnerUpper)
    
    ' Combined probability assessment
    IsHighProbabilitySignal = volumeConfirmation And _
                             trendConfirmation And _
                             volatilityConfirmation
End Function

Sub UpdateSignalChecking()
    ' Update the existing IsSignal function call in your main code to:
    If IsHighProbabilitySignal(row) Then
        ' Your existing signal handling code
    End If
End Sub


' Function to calculate True Range with explicit type handling
Private Function TrueRange(ByVal high1 As Double, ByVal low1 As Double, ByVal close0 As Double, ByVal close1 As Double) As Double
    Dim tr1 As Double, tr2 As Double, tr3 As Double
    
    tr1 = Abs(high1 - low1)
    tr2 = Abs(high1 - close0)
    tr3 = Abs(low1 - close0)
    
    TrueRange = Application.max(tr1, Application.max(tr2, tr3))
End Function

' Safe conversion function
Private Function SafeConvertToDouble(value As Variant) As Double
    On Error Resume Next
    SafeConvertToDouble = CDbl(value)
    If Err.Number <> 0 Then
        SafeConvertToDouble = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

' Function to calculate Directional Movement Index (DMI) and output to specific columns
Sub toCalculateDMI(dataRange As Range, period As Integer)
    Dim i As Long
    Dim dates(), highs(), lows(), closes() As Variant
    Dim trueRanges() As Double
    Dim plusDM(), minusDM() As Double
    Dim plusDI(), minusDI(), ADX() As Double
    Dim lastValidIndex As Long
    Dim ws As Worksheet
   
    'Set ws = ThisWorkbook.Sheets(selectedSheetName)
       
    ' Get the worksheet
    Set ws = ThisWorkbook.Worksheets("TechnicalAnalysis")
    'lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    ws.Range("AD:AH").ClearContents
    
    ' Validate input
    If period < 2 Then
        MsgBox "Error: Period must be at least 2"
        Exit Sub
    End If
    
    ' Determine last valid row
    For i = dataRange.Rows.count To 1 Step -1
        If Not IsEmpty(dataRange.Cells(i, 1)) Then
            lastValidIndex = i
            Exit For
        End If
    Next i
    
    ' Check if we have enough data
    If lastValidIndex <= period Then
        MsgBox "Error: Not enough data points"
        Exit Sub
    End If
    
    ' Resize arrays
    ReDim dates(1 To lastValidIndex)
    ReDim highs(1 To lastValidIndex)
    ReDim lows(1 To lastValidIndex)
    ReDim closes(1 To lastValidIndex)
    ReDim trueRanges(1 To lastValidIndex)
    ReDim plusDM(1 To lastValidIndex)
    ReDim minusDM(1 To lastValidIndex)
    ReDim plusDI(1 To lastValidIndex)
    ReDim minusDI(1 To lastValidIndex)
    ReDim ADX(1 To lastValidIndex)
    
    ' Extract data
    For i = 1 To lastValidIndex
        dates(i) = dataRange.Cells(i, 1).value
        highs(i) = SafeConvertToDouble(dataRange.Cells(i, 3).value)
        lows(i) = SafeConvertToDouble(dataRange.Cells(i, 4).value)
        closes(i) = SafeConvertToDouble(dataRange.Cells(i, 5).value)
    Next i
    
    ' Calculate initial true ranges and directional movement
    For i = 2 To lastValidIndex
        trueRanges(i) = TrueRange(highs(i), lows(i), closes(i - 1), closes(i))
        
        Dim upMove As Double, downMove As Double
        upMove = highs(i) - highs(i - 1)
        downMove = lows(i - 1) - lows(i)
        
        If upMove > downMove And upMove > 0 Then
            plusDM(i) = upMove
        Else
            plusDM(i) = 0
        End If
        
        If downMove > upMove And downMove > 0 Then
            minusDM(i) = downMove
        Else
            minusDM(i) = 0
        End If
    Next i
    
    ' Calculate smoothed values using Wilder's smoothing
    Dim smoothedTR As Double, smoothedPlusDM As Double, smoothedMinusDM As Double
    
    ' Initial smoothing
    For i = 2 To period
        smoothedTR = smoothedTR + trueRanges(i)
        smoothedPlusDM = smoothedPlusDM + plusDM(i)
        smoothedMinusDM = smoothedMinusDM + minusDM(i)
    Next i
    
    ' Calculate first DIs after period
    If smoothedTR > 0 Then
        plusDI(period) = 100 * smoothedPlusDM / smoothedTR
        minusDI(period) = 100 * smoothedMinusDM / smoothedTR
    End If
    
    ' Continue calculations for remaining periods
    For i = period + 1 To lastValidIndex
        ' Update smoothed values
        smoothedTR = smoothedTR - (smoothedTR / period) + trueRanges(i)
        smoothedPlusDM = smoothedPlusDM - (smoothedPlusDM / period) + plusDM(i)
        smoothedMinusDM = smoothedMinusDM - (smoothedMinusDM / period) + minusDM(i)
        
        ' Calculate DI values
        If smoothedTR > 0 Then
            plusDI(i) = 100 * smoothedPlusDM / smoothedTR
            minusDI(i) = 100 * smoothedMinusDM / smoothedTR
            
            ' Calculate ADX
            Dim DX As Double
            DX = 100 * Abs(plusDI(i) - minusDI(i)) / (plusDI(i) + minusDI(i))
            
            ' Smooth ADX
            If i = period + 1 Then
                ADX(i) = DX
            Else
                ADX(i) = ((period - 1) * ADX(i - 1) + DX) / period
            End If
        End If
    Next i
    
    ' Add headers
    With ws
        .Cells(1, 30).value = "+DI"
        .Cells(1, 31).value = "-DI"
        .Cells(1, 32).value = "ADX"
        .Cells(1, 33).value = "Signal"
        .Cells(1, 34).value = "Trend"
    End With
    
    ' Output all values
 
'For ' Output all values
For i = period + 1 To lastValidIndex  ' Changed from lastValidIndex to lastValidIndex - 1
    With ws
        .Cells(i, 30).value = Round(plusDI(i), 2)     ' +DI in column AD (30)
        .Cells(i, 31).value = Round(minusDI(i), 2)    ' -DI in column AE (31)
        .Cells(i, 32).value = Round(ADX(i), 2)        ' ADX in column AF (32)
        
        ' Crossover signal - needs current and previous values
        If plusDI(i) > minusDI(i) And _
           plusDI(i - 1) <= minusDI(i - 1) Then
            .Cells(i, 33).value = "BUY"               ' Signal in column AG (33)
        ElseIf plusDI(i) < minusDI(i) And _
               plusDI(i - 1) >= minusDI(i - 1) Then
            .Cells(i, 33).value = "SELL"
        Else
            .Cells(i, 33).value = ""
        End If
        
        ' Trend strength
        If ADX(i) > 25 Then
            .Cells(i, 34).value = "STRONG"            ' Trend in column AH (34)
        ElseIf ADX(i) < 20 Then
            .Cells(i, 34).value = ""
        Else
            .Cells(i, 34).value = ""
        End If
    End With
Next i

' Add the final row separately without needing previous values
With ws
    .Cells(lastValidIndex, 30).value = Round(plusDI(lastValidIndex), 2)
    .Cells(lastValidIndex, 31).value = Round(minusDI(lastValidIndex), 2)
    .Cells(lastValidIndex, 32).value = Round(ADX(lastValidIndex), 2)
    
    ' For the last row, we can only show current state without crossover
    If plusDI(lastValidIndex) > minusDI(lastValidIndex) Then
        .Cells(lastValidIndex, 33).value = ""
    Else
        .Cells(lastValidIndex, 33).value = ""
    End If
    
    ' Trend strength for last row
    If ADX(lastValidIndex) > 25 Then
        .Cells(lastValidIndex, 34).value = "STRONG"
    ElseIf ADX(lastValidIndex) < 20 Then
        .Cells(lastValidIndex, 34).value = ""
    Else
        .Cells(lastValidIndex, 34).value = ""
    End If
End With
End Sub

