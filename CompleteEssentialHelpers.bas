Attribute VB_Name = "CompleteEssentialHelpers"
Option Explicit

'*******6. ESSENTIAL HELPER FUNCTIONS*******
Sub SetupTradingSignalsSheet(ws As Worksheet)
    ws.Cells.Clear
    
    With ws
        .Range("A1").value = "Trading Signals - " & Format(Date, "yyyy-mm-dd")
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        
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

Public Function Nz(value As Variant, defaultVal As Double) As Double
    If IsEmpty(value) Or IsError(value) Or value = "" Then
        Nz = defaultVal
    Else
        Nz = CDbl(value)
    End If
End Function

Function GetMarketRegime(wsDash As Worksheet) As String
    GetMarketRegime = "NORMAL"
End Function

Function GetRegimeAdjustedThreshold(regime As String, baseThreshold As Double) As Double
    Select Case regime
        Case "HIGH_VOLATILITY"
            ' Require stronger signals in high volatility
            GetRegimeAdjustedThreshold = baseThreshold * 1.3
        Case "STRONG_TREND"
            ' Can use lower thresholds in strong trends
            GetRegimeAdjustedThreshold = baseThreshold * 0.8
        Case "RANGING"
            ' Require slightly stronger signals in ranging markets
            GetRegimeAdjustedThreshold = baseThreshold * 1.1
        Case Else
            GetRegimeAdjustedThreshold = baseThreshold
    End Select
End Function

Function HasVolumeConfirmation(batchData As Variant, rowIndex As Long, volumeCol As Long, signalScore As Double) As Boolean
       ' Default to true if volume data not available
    If volumeCol > UBound(batchData, 2) Or Not IsNumeric(batchData(rowIndex, volumeCol)) Then
        HasVolumeConfirmation = True
        Exit Function
    End If
    
    Dim volume As Double
    volume = CDbl(batchData(rowIndex, volumeCol))
    
    ' Simple volume confirmation logic
    ' In practice, you'd compare to average volume
    If signalScore > 0 Then
        ' Buy signals should have decent volume
        HasVolumeConfirmation = (volume > 50000) ' Adjust threshold as needed
    Else
        ' Sell signals - volume less critical but still helpful
        HasVolumeConfirmation = (volume > 30000) ' Adjust threshold as needed
    End If
End Function

Function GetRecentPerformance(ticker As String, analysisDate As Date, daysBack As Long) As Double
    ' Returns the price % change over the last daysBack trading days for the given ticker
    ' Looks up data from the BackupAll sheet (col1=Date, col5=Close, col7=Ticker)
    On Error GoTo Fail
    Dim wsBackup As Worksheet
    Set wsBackup = ThisWorkbook.Sheets("BackupAll")
    Dim lastRow As Long
    lastRow = wsBackup.Cells(wsBackup.Rows.count, 1).End(xlUp).row

    Dim latestClose As Double, earliestClose As Double
    Dim latestDate As Date, earliestDate As Date
    Dim latestFound As Boolean, earliestFound As Boolean
    Dim i As Long
    latestFound = False: earliestFound = False

    ' Scan backwards to find the two price points for this ticker
    For i = lastRow To 2 Step -1
        If CStr(wsBackup.Cells(i, 7).Value) = ticker Then
            Dim rowDate As Date
            rowDate = CDate(wsBackup.Cells(i, 1).Value)
            If rowDate <= analysisDate Then
                If Not latestFound Then
                    latestClose = CDbl(wsBackup.Cells(i, 5).Value)
                    latestDate = rowDate
                    latestFound = True
                ElseIf latestFound And DateDiff("d", rowDate, latestDate) >= daysBack Then
                    earliestClose = CDbl(wsBackup.Cells(i, 5).Value)
                    earliestFound = True
                    Exit For
                End If
            End If
        End If
    Next i

    If latestFound And earliestFound And earliestClose <> 0 Then
        GetRecentPerformance = (latestClose - earliestClose) / earliestClose
    Else
        GetRecentPerformance = 0  ' Neutral if data not available
    End If
    Exit Function
Fail:
    GetRecentPerformance = 0
End Function

Function IsFalsePositive(ticker As String, signalScore As Double, analysisDate As Date) As Boolean
        ' Avoid buying into strong downtrends
    If signalScore > 0 Then
        Dim recentPerformance As Double
        recentPerformance = GetRecentPerformance(ticker, analysisDate, 5) ' 5-day performance
        
        If recentPerformance < -0.08 Then ' 8% drop recently
            IsFalsePositive = True
            Exit Function
        End If
    End If
    
    ' Avoid selling into strong uptrends
    If signalScore < 0 Then
        recentPerformance = GetRecentPerformance(ticker, analysisDate, 5)
        
        If recentPerformance > 0.08 Then ' 8% gain recently
            IsFalsePositive = True
            Exit Function
        End If
    End If
    
    ' Add other false positive patterns here
    ' e.g., earnings announcements, gap ups/downs, etc.
    
    IsFalsePositive = False
End Function

Sub VerifyATRDataCompleteness(ws As Worksheet, expectedRows As Long)
    Dim lastCalcRow As Long
    lastCalcRow = ws.Cells(ws.Rows.count, "O").End(xlUp).row
    Debug.Print "ATR rows calculated: " & lastCalcRow & " of " & expectedRows
End Sub
'*** 6. ESSENTIAL HELPER FUNCTIONS END***

