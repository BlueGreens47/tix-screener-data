Attribute VB_Name = "ValidateData"
Sub DetectAnomalies()
    Dim ws As Worksheet
    Dim wsAnomalies As Worksheet
    Dim lastRow As Long
    Dim lastRowAnomalies As Long
    Dim i As Long
    Dim prevClose As Double
    Dim currentOpen As Double, currentHigh As Double, currentLow As Double, currentClose As Double
    Dim pctChangeOpen As Double, pctChangeHigh As Double, pctChangeLow As Double, pctChangeClose As Double
    Const PCT_THRESHOLD As Double = 1#  ' Example: 50% change threshold (adjust as needed, e.g., 1.0 for 100%)
    Const PCT_THRESHOLD_VOL As Double = 3#   ' Example: 300% change threshold for Volume

    Set ws = ThisWorkbook.Sheets("Test") ' Adjust main data sheet name as needed

    ' Check if "anomaliesList" sheet exists, if not, create it
    On Error Resume Next
    Set wsAnomalies = ThisWorkbook.Sheets("AnomaliesList")
    On Error GoTo 0

    If wsAnomalies Is Nothing Then
        Set wsAnomalies = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsAnomalies.Name = "AnomaliesList"
        ' Add headers to the new sheet
        wsAnomalies.Cells(1, 1).Value = "Date"
        wsAnomalies.Cells(1, 2).Value = "Ticker"
        wsAnomalies.Cells(1, 3).Value = "Open"
        wsAnomalies.Cells(1, 4).Value = "High"
        wsAnomalies.Cells(1, 5).Value = "Low"
        wsAnomalies.Cells(1, 6).Value = "Close"
        wsAnomalies.Cells(1, 7).Value = "Volume"
        wsAnomalies.Cells(1, 8).Value = "Anomaly Type"
        wsAnomalies.Cells(1, 9).Value = "Details"
        wsAnomalies.Rows(1).Font.Bold = True
    End If

    ' Find the last row with data in Column A (Date)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop from the second data row (row 3, assuming header in row 1, first data in row 2)
    For i = 3 To lastRow
    
    ' Jump windowSize and restart if not same Ticker
     If ws.Cells(i - 1, 7).Value <> ws.Cells(i, 7).Value Then i = i + 7 'then 1 week later
     
        ' Get previous day's closing price (from row i-1)
        prevClose = ws.Cells(i - 1, "E").Value ' Column E is 'Close'

        ' Get current day's prices
        currentOpen = ws.Cells(i, "B").Value ' Column B is 'Open'
        currentHigh = ws.Cells(i, "C").Value ' Column C is 'High'
        currentLow = ws.Cells(i, "D").Value ' Column D is 'Low'
        currentClose = ws.Cells(i, "E").Value ' Column E is 'Close'

        ' Calculate percentage changes relative to previous day's close
        If prevClose <> 0 Then ' Avoid division by zero
            pctChangeOpen = Abs((currentOpen - prevClose) / prevClose)
            pctChangeHigh = Abs((currentHigh - prevClose) / prevClose)
            pctChangeLow = Abs((currentLow - prevClose) / prevClose)
            pctChangeClose = Abs((currentClose - prevClose) / prevClose)
        Else
            ' Handle cases where previous close is zero (e.g., new listing, or extreme anomaly)
            ' If any current price is non-zero and prevClose is zero, it's a significant change
            If currentOpen <> 0 Or currentHigh <> 0 Or currentLow <> 0 Or currentClose <> 0 Then
                pctChangeOpen = 999 ' Assign a very high value to trigger anomaly
                pctChangeHigh = 999
                pctChangeLow = 999
                pctChangeClose = 999
            Else
                pctChangeOpen = 0 ' No change if both are zero
                pctChangeHigh = 0
                pctChangeLow = 0
                pctChangeClose = 0
            End If
        End If

        Dim isAnomaly As Boolean
        isAnomaly = False
        Dim anomalyDetails As String

        ' Check for anomaly based on price change threshold
        If pctChangeOpen > PCT_THRESHOLD Or _
           pctChangeHigh > PCT_THRESHOLD Or _
           pctChangeLow > PCT_THRESHOLD Or _
           pctChangeClose > PCT_THRESHOLD Then
            isAnomaly = True
            anomalyDetails = "Anomaly: Price change > " & Format(PCT_THRESHOLD, "0%") & ". Open: " & Format(pctChangeOpen, "0.0%") & ", High: " & Format(pctChangeHigh, "0.0%") & ", Low: " & Format(pctChangeLow, "0.0%") & ", Close: " & Format(pctChangeClose, "0.0%")
        End If

        ' Additional check for Volume anomaly (e.g., extreme spike in volume)
        Dim currentVolume As Double
        Dim prevVolume As Double
        currentVolume = ws.Cells(i, "F").Value ' Column F is 'Volume'
        prevVolume = ws.Cells(i - 1, "F").Value

        If prevVolume <> 0 Then
            Dim pctChangeVolume As Double
            pctChangeVolume = Abs((currentVolume - prevVolume) / prevVolume)
            If pctChangeVolume > PCT_THRESHOLD_VOL Then
                If Not isAnomaly Then ' If not already flagged by price, set to true
                    isAnomaly = True
                    anomalyDetails = "Anomaly: Volume change > " & Format(PCT_THRESHOLD_VOL, "0%") & ". Volume: " & Format(pctChangeVolume, "0.0%")
                Else ' Append to existing details if already a price anomaly
                    anomalyDetails = anomalyDetails & vbNewLine & "Also Volume change > " & Format(PCT_THRESHOLD_VOL, "0%") & ". Volume: " & Format(pctChangeVolume, "0.0%")
                End If
            End If
        End If

        If isAnomaly Then
            ' Mark row as anomalous on the main sheet (e.g., highlight or add comment)
            ws.Rows(i).Interior.Color = RGB(255, 255, 0) ' Yellow highlight [20, 21, 22, 23]
            If ws.Cells(i, "A").Comment Is Nothing Then ' Add new comment if none exists
                ws.Cells(i, "A").AddComment anomalyDetails '[24, 25, 26, 27, 28, 29]
            Else ' Append to existing comment
                ws.Cells(i, "A").Comment.Text ws.Cells(i, "A").Comment.Text & vbNewLine & anomalyDetails
            End If

            ' Add anomaly details to anomaliesList sheet
            lastRowAnomalies = wsAnomalies.Cells(wsAnomalies.Rows.Count, "A").End(xlUp).Row + 1
            wsAnomalies.Cells(lastRowAnomalies, 1).Value = ws.Cells(i, "A").Value ' Date
            wsAnomalies.Cells(lastRowAnomalies, 2).Value = ws.Cells(i, "G").Value ' Ticker
            wsAnomalies.Cells(lastRowAnomalies, 3).Value = ws.Cells(i, "B").Value ' Open
            wsAnomalies.Cells(lastRowAnomalies, 4).Value = ws.Cells(i, "C").Value ' High
            wsAnomalies.Cells(lastRowAnomalies, 5).Value = ws.Cells(i, "D").Value ' Low
            wsAnomalies.Cells(lastRowAnomalies, 6).Value = ws.Cells(i, "E").Value ' Close
            wsAnomalies.Cells(lastRowAnomalies, 7).Value = ws.Cells(i, "F").Value ' Volume
            wsAnomalies.Cells(lastRowAnomalies, 8).Value = "Percentage Change Anomaly"
            wsAnomalies.Cells(lastRowAnomalies, 9).Value = anomalyDetails
        End If
    Next i
    wsAnomalies.Columns.AutoFit ' Auto-fit columns for readability
    MsgBox "Anomaly detection complete based on percentage change. Anomalies listed on 'anomaliesList' sheet.", vbInformation
End Sub

Sub ClearAnomalousCells()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Test") ' Adjust sheet name as needed
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False ' Turn off screen updating for performance

    ' Loop through rows from bottom up (good practice, though not strictly necessary for ClearContents)
    For i = lastRow To 2 Step -1 ' Assuming data starts from row 2
        ' Check if the row is flagged (e.g., by yellow highlight)
        If ws.Rows(i).Interior.Color = RGB(255, 255, 0) Then ' Check for yellow highlight
            ' Clear contents of price and volume columns (B:F)
            ws.Range(ws.Cells(i, "B"), ws.Cells(i, "F")).ClearContents [40]
            ' Optionally remove the comment and highlight after clearing
            If Not ws.Cells(i, "A").Comment Is Nothing Then
                ws.Cells(i, "A").Comment.Delete ' [25]
            End If
            ws.Rows(i).Interior.ColorIndex = xlNone ' Remove highlight [20]
        End If
    Next i
    Application.ScreenUpdating = True
    MsgBox "Anomalous cell contents cleared.", vbInformation
End Sub

Sub DetectStatisticalAnomalies()
    Dim ws As Worksheet
    Dim wsAnomalies As Worksheet
    Dim lastRow As Long
    Dim lastRowAnomalies As Long
    Dim i As Long
    Dim windowSize As Integer ' Number of previous days for rolling calculation
    Dim priceRange As Range
    Dim avgPrice As Double
    Dim stdDevPrice As Double
    Dim medianPrice As Double
    Const STD_DEV_MULTIPLIER As Double = 3 ' Number of standard deviations for anomaly threshold

    Set ws = ThisWorkbook.Sheets("Test") ' Adjust main data sheet name as needed

    ' Check if "anomaliesList" sheet exists, if not, create it
    On Error Resume Next
    Set wsAnomalies = ThisWorkbook.Sheets("anomaliesList")
    On Error GoTo 0

    If wsAnomalies Is Nothing Then
        Set wsAnomalies = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsAnomalies.Name = "anomaliesList"
        ' Add headers to the new sheet
        wsAnomalies.Cells(1, 1).Value = "Date"
        wsAnomalies.Cells(1, 2).Value = "Ticker"
        wsAnomalies.Cells(1, 3).Value = "Open"
        wsAnomalies.Cells(1, 4).Value = "High"
        wsAnomalies.Cells(1, 5).Value = "Low"
        wsAnomalies.Cells(1, 6).Value = "Close"
        wsAnomalies.Cells(1, 7).Value = "Volume"
        wsAnomalies.Cells(1, 8).Value = "Anomaly Type"
        wsAnomalies.Cells(1, 9).Value = "Details"
        wsAnomalies.Rows(1).Font.Bold = True
    End If

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    windowSize = 10 ' Example: 10-day rolling window (adjust as needed)

    ' Loop from row where a full window of data is available (e.g., row 12 for 10-day window + 1 header + 1 first data row)
    For i = (2 + windowSize) To lastRow ' Start after enough data for the window
        
    ' Jump windowSize and restart if not same Ticker
     If ws.Cells(i - 1, 7).Value <> ws.Cells(i, 7).Value Then i = i + windowSize
     
        ' Define the range for the rolling window (e.g., Close prices for the last 'windowSize' days)
        ' The range goes from 'windowSize' rows above the current row (i - windowSize) to the row just before the current row (i - 1)
        Set priceRange = ws.Range(ws.Cells(i - windowSize, "E"), ws.Cells(i - 1, "E")) ' Column E is 'Close'

        ' Calculate rolling average, median, and standard deviation
        On Error Resume Next ' Handle potential errors (e.g., non-numeric data, insufficient data)
        avgPrice = Application.WorksheetFunction.Average(priceRange)
        stdDevPrice = Application.WorksheetFunction.StDev_S(priceRange) ' StDev.S for sample standard deviation
        medianPrice = Application.WorksheetFunction.Median(priceRange)
        On Error GoTo 0

        ' Get current closing price
        Dim currentClose As Double
        currentClose = ws.Cells(i, "E").Value

        ' Check if current closing price deviates significantly from the rolling average/median
        If stdDevPrice > 0 Then ' Avoid division by zero if all values in window are same
            ' Anomaly if current price is outside STD_DEV_MULTIPLIER * stdDevPrice from average OR median
            If Abs(currentClose - avgPrice) > (STD_DEV_MULTIPLIER * stdDevPrice) Or _
               Abs(currentClose - medianPrice) > (STD_DEV_MULTIPLIER * stdDevPrice) Then ' Using std dev with median for robustness
                ' Mark row as anomalous on the main sheet
                If ws.Rows(i).Interior.Color <> RGB(255, 255, 0) Then ' Only highlight if not already highlighted
                    ws.Rows(i).Interior.Color = RGB(255, 255, 0) ' Yellow highlight
                End If
                Dim anomalyComment As String
                anomalyComment = "Anomaly: Price deviation from MA/Median. Current: " & Format(currentClose, "0.00") & ", Avg: " & Format(avgPrice, "0.00") & ", Median: " & Format(medianPrice, "0.00") & ", StdDev: " & Format(stdDevPrice, "0.00")
                If ws.Cells(i, "A").Comment Is Nothing Then
                    ws.Cells(i, "A").AddComment anomalyComment
                Else
                    ws.Cells(i, "A").Comment.Text ws.Cells(i, "A").Comment.Text & vbNewLine & anomalyComment
                End If

                ' Add anomaly details to anomaliesList sheet
                lastRowAnomalies = wsAnomalies.Cells(wsAnomalies.Rows.Count, "A").End(xlUp).Row + 1
                wsAnomalies.Cells(lastRowAnomalies, 1).Value = ws.Cells(i, "A").Value ' Date
                wsAnomalies.Cells(lastRowAnomalies, 2).Value = ws.Cells(i, "G").Value ' Ticker
                wsAnomalies.Cells(lastRowAnomalies, 3).Value = ws.Cells(i, "B").Value ' Open
                wsAnomalies.Cells(lastRowAnomalies, 4).Value = ws.Cells(i, "C").Value ' High
                wsAnomalies.Cells(lastRowAnomalies, 5).Value = ws.Cells(i, "D").Value ' Low
                wsAnomalies.Cells(lastRowAnomalies, 6).Value = ws.Cells(i, "E").Value ' Close
                wsAnomalies.Cells(lastRowAnomalies, 7).Value = ws.Cells(i, "F").Value ' Volume
                wsAnomalies.Cells(lastRowAnomalies, 8).Value = "Statistical Price Anomaly"
                wsAnomalies.Cells(lastRowAnomalies, 9).Value = anomalyComment
            End If
        ElseIf currentClose <> avgPrice Then ' If stdDev is 0 (all previous values were identical) but current price is different (spike from flat line)
            If ws.Rows(i).Interior.Color <> RGB(255, 255, 0) Then
                ws.Rows(i).Interior.Color = RGB(255, 255, 0)
            End If
            Dim anomalyCommentZeroStdDev As String
            anomalyCommentZeroStdDev = "Anomaly: Price deviation from constant MA/Median. Current: " & Format(currentClose, "0.00") & ", Avg: " & Format(avgPrice, "0.00")
            If ws.Cells(i, "A").Comment Is Nothing Then
                ws.Cells(i, "A").AddComment anomalyCommentZeroStdDev
            Else
                ws.Cells(i, "A").Comment.Text ws.Cells(i, "A").Comment.Text & vbNewLine & anomalyCommentZeroStdDev
            End If

            ' Add anomaly details to anomaliesList sheet
            lastRowAnomalies = wsAnomalies.Cells(wsAnomalies.Rows.Count, "A").End(xlUp).Row + 1
            wsAnomalies.Cells(lastRowAnomalies, 1).Value = ws.Cells(i, "A").Value ' Date
            wsAnomalies.Cells(lastRowAnomalies, 2).Value = ws.Cells(i, "G").Value ' Ticker
            wsAnomalies.Cells(lastRowAnomalies, 3).Value = ws.Cells(i, "B").Value ' Open
            wsAnomalies.Cells(lastRowAnomalies, 4).Value = ws.Cells(i, "C").Value ' High
            wsAnomalies.Cells(lastRowAnomalies, 5).Value = ws.Cells(i, "D").Value ' Low
            wsAnomalies.Cells(lastRowAnomalies, 6).Value = ws.Cells(i, "E").Value ' Close
            wsAnomalies.Cells(lastRowAnomalies, 7).Value = ws.Cells(i, "F").Value ' Volume
            wsAnomalies.Cells(lastRowAnomalies, 8).Value = "Statistical Price Anomaly (Zero StdDev)"
            wsAnomalies.Cells(lastRowAnomalies, 9).Value = anomalyCommentZeroStdDev
        End If
    Next i
    wsAnomalies.Columns.AutoFit ' Auto-fit columns for readability
    MsgBox "Anomaly detection complete based on statistical deviation. Anomalies listed on 'anomaliesList' sheet.", vbInformation
End Sub

Sub origDetectStatisticalAnomalies()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim windowSize As Integer ' Number of previous days for rolling calculation
    Dim priceRange As Range
    Dim avgPrice As Double
    Dim stdDevPrice As Double
    Dim medianPrice As Double
    Const STD_DEV_MULTIPLIER As Double = 3.25 ' Number of standard deviations for anomaly threshold

    Set ws = ThisWorkbook.Sheets("Test") ' Adjust sheet name as needed
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    windowSize = 10 ' Example: 10-day rolling window (adjust as needed)

    ' Loop from row where a full window of data is available (e.g., row 12 for 10-day window + 1 header + 1 first data row)
    For i = (2 + windowSize) To lastRow ' Start after enough data for the window
    
    ' Skip 10 if not same Ticker
     If ws.Cells(i - 1, 7).Value <> ws.Cells(i, 7).Value Then i = i + windowSize
            
        ' Define the range for the rolling window (e.g., Close prices for the last 'windowSize' days)
        ' The range goes from 'windowSize' rows above the current row (i - windowSize) to the row just before the current row (i - 1)
        Set priceRange = ws.Range(ws.Cells(i - windowSize, "E"), ws.Cells(i - 1, "E")) ' Column E is 'Close'

        ' Calculate rolling average, median, and standard deviation
        On Error Resume Next ' Handle potential errors (e.g., non-numeric data, insufficient data)
        avgPrice = Application.WorksheetFunction.Average(priceRange)
        stdDevPrice = Application.WorksheetFunction.StDev_S(priceRange) ' StDev.S for sample standard deviation
        medianPrice = Application.WorksheetFunction.Median(priceRange)
        On Error GoTo 0

        ' Get current closing price
        Dim currentClose As Double
        currentClose = ws.Cells(i, "E").Value

        ' Check if current closing price deviates significantly from the rolling average/median
        If stdDevPrice > 0 Then ' Avoid division by zero if all values in window are same
            ' Anomaly if current price is outside STD_DEV_MULTIPLIER * stdDevPrice from average OR median
            If Abs(currentClose - avgPrice) > (STD_DEV_MULTIPLIER * stdDevPrice) Or _
               Abs(currentClose - medianPrice) > (STD_DEV_MULTIPLIER * stdDevPrice) Then ' Using std dev with median for robustness
                ' Mark row as anomalous
                If ws.Rows(i).Interior.Color <> RGB(255, 255, 0) Then ' Only highlight if not already highlighted
                    ws.Rows(i).Interior.Color = RGB(255, 255, 0) ' Yellow highlight
                End If
                Dim anomalyComment As String
                anomalyComment = "Anomaly: Price deviation from MA/Median. Current: " & Format(currentClose, "0.00") & ", Avg: " & Format(avgPrice, "0.00") & ", Median: " & Format(medianPrice, "0.00") & ", StdDev: " & Format(stdDevPrice, "0.00")
                If ws.Cells(i, "A").Comment Is Nothing Then
                    ws.Cells(i, "A").AddComment anomalyComment
                Else
                    ws.Cells(i, "A").Comment.Text ws.Cells(i, "A").Comment.Text & vbNewLine & anomalyComment
                End If
            End If
        ElseIf currentClose <> avgPrice Then ' If stdDev is 0 (all previous values were identical) but current price is different (spike from flat line)
            If ws.Rows(i).Interior.Color <> RGB(255, 255, 0) Then
                ws.Rows(i).Interior.Color = RGB(255, 255, 0)
            End If
            debig.Print ws.Rows(i)
            Dim anomalyCommentZeroStdDev As String
            anomalyCommentZeroStdDev = "Anomaly: Price deviation from constant MA/Median. Current: " & Format(currentClose, "0.00") & ", Avg: " & Format(avgPrice, "0.00")
            If ws.Cells(i, "A").Comment Is Nothing Then
                ws.Cells(i, "A").AddComment anomalyCommentZeroStdDev
            Else
                ws.Cells(i, "A").Comment.Text ws.Cells(i, "A").Comment.Text & vbNewLine & anomalyCommentZeroStdDev
            End If
        End If
    Next i
    MsgBox "Anomaly detection complete based on statistical deviation.", vbInformation
End Sub

Sub ClearAnomalousCells()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Test") ' Adjust sheet name as needed
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False ' Turn off screen updating for performance

    ' Loop through rows from bottom up (good practice, though not strictly necessary for ClearContents)
    For i = lastRow To 2 Step -1 ' Assuming data starts from row 2
        ' Check if the row is flagged (e.g., by yellow highlight)
        If ws.Rows(i).Interior.Color = RGB(255, 255, 0) Then ' Check for yellow highlight
            ' Clear contents of price and volume columns (B:F)
            ws.Range(ws.Cells(i, "B"), ws.Cells(i, "F")).ClearContents [40]
            ' Optionally remove the comment and highlight after clearing
            If Not ws.Cells(i, "A").Comment Is Nothing Then
                ws.Cells(i, "A").Comment.Delete ' [25]
            End If
            ws.Rows(i).Interior.ColorIndex = xlNone ' Remove highlight [20]
        End If
    Next i
    Application.ScreenUpdating = True
    MsgBox "Anomalous cell contents cleared.", vbInformation
End Sub

Sub origDetectPriceChangeAnomalies()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim prevClose As Double
    Dim currentOpen As Double, currentHigh As Double, currentLow As Double, currentClose As Double
    Dim pctChangeOpen As Double, pctChangeHigh As Double, pctChangeLow As Double, pctChangeClose As Double
    Const PCT_THRESHOLD As Double = 1#  ' Example: 50% change threshold (adjust as needed, e.g., 1.0 for 100%)
    Const PCT_THRESHOLD_VOL As Double = 2#  ' Example: 200% change threshold for Volume

    Set ws = ThisWorkbook.Sheets("Test") ' Adjust sheet name as needed
    ' Find the last row with data in Column A (Date)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop from the second data row (row 3, assuming header in row 1, first data in row 2)
    For i = 3 To lastRow
        ' Get previous day's closing price (from row i-1)
        prevClose = ws.Cells(i - 1, "E").Value ' Column E is 'Close'

        ' Get current day's prices
        currentOpen = ws.Cells(i, "B").Value ' Column B is 'Open'
        currentHigh = ws.Cells(i, "C").Value ' Column C is 'High'
        currentLow = ws.Cells(i, "D").Value ' Column D is 'Low'
        currentClose = ws.Cells(i, "E").Value ' Column E is 'Close'

        ' Calculate percentage changes relative to previous day's close
        If prevClose <> 0 Then ' Avoid division by zero
            pctChangeOpen = Abs((currentOpen - prevClose) / prevClose)
            pctChangeHigh = Abs((currentHigh - prevClose) / prevClose)
            pctChangeLow = Abs((currentLow - prevClose) / prevClose)
            pctChangeClose = Abs((currentClose - prevClose) / prevClose)
        Else
            ' Handle cases where previous close is zero (e.g., new listing, or extreme anomaly)
            ' If any current price is non-zero and prevClose is zero, it's a significant change
            If currentOpen <> 0 Or currentHigh <> 0 Or currentLow <> 0 Or currentClose <> 0 Then
                pctChangeOpen = 999 ' Assign a very high value to trigger anomaly
                pctChangeHigh = 999
                pctChangeLow = 999
                pctChangeClose = 999
            Else
                pctChangeOpen = 0 ' No change if both are zero
                pctChangeHigh = 0
                pctChangeLow = 0
                pctChangeClose = 0
            End If
        End If

        ' Check for anomaly based on price change threshold
        If pctChangeOpen > PCT_THRESHOLD Or _
           pctChangeHigh > PCT_THRESHOLD Or _
           pctChangeLow > PCT_THRESHOLD Or _
           pctChangeClose > PCT_THRESHOLD Then
            ' Mark row as anomalous (e.g., highlight or add comment)
            ws.Rows(i).Interior.Color = RGB(255, 255, 0) ' Yellow highlight [20, 21, 22, 23]
            If ws.Cells(i, "A").Comment Is Nothing Then ' Add new comment if none exists
                ws.Cells(i, "A").AddComment "Anomaly: Price change > " & Format(PCT_THRESHOLD, "0%") ' [24, 25, 26, 27, 28, 29]
            Else ' Append to existing comment
                ws.Cells(i, "A").Comment.Text ws.Cells(i, "A").Comment.Text & vbNewLine & "Anomaly: Price change > " & Format(PCT_THRESHOLD, "0%")
            End If
        End If

        ' Additional check for Volume anomaly (e.g., extreme spike in volume)
        Dim currentVolume As Double
        Dim prevVolume As Double
        currentVolume = ws.Cells(i, "F").Value ' Column F is 'Volume'
        prevVolume = ws.Cells(i - 1, "F").Value

        If prevVolume <> 0 Then
            Dim pctChangeVolume As Double
            pctChangeVolume = Abs((currentVolume - prevVolume) / prevVolume)
            If pctChangeVolume > PCT_THRESHOLD_VOL Then
                If ws.Rows(i).Interior.Color <> RGB(255, 255, 0) Then ' Only highlight if not already highlighted
                    ws.Rows(i).Interior.Color = RGB(255, 255, 0)
                End If
                'If ws.Cells(i, "A").Comment Is Nothing Then
                 '   ws.Cells(i, "A").AddComment "Anomaly: Volume change > " & Format(PCT_THRESHOLD_VOL, "0%")
                'Else
                '    ws.Cells(i, "A").Comment.Text ws.Cells(i, "A").Comment.Text & vbNewLine & "Anomaly: Volume change > " & Format(PCT_THRESHOLD_VOL, "0%")
                'End If
            End If
        End If
    Next i
    MsgBox "Anomaly detection complete based on percentage change.", vbInformation
End Sub

Sub DeleteAnomalousRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Test") ' Adjust sheet name as needed
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Application.DisplayAlerts = False ' Disable alerts to prevent pop-ups [44]
    Application.ScreenUpdating = False ' Turn off screen updating for performance

    ' Loop from the bottom up to avoid skipping rows when deleting [42, 43]
    For i = lastRow To 2 Step -1 ' Assuming data starts from row 2
        ' Check if the row is flagged (e.g., by yellow highlight)
        If ws.Rows(i).Interior.Color = RGB(255, 255, 0) Then ' Check for yellow highlight
            ws.Rows(i).Delete Shift:=xlUp ' Delete the entire row, shifting cells up [42]
        End If
    Next i

    Application.DisplayAlerts = True ' Re-enable alerts
    Application.ScreenUpdating = True
    MsgBox "Anomalous rows deleted.", vbInformation
End Sub

Sub CalcPnL()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("TRADE LOG")
  Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
  Dim i As Long

  For i = 4 To lastRow
    ws.Cells(i, "R").Value = (ws.Cells(i, "D").Value - ws.Cells(i, "O").Value) / ws.Cells(i, "O").Value
    ws.Cells(i, "S").Value = (ws.Cells(i, "O").Value - ws.Cells(i, "D").Value) / ws.Cells(i, "D").Value
  Next i

  MsgBox "PnL% calculated for " & lastRow - 1 & " trades"
End Sub

Sub qFilter()

 Range("A1:G1").AutoFilter
End Sub
