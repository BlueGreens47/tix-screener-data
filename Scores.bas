Attribute VB_Name = "Scores"
Option Explicit

' ============================================================================
' STOCK SCORING SYSTEM - VBA MODULE (Improved to match Python layout)
' ============================================================================
' This module calculates composite stock scores using sector-adjusted percentiles
' Formula: Weighted average of key financial metrics converted to percentile ranks
' Updated to match Python script column layout and data structure
' ============================================================================

' Main function to calculate composite stock score
Public Function CalculateStockScore(ticker As String, sector As String, _
    current_price As Double, eps As Double, industry_pe As Double, _
    pe_to_industry As Double, growth_rate As Double, fair_value As Double, _
    fair_value_ratio As Double, roe_percent As Double, debt_equity As Double, _
    fcf_margin_percent As Double, _
    Optional price_weight As Double = 0.1, _
    Optional eps_weight As Double = 0.1, _
    Optional pe_weight As Double = 0.15, _
    Optional growth_weight As Double = 0.2, _
    Optional value_weight As Double = 0.15, _
    Optional roe_weight As Double = 0.15, _
    Optional debt_weight As Double = 0.1, _
    Optional fcf_weight As Double = 0.05) As Double
    
    ' Get the data range for sector comparison
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TJX")
    
    ' Find the data range (assumes data starts in row 3, headers in rows 1-2)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Calculate percentile ranks for each metric within the sector
    Dim price_percentile As Double
    Dim eps_percentile As Double
    Dim pe_industry_percentile As Double
    Dim growth_percentile As Double
    Dim value_percentile As Double
    Dim roe_percentile As Double
    Dim debt_percentile As Double
    Dim fcf_percentile As Double
    
    ' Calculate percentiles (lower is better for some metrics)
    price_percentile = CalculateSectorPercentile(ws, sector, "PRICE2", current_price, lastRow)
    eps_percentile = CalculateSectorPercentile(ws, sector, "EPS", eps, lastRow)
    pe_industry_percentile = CalculateSectorPercentile(ws, sector, "PE_INDUSTRY", pe_to_industry, lastRow, True) ' Lower is better
    growth_percentile = CalculateSectorPercentile(ws, sector, "GROWTH", growth_rate, lastRow)
    value_percentile = CalculateSectorPercentile(ws, sector, "VALUE_RATIO", fair_value_ratio, lastRow, True) ' Lower is better (undervalued)
    roe_percentile = CalculateSectorPercentile(ws, sector, "ROE", roe_percent, lastRow)
    debt_percentile = CalculateSectorPercentile(ws, sector, "DEBT", debt_equity, lastRow, True) ' Lower is better
    fcf_percentile = CalculateSectorPercentile(ws, sector, "FCF", fcf_margin_percent, lastRow)
    
    ' Calculate weighted composite score
    CalculateStockScore = (price_percentile * price_weight) + _
                         (eps_percentile * eps_weight) + _
                         (pe_industry_percentile * pe_weight) + _
                         (growth_percentile * growth_weight) + _
                         (value_percentile * value_weight) + _
                         (roe_percentile * roe_weight) + _
                         (debt_percentile * debt_weight) + _
                         (fcf_percentile * fcf_weight)
    
End Function

' Helper function to calculate sector-adjusted percentile rank
Private Function CalculateSectorPercentile(ws As Worksheet, targetSector As String, _
    metricType As String, targetValue As Double, lastRow As Long, _
    Optional lowerIsBetter As Boolean = False) As Double
    
    ' Column mapping to match Python script layout exactly
    Dim tickerCol As Long: tickerCol = 1          ' Column A for Ticker
    Dim industryCol As Long: industryCol = 6      ' Column F for Industry
    Dim sectorCol As Long: sectorCol = 8          ' Column H for Sector
    Dim priceCol As Long: priceCol = 12           ' Column L for Current Price
    Dim epsCol As Long: epsCol = 13               ' Column M for EPS
    Dim industryPeCol As Long: industryPeCol = 14 ' Column N for Industry PE
    Dim peToIndustryCol As Long: peToIndustryCol = 15 ' Column O for PE to Industry
    Dim growthCol As Long: growthCol = 16         ' Column P for Growth Rate
    Dim fairValueCol As Long: fairValueCol = 17   ' Column Q for Fair Value
    Dim valueRatioCol As Long: valueRatioCol = 18 ' Column R for Fair Value Ratio
    Dim roeCol As Long: roeCol = 19               ' Column S for ROE%
    Dim debtCol As Long: debtCol = 20             ' Column T for Debt/Equity
    Dim fcfCol As Long: fcfCol = 21               ' Column U for FCF Margin%
    
    Dim metricCol As Long
    
    ' Determine which column to use based on metric type
    Select Case UCase(metricType)
        Case "PRICE2": metricCol = priceCol
        Case "EPS": metricCol = epsCol
        Case "INDUSTRY_PE": metricCol = industryPeCol
        Case "PE_INDUSTRY": metricCol = peToIndustryCol
        Case "GROWTH": metricCol = growthCol
        Case "FAIR_VALUE": metricCol = fairValueCol
        Case "VALUE_RATIO": metricCol = valueRatioCol
        Case "ROE": metricCol = roeCol
        Case "DEBT": metricCol = debtCol
        Case "FCF": metricCol = fcfCol
        Case Else: CalculateSectorPercentile = 50: Exit Function ' Neutral score for unknown metrics
    End Select
    
    ' Collect all values for the same sector
    Dim sectorValues As Collection
    Set sectorValues = New Collection
    
    Dim i As Long
    For i = 3 To lastRow ' Start from row 3 to match Python script
        If UCase(ws.Cells(i, sectorCol).value) = UCase(targetSector) Then
            If IsNumeric(ws.Cells(i, metricCol).value) And ws.Cells(i, metricCol).value <> "" Then
                Dim cellValue As Variant
                cellValue = ws.Cells(i, metricCol).value
                ' Handle "N/A", "Error" and error values (matching Python)
                If cellValue <> "N/A" And cellValue <> "Error" And Not IsError(cellValue) Then
                    ' Basic sanity check - filter out unreasonable values
                    Dim numValue As Double
                    numValue = CDbl(cellValue)
                    If IsValidMetricValue(metricType, numValue) Then
                        sectorValues.Add numValue
                    End If
                End If
            End If
        End If
    Next i
    
    ' If no sector data found, return 50 (neutral score)
    If sectorValues.count = 0 Then
        CalculateSectorPercentile = 50
        Exit Function
    End If
    
    ' Count how many values are better/worse than target
    Dim betterCount As Long
    Dim totalCount As Long
    Dim value As Variant
    
    totalCount = sectorValues.count
    betterCount = 0
    
    For Each value In sectorValues
        If lowerIsBetter Then
            If CDbl(value) > targetValue Then betterCount = betterCount + 1
        Else
            If CDbl(value) < targetValue Then betterCount = betterCount + 1
        End If
    Next value
    
    ' Calculate percentile (higher percentile = better performance)
    CalculateSectorPercentile = (betterCount / totalCount) * 100
    
End Function

' New function to validate metric values (matches Python filtering logic)
Private Function IsValidMetricValue(metricType As String, value As Double) As Boolean
    IsValidMetricValue = True
    
    Select Case UCase(metricType)
        Case "PE_INDUSTRY", "GROWTH"
            ' PE ratios and growth rates should be reasonable
            If value < 0 Or value > 200 Then IsValidMetricValue = False
        Case "PRICE2"
            ' Stock prices should be positive
            If value <= 0 Then IsValidMetricValue = False
        Case "EPS"
            ' EPS can be negative but not extremely so
            If value < -100 Then IsValidMetricValue = False
        Case "ROE"
            ' ROE as percentage should be reasonable
            If value < -100 Or value > 200 Then IsValidMetricValue = False
        Case "DEBT"
            ' Debt/Equity should be non-negative and reasonable
            If value < 0 Or value > 10 Then IsValidMetricValue = False
        Case "FCF"
            ' FCF margin can be negative but not extremely so
            If value < -200 Or value > 200 Then IsValidMetricValue = False
        Case "VALUE_RATIO"
            ' Price-to-fair-value ratio should be positive and reasonable
            If value <= 0 Or value > 10 Then IsValidMetricValue = False
    End Select
End Function

' Updated utility function to calculate score for entire range
Public Sub CalculateAllScores()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TJX")
    
    ' Find last row with data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Add header for score column (Column V - 22)
    ws.Cells(2, 22).value = "Composite Score"
    ws.Cells(2, 22).Font.Bold = True
    
    ' Progress tracking
    Dim processedCount As Long
    Dim totalCount As Long
    totalCount = lastRow - 2 ' Subtract header rows
    processedCount = 0
    
    ' Performance optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Calculate score for each stock
    Dim i As Long
    For i = 3 To lastRow ' Start from row 3
        If ws.Cells(i, 1).value <> "" And ws.Cells(i, 1).value <> "N/A" Then ' Check if ticker exists
            
            ' Get sector information
            Dim sector As String
            sector = ws.Cells(i, 8).value ' Column H
            
            ' Skip if no sector information
            If sector = "" Or sector = "N/A" Or sector = "Error" Then
                ws.Cells(i, 22).value = "N/A"
                GoTo NextIteration
            End If
            
            ' Get all required values with error checking
            Dim ticker As String, current_price As Double, eps As Double
            Dim industry_pe As Double, pe_to_industry As Double, growth_rate As Double
            Dim fair_value As Double, fair_value_ratio As Double
            Dim roe_percent As Double, debt_equity As Double, fcf_margin_percent As Double
            
            ticker = ws.Cells(i, 1).value
            
            ' Helper function to safely convert cell values to numbers
            current_price = SafeConvertToDouble(ws.Cells(i, 12).value)
            eps = SafeConvertToDouble(ws.Cells(i, 13).value)
            industry_pe = SafeConvertToDouble(ws.Cells(i, 14).value)
            pe_to_industry = SafeConvertToDouble(ws.Cells(i, 15).value)
            growth_rate = SafeConvertToDouble(ws.Cells(i, 16).value)
            fair_value = SafeConvertToDouble(ws.Cells(i, 17).value)
            fair_value_ratio = SafeConvertToDouble(ws.Cells(i, 18).value)
            roe_percent = SafeConvertToDouble(ws.Cells(i, 19).value)
            debt_equity = SafeConvertToDouble(ws.Cells(i, 20).value)
            fcf_margin_percent = SafeConvertToDouble(ws.Cells(i, 21).value)
            
            ' Check if we have enough valid data to calculate a meaningful score
            Dim validDataCount As Long
            validDataCount = 0
            If current_price <> 0 Then validDataCount = validDataCount + 1
            If eps <> 0 Then validDataCount = validDataCount + 1
            If pe_to_industry <> 0 Then validDataCount = validDataCount + 1
            If growth_rate <> 0 Then validDataCount = validDataCount + 1
            If fair_value_ratio <> 0 Then validDataCount = validDataCount + 1
            If roe_percent <> 0 Then validDataCount = validDataCount + 1
            If debt_equity <> 0 Then validDataCount = validDataCount + 1
            If fcf_margin_percent <> 0 Then validDataCount = validDataCount + 1
            
            ' Only calculate score if we have at least 4 valid metrics
            If validDataCount >= 4 Then
                ' Calculate composite score
                Dim score As Double
                score = CalculateStockScore(ticker, sector, current_price, eps, industry_pe, _
                                          pe_to_industry, growth_rate, fair_value, fair_value_ratio, _
                                          roe_percent, debt_equity, fcf_margin_percent)
                
                ' Write score to column V (22)
                ws.Cells(i, 22).value = score
            Else
                ' Not enough data for meaningful score
                ws.Cells(i, 22).value = ""
            End If
            
NextIteration:
            ' Update progress
            processedCount = processedCount + 1
            If processedCount Mod 25 = 0 Or processedCount = totalCount Then
                Application.StatusBar = "Processing stocks: " & processedCount & " of " & totalCount & " (" & Format(processedCount / totalCount, "0%") & ")"
                DoEvents
            End If
        End If
    Next i
    
    ' Format the score column
    ws.Range("V3:V" & lastRow).NumberFormat = "0.00"
    
    ' Add conditional formatting for easy visualization
    Call AddConditionalFormatting(ws.Range("V3:V" & lastRow))
    
    ' Auto-fit the score column
    ws.Columns(22).AutoFit
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Clear status bar
    Application.StatusBar = False
    
    MsgBox "Stock scores calculated successfully for " & processedCount & " stocks!" & vbCrLf & _
           "Scores range from 0-100, with higher scores indicating better relative performance within sector.", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "Error occurred during score calculation: " & Err.Description, vbCritical
End Sub

' Enhanced helper function to safely convert values to Double (matches Python logic)
Private Function SafeConvertToDouble(cellValue As Variant) As Double
    If IsEmpty(cellValue) Or cellValue = "" Or cellValue = "N/A" Or cellValue = "Error" Or IsError(cellValue) Then
        SafeConvertToDouble = 0
    Else
        If IsNumeric(cellValue) Then
            Dim numValue As Double
            numValue = CDbl(cellValue)
            ' Additional validation to match Python filtering
            If numValue <> numValue Then ' Check for NaN
                SafeConvertToDouble = 0
            Else
                SafeConvertToDouble = numValue
            End If
        Else
            SafeConvertToDouble = 0
        End If
    End If
End Function

' Enhanced conditional formatting (improved color scheme)
Private Sub AddConditionalFormatting(rng As Range)
    rng.FormatConditions.Delete
    
    ' Excellent performers (>80)
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="80", Formula2:="100")
        .Interior.Color = RGB(34, 139, 34) ' Forest Green
        .Font.Color = RGB(255, 255, 255) ' White text
        .Font.Bold = True
    End With
    
    ' Good performers (60-80)
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="60", Formula2:="80")
        .Interior.Color = RGB(144, 238, 144) ' Light green
        .Font.Color = RGB(0, 100, 0) ' Dark green
    End With
    
    ' Average performers (40-60)
   ' With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="40", Formula2:="60")
   '     .Interior.Color = RGB(255, 255, 224) ' Light yellow
   '     .Font.Color = RGB(139, 69, 19) ' Brown
   ' End With
    
    ' Below average performers (20-40)
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="20", Formula2:="40")
        .Interior.Color = RGB(255, 255, 224) ' Light yellow' RGB(255, 165, 0) ' Orange
        .Font.Color = RGB(139, 69, 19) ' Brown RGB(255, 255, 255) ' White text
    End With
    
    ' Poor performers (<20)
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="1", Formula2:="20")
        .Interior.Color = RGB(220, 20, 60) ' Crimson
        .Font.Color = RGB(255, 255, 255) ' White text
        .Font.Bold = True
    End With
End Sub

' Enhanced sector statistics function
Public Function GetSectorStats(sector As String, metric As String) As String
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Updated column mapping to match exactly
    Dim sectorCol As Long: sectorCol = 8
    Dim metricCol As Long
    
    Select Case UCase(metric)
        Case "PRICE2": metricCol = 12
        Case "EPS": metricCol = 13
        Case "INDUSTRY_PE": metricCol = 14
        Case "PE_INDUSTRY": metricCol = 15
        Case "GROWTH": metricCol = 16
        Case "FAIR_VALUE": metricCol = 17
        Case "VALUE_RATIO": metricCol = 18
        Case "ROE": metricCol = 19
        Case "DEBT": metricCol = 20
        Case "FCF": metricCol = 21
        Case "SCORE": metricCol = 22
        Case Else: GetSectorStats = "Invalid metric": Exit Function
    End Select
    
    ' Collect sector values with filtering
    Dim values As Collection
    Set values = New Collection
    
    Dim i As Long
    For i = 3 To lastRow
        If UCase(ws.Cells(i, sectorCol).value) = UCase(sector) Then
            Dim cellValue As Variant
            cellValue = ws.Cells(i, metricCol).value
            If IsNumeric(cellValue) And cellValue <> "N/A" And cellValue <> "Error" And Not IsError(cellValue) Then
                Dim numValue As Double
                numValue = CDbl(cellValue)
                If IsValidMetricValue(metric, numValue) Then
                    values.Add numValue
                End If
            End If
        End If
    Next i
    
    If values.count = 0 Then
        GetSectorStats = "No valid data found for " & sector & " in " & metric
        Exit Function
    End If
    
    ' Calculate comprehensive statistics with outlier removal
    Dim sum As Double, avg As Double, min As Double, max As Double
    Dim median As Double, q1 As Double, q3 As Double
    Dim value As Variant
    
    sum = 0
    min = 999999
    max = -999999
    
    ' Calculate basic stats
    For Each value In values
        sum = sum + CDbl(value)
        If CDbl(value) < min Then min = CDbl(value)
        If CDbl(value) > max Then max = CDbl(value)
    Next value
    
    avg = sum / values.count
    
    ' Calculate standard deviation for outlier detection
    Dim variance As Double, stdDev As Double
    variance = 0
    For Each value In values
        variance = variance + (CDbl(value) - avg) ^ 2
    Next value
    variance = variance / values.count
    stdDev = Sqr(variance)
    
    ' Filter outliers (more than 2 standard deviations from mean)
    Dim filteredValues As Collection
    Set filteredValues = New Collection
    For Each value In values
        If Abs(CDbl(value) - avg) <= 2 * stdDev Then
            filteredValues.Add CDbl(value)
        End If
    Next value
    
    ' Use filtered values if we removed outliers
    If filteredValues.count > 0 And filteredValues.count < values.count Then
        Set values = filteredValues
        ' Recalculate stats
        sum = 0
        min = 999999
        max = -999999
        For Each value In values
            sum = sum + CDbl(value)
            If CDbl(value) < min Then min = CDbl(value)
            If CDbl(value) > max Then max = CDbl(value)
        Next value
        avg = sum / values.count
    End If
    
    ' Sort values for quartile calculations
    Dim sortedValues() As Double
    ReDim sortedValues(1 To values.count)
    Dim idx As Long
    idx = 1
    For Each value In values
        sortedValues(idx) = CDbl(value)
        idx = idx + 1
    Next value
    
    ' Simple bubble sort
    Dim j As Long, temp As Double
    For i = 1 To values.count - 1
        For j = 1 To values.count - i
            If sortedValues(j) > sortedValues(j + 1) Then
                temp = sortedValues(j)
                sortedValues(j) = sortedValues(j + 1)
                sortedValues(j + 1) = temp
            End If
        Next j
    Next i
    
    ' Calculate quartiles
    Dim midpoint As Long
    midpoint = values.count \ 2
    If values.count Mod 2 = 0 Then
        median = (sortedValues(midpoint) + sortedValues(midpoint + 1)) / 2
    Else
        median = sortedValues(midpoint + 1)
    End If
    
    q1 = sortedValues(values.count \ 4 + 1)
    q3 = sortedValues((3 * values.count) \ 4 + 1)
    
    GetSectorStats = "Sector: " & sector & " | Metric: " & metric & vbCrLf & _
                    "Count: " & values.count & " | Avg: " & Format(avg, "0.00") & " | StdDev: " & Format(stdDev, "0.00") & vbCrLf & _
                    "Min: " & Format(min, "0.00") & " | Q1: " & Format(q1, "0.00") & _
                    " | Median: " & Format(median, "0.00") & " | Q3: " & Format(q3, "0.00") & " | Max: " & Format(max, "0.00")
End Function

' Enhanced top performers function with additional filtering
Public Function GetTopPerformers(sector As String, Optional topN As Long = 5) As String
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Collect tickers and scores for the sector
    Dim performers As Collection
    Set performers = New Collection
    
    Dim i As Long
    For i = 3 To lastRow
        If UCase(ws.Cells(i, 8).value) = UCase(sector) Then
            Dim scoreValue As Variant
            scoreValue = ws.Cells(i, 22).value
            
            If IsNumeric(scoreValue) And scoreValue <> "N/A" And scoreValue <> "Error" And scoreValue <> "Insufficient Data" Then
                Dim tickerScore As String
                tickerScore = ws.Cells(i, 1).value & "|" & scoreValue
                performers.Add tickerScore
            End If
        End If
    Next i
    
    If performers.count = 0 Then
        GetTopPerformers = "No scored stocks found in " & sector
        Exit Function
    End If
    
    ' Sort performers by score (simple selection sort for top N)
    Dim result As String
    result = "Top " & topN & " performers in " & sector & ":" & vbCrLf
    
    Dim count As Long
    count = 0
    Dim maxScore As Double
    Dim maxTicker As String
    Dim usedItems As Collection
    Set usedItems = New Collection
    
    Do While count < topN And count < performers.count
        maxScore = -1
        maxTicker = ""
        
        Dim item As Variant
        For Each item In performers
            Dim parts As Variant
            parts = Split(CStr(item), "|")
            Dim ticker As String, score As Double
            ticker = parts(0)
            score = CDbl(parts(1))
            
            ' Check if already used
            Dim alreadyUsed As Boolean
            alreadyUsed = False
            Dim usedItem As Variant
            For Each usedItem In usedItems
                If CStr(usedItem) = ticker Then
                    alreadyUsed = True
                    Exit For
                End If
            Next usedItem
            
            If Not alreadyUsed And score > maxScore Then
                maxScore = score
                maxTicker = ticker
            End If
        Next item
        
        If maxTicker <> "" Then
            result = result & (count + 1) & ". " & maxTicker & " - " & Format(maxScore, "0.00") & vbCrLf
            usedItems.Add maxTicker
            count = count + 1
        Else
            Exit Do
        End If
    Loop
    
    GetTopPerformers = result
End Function

' Enhanced update function matching Python's bulk update approach
Sub UpdateTJXFromTickerFile()
    Dim sourceWb As Workbook
    Dim targetWb As Workbook
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim tickerFilePath As String
    Dim lastRow As Long
    
    ' Turn off screen updating and alerts for better performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Set the target workbook
    Set targetWb = ThisWorkbook
    
    ' Specify the path to TickerFile.xlsx
    tickerFilePath = ThisWorkbook.Path & "\TickerFile.xlsx"
    
    ' Check if TickerFile exists
    If Dir(tickerFilePath) = "" Then
        MsgBox "TickerFile.xlsx not found at: " & tickerFilePath
        GoTo CleanUp
    End If
    
    ' Open TickerFile.xlsx
    Set sourceWb = Workbooks.Open(tickerFilePath, UpdateLinks:=False, ReadOnly:=True)
    Set sourceWs = sourceWb.Sheets("TJX")
    
    ' Find TJX sheet in target workbook
    Set targetWs = targetWb.Sheets("TJX")
    
    ' Find the data range
    lastRow = sourceWs.Cells(sourceWs.Rows.count, 1).End(xlUp).row
    
    ' Clear existing data in target (except headers)
    targetWs.Range("A3:V" & targetWs.Cells(targetWs.Rows.count, 1).End(xlUp).row).ClearContents
    
    ' Copy the entire data range (matching Python's approach)
    If lastRow > 2 Then
        sourceWs.Range("A3:V" & lastRow).Copy
        targetWs.Range("A3").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    End If
    
    MsgBox "TJX data updated successfully from TickerFile.xlsx" & vbCrLf & _
           "Rows processed: " & (lastRow - 2), vbInformation
    
CleanUp:
    ' Close TickerFile without saving
    If Not sourceWb Is Nothing Then
        sourceWb.Close SaveChanges:=False
    End If
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    GoTo CleanUp
    
End Sub

' Enhanced update function matching Python's bulk update approach
Sub origUpdateTJXFromTickerFile()
    Dim sourceWb As Workbook
    Dim targetWb As Workbook
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim tickerFilePath As String
    Dim lastRow As Long
    
    ' Turn off screen updating and alerts for better performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' Set the target workbook
    Set targetWb = ThisWorkbook
    
    ' Specify the path to TickerFile.xlsx
    tickerFilePath = ThisWorkbook.Path & "\TickerFile.xlsx"
    
    ' Check if TickerFile exists
    If Dir(tickerFilePath) = "" Then
        MsgBox "TickerFile.xlsx not found at: " & tickerFilePath
        GoTo CleanUp
    End If
    
    ' Open TickerFile.xlsx
    Set sourceWb = Workbooks.Open(tickerFilePath, UpdateLinks:=False, ReadOnly:=True)
    Set sourceWs = sourceWb.Sheets("TJX")
    
    ' Find TJX sheet in target workbook
    Set targetWs = targetWb.Sheets("TJX")
    
    ' Find the data range
    lastRow = sourceWs.Cells(sourceWs.Rows.count, 1).End(xlUp).row
    
    ' Clear existing data in target (except headers)
    targetWs.Range("A3:U" & targetWs.Cells(targetWs.Rows.count, 1).End(xlUp).row).ClearContents
    
    ' Copy the entire data range (matching Python's approach)
    
    If lastRow > 2 Then
        targetWs.Range("A3:u" & lastRow).value = sourceWs.Range("A3:u" & lastRow).value
    End If
    
    If lastRow > 2 Then
        sourceWs.Range("c3:c" & lastRow).Copy
        targetWs.Range("c3").PasteSpecial Paste:=xlPasteFormulas
        Application.CutCopyMode = False
    End If
    
    MsgBox "TJX data updated successfully from TickerFile.xlsx" & vbCrLf & _
           "Rows processed: " & (lastRow - 2), vbInformation
    
CleanUp:
    ' Close TickerFile without saving
    If Not sourceWb Is Nothing Then
        sourceWb.Close SaveChanges:=False
    End If
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    GoTo CleanUp
    
End Sub

Public Function GetSectorDistribution() As String
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    Dim sectorCounts As Object
    Set sectorCounts = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = 3 To lastRow
        Dim sectorValue As String
        sectorValue = Trim(CStr(ws.Cells(i, 8).value))
        
        If sectorValue <> "" And sectorValue <> "N/A" And sectorValue <> "Error" Then
            If sectorCounts.Exists(sectorValue) Then
                sectorCounts(sectorValue) = sectorCounts(sectorValue) + 1
            Else
                sectorCounts.Add sectorValue, 1
            End If
        End If
    Next i
    
    Dim result As String
    result = "Sector Distribution:" & vbCrLf & String(20, "-") & vbCrLf
    
    Dim sector As Variant
    Dim totalStocks As Long
    totalStocks = 0
    
    ' Calculate total stocks first
    For Each sector In sectorCounts.Keys
        totalStocks = totalStocks + sectorCounts(sector)
    Next sector
    
    ' Build the result string with counts and percentages
    For Each sector In sectorCounts.Keys
        Dim count As Long
        Dim percentage As Double
        count = sectorCounts(sector)
        percentage = (count / totalStocks) * 100
        
        result = result & sector & ": " & count & " (" & Format(percentage, "0.0") & "%)" & vbCrLf
    Next sector
    
    ' Add summary line
    result = result & vbCrLf & "Total Stocks: " & totalStocks
    
    GetSectorDistribution = result
End Function
