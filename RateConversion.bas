Attribute VB_Name = "RateConversion"
Private Sub UpdateExchangeRates()  ' Private: canonical public version in XchgeRates.bas
    Dim http As Object
    Dim json As Object
    Dim wsRates As Worksheet
    Dim url As String
    Dim i As Integer
    
    ' Set the URL for the Bank of Canada's Valet API
    url = "https://www.bankofcanada.ca/valet/observations/FXUSDCAD/json"
    
    ' Create the HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send
    
    ' Parse the JSON response
    Set json = JsonConverter.ParseJson(http.responseText)
    
    ' Write the exchange rates to the "ConversionRates" sheet
    Set wsRates = ThisWorkbook.Sheets("ConversionRates")
    wsRates.Cells.ClearContents
    wsRates.Cells(1, 1).value = "Exchange"
    wsRates.Cells(1, 2).value = "ToUSD"
    'wsRates.Cells(1, 3).value = "ToCAD"
    
    ' Example static rates (replace with actual rates as needed)
    Dim exchanges As Variant
    exchanges = Array("XASX", "XMIL", "XCNQ", "XFRA", "XAMS", "XSTC", "XLON", "XNAS", "XNSE", "NEOE", "XNYS", "ARCX", "OTCM", "XTSE", "XTSX", "XETR")
    
    For i = 1 To UBound(exchanges)
        wsRates.Cells(i + 1, 1).value = exchanges(i)
        wsRates.Cells(i + 1, 2).value = 1 ' Placeholder for USD conversion rates
          Next i
        
    MsgBox "Exchange rates updated!", vbInformation
End Sub

Private Sub ConvertStockPrices()  ' Private: canonical public version in XchgeRates.bas
    Dim ws As Worksheet, wsRates As Worksheet
    Dim lastRow As Long, i As Long
    Dim exchange As String, ticker As String
    Dim price As Double, toUSD As Double, toCAD As Double

    Set ws = ThisWorkbook.Sheets("Sheet1") ' Adjust sheet name as necessary
    Set wsRates = ThisWorkbook.Sheets("ConversionRates")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    For i = 2 To lastRow ' Assuming headers in the first row
        ' Extract exchange from "Stock" column
        exchange = Split(Split(ws.Cells(i, 3).value, "(")(1), ":")(0)
        price = ws.Cells(i, 4).value

        ' Lookup conversion rates
        toUSD = Application.WorksheetFunction.VLookup(exchange, wsRates.Range("A:C"), 2, False)
        toCAD = Application.WorksheetFunction.VLookup(exchange, wsRates.Range("A:C"), 3, False)

        ' Convert and write the converted prices to adjacent columns
        ws.Cells(i, 5).value = price * toUSD ' Column E for USD
        ws.Cells(i, 6).value = price * toCAD ' Column F for CAD
    Next i

    MsgBox "Conversion completed!", vbInformation
End Sub

Sub ConvertStockPricesToUSD()
    Dim ws As Worksheet, wsRates As Worksheet
    Dim lastRow As Long, i As Long
    Dim stockInfo As String, exchange As String
    Dim originalPrice As Double, exchangeRate As Double
    
    Set ws = ThisWorkbook.Sheets("TJX")
    Set wsRates = ThisWorkbook.Sheets("ConvRates")
    exchangeRate = GetExchangeRate()
    
    If exchangeRate <= 0 Then
        MsgBox "Failed to retrieve exchange rate. Conversion aborted.", vbExclamation
        Exit Sub
    End If
    
    With ws
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        
        'Add USD Price header
        .Cells(2, 7) = "Price (USD)"
        
        For i = 3 To lastRow
            'stockInfo = .Cells(i, 6).value 'Stock info from column F
            
            'Extract exchange code
            'If InStr(1, stockInfo, ":") > 0 Then
             '  exchange = Split(Split(stockInfo, ":")(1), ":")(0)
            'Else
                exchange = .Cells(i, 6).value
            'End If
            
            'Get original price (remove currency symbol)
            originalPrice = Replace(.Cells(i, 4).value, "$ ", "")
            
            'Convert based on exchange
            Select Case UCase(exchange)
                Case "XNAS", "XNYS" 'US exchanges
                    .Cells(i, 7) = originalPrice
                Case "XTSX", "XTSE", "XCNQ" 'Canadian exchanges
                    .Cells(i, 7) = originalPrice * exchangeRate
                Case "XTSX", "XTSE" 'Canadian exchanges
                    .Cells(i, 7) = originalPrice * exchangeRate
                    
                Case Else
                    .Cells(i, 7) = "N/A"
            End Select
            
            'Format as currency
            .Cells(i, 7).NumberFormat = "$#,##0.00"
        Next i
    End With
    
    MsgBox "Conversion complete!", vbInformation
End Sub

Function GetExchangeRate() As Double
    Dim httpRequest As Object
    Dim responseText As String
    Dim cadPerUsd As Double
    Dim rateStart As Long, rateEnd As Long
    
    On Error GoTo errHandler
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    httpRequest.Open "GET", "https://www.bankofcanada.ca/valet/observations/FXUSDCAD/json?recent=1", False
    httpRequest.send
    responseText = httpRequest.responseText
    
    ' Parse response to find CAD/USD rate
    If InStr(responseText, "v"": """) > 0 Then
        rateStart = InStr(responseText, "v"": """) + Len("v"": """)
        rateEnd = InStr(rateStart, responseText, """")
        cadPerUsd = CDbl(Mid(responseText, rateStart, rateEnd - rateStart))
        GetExchangeRate = 1 / cadPerUsd
    Else
        MsgBox "Failed to parse exchange rate from API response.", vbExclamation
        GetExchangeRate = 0
    End If
    Exit Function
    
errHandler:
    MsgBox "Error retrieving exchange rate: " & Err.Description, vbCritical
    GetExchangeRate = 0
End Function
Function ExtractExchange(s As String) As String
    Dim openParenPos As Integer
    Dim colonPos As Integer

    openParenPos = InStr(s, "(") + 1
    colonPos = InStr(s, ":")

    ExtractExchange = Mid(s, openParenPos, colonPos - openParenPos)
End Function




