Attribute VB_Name = "XchgeRates"
#If VBA7 Then
Private Declare PtrSafe Function VarPtrArray Lib "VBA" Alias "VarPtr" (var() As Any) As LongPtr
#Else
Private Declare Function VarPtrArray Lib "VBA" Alias "VarPtr" (var() As Any) As Long
#End If

Sub UpdateExchangeRates()
    Dim http As Object
    Dim json As Object
    Dim wsRates As Worksheet
    Dim url As String
    Dim exchangeRate As Variant
    
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
    wsRates.Cells(1, 3).value = "ToCAD"
    
    ' Example static rates (replace with actual rates as needed)
    Dim exchanges As Variant
    exchanges = Array("XASX", "XMIL", "XCNQ", "XFRA", "XAMS", "XSTC", "XLON", "XNAS", "XNSE", "NEOE", "XNYS", "ARCX", "OTCM", "XTSE", "XTSX", "XETR")
    
    For i = 0 To UBound(exchanges)
        wsRates.Cells(i + 2, 1).value = exchanges(i)
        wsRates.Cells(i + 2, 2).value = 1 ' Placeholder for USD conversion rates
        wsRates.Cells(i + 2, 3).value = json("observations")(1)("d")("FXUSDCAD")("v") ' CAD conversion rate
    Next i
    
    ' Add the USD to CAD rate
    wsRates.Cells(2, 1).value = "USDCAD"
    wsRates.Cells(2, 2).value = 1
    wsRates.Cells(2, 3).value = json("observations")(1)("d")("FXUSDCAD")("v")
    
    MsgBox "Exchange rates updated!", vbInformation
End Sub

Sub ConvertStockPrices()
    Dim ws As Worksheet, wsRates As Worksheet
    Dim lastRow As Long, i As Long
    Dim exchange As String, ticker As String
    Dim price As Double, toUSD As Double, toCAD As Double

    Set ws = ThisWorkbook.Sheets("TJX")
    Set wsRates = ThisWorkbook.Sheets("ConversionRates")

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    For i = 3 To lastRow ' Assuming headers in the first row
        ' Extract exchange from "Stock" column
        price = ws.Cells(i, 4).value
        exchange = ws.Cells(i, 6).value
        
        ' Lookup conversion rates
        toUSD = Application.WorksheetFunction.VLookup(exchange, wsRates.Range("A:C"), 2, False)
        toCAD = Application.WorksheetFunction.VLookup(exchange, wsRates.Range("A:C"), 3, False)

        ' Convert and write the converted prices to adjacent columns
        ws.Cells(i, 7).value = price * toUSD ' Column E for USD
        ws.Cells(i, 8).value = price * toCAD ' Column F for CAD
    Next i

    MsgBox "Conversion completed!", vbInformation
End Sub

