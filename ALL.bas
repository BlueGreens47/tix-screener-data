Attribute VB_Name = "ALL"
Option Explicit
' Global variables
Public gStopMacro As Boolean
Public pubNotice As Boolean
Public minScore As Long
Public perfTest As Boolean
Public endDate As Date
Public sigTest As Boolean

 'Global dictionary to store exchange rates
Private ExchangeRates As Scripting.dictionary

' E-Stop
Sub StopButton_Click()
    gStopMacro = True
    MsgBox "E-Stop!", vbOKOnly
    ResetApplicationSettings
    End
End Sub

' Main Processing Subroutine
Sub ProcessAll()
    Dim startTime As Double
    Dim processResult As Boolean
    
    On Error GoTo ErrorHandler
    startTime = Timer
    
    'If Not ConfirmProcessing() Then Exit Sub
    
    ClearAllFilters
    
    ' Disable application events for performance
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ' Process stocks
    processResult = ProcessAllStocks()
    
    ' Finalize processing
    If processResult Then
        FinalizeStockProcessing
        DisplayCompletionMessage startTime
        ThisWorkbook.Save
    End If

Cleanup:
    Call UpdateBackupAll
    'Call UploadToDrive
    
    Exit Sub
    
ErrorHandler:
    HandleProcessingError "ProcessAll", Err
    Resume Cleanup
End Sub

' User Confirmation Function
Public Function ConfirmProcessing() As Boolean
    If Not pubNotice Then
        Dim userResponse As VbMsgBoxResult
        userResponse = MsgBox("Confirm processing?" & vbCrLf & "'No' to exit.", vbYesNo)
        ConfirmProcessing = (userResponse = vbYes)
    Else
        ConfirmProcessing = True
    End If
End Function

' Process All Stocks
Private Function ProcessAllStocks() As Boolean
    Dim wsTJX As Worksheet
    Dim HistPage As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim currentSymbol As String
    Dim processedSymbols As Long
    
   ' On Error GoTo ErrorHandler
    
    Set wsTJX = ThisWorkbook.Worksheets("TJX")
    Set HistPage = ThisWorkbook.Worksheets("DataHistory")
        
    ' Prepare history sheet
    PrepareHistorySheet HistPage
    
    wsTJX.AutoFilterMode = False
    lastRow = wsTJX.Cells(wsTJX.Rows.count, "E").End(xlUp).row
    
    For i = 3 To lastRow Step 1
        ' Check for user interruption
        If gStopMacro Then
            MsgBox "E-Stopped by user.", vbInformation
            ProcessAllStocks = (processedSymbols > 0)
            Exit Function
        End If
        
        currentSymbol = Trim(wsTJX.Range("E" & i).Value)
                
        If ProcessSymbolData(currentSymbol, HistPage) Then
            processedSymbols = processedSymbols + 1
        End If
    Next i
    
    ProcessAllStocks = (processedSymbols > 0)
    
'ErrorHandler:
    'HandleProcessingError "ProcessAllStocks", Err
    'ProcessAllStocks = False
    
End Function

' Prepare History Sheet
Private Sub PrepareHistorySheet(ByRef HistPage As Worksheet)
    HistPage.Range("A1:G" & HistPage.Rows.count).ClearContents
  
    dataHeader HistPage
    With HistPage.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
        
    End With
    
    ' Ensure column widths are appropriate
    HistPage.Range("A:A").NumberFormat = "m/d/yyyy"
    HistPage.Columns("A:G").AutoFit
    HistPage.Range("G:G").NumberFormat = "@"
    
End Sub

' Process Individual Symbol Data
Private Function ProcessSymbolData(Symbol As String, HistPage As Worksheet) As Boolean
    Dim processResult As Boolean
    Dim dendDate As Date, dstartDate As Date
    Dim endDate As String, startDate As String
   
    ' Initialize to false
    processResult = False
    
    dendDate = GetPreviousWorkday(Date)
       
    dstartDate = DateAdd("d", -7, dendDate) '-7, -28, -366
          
    endDate = Format(dendDate, "yyyy-mm-dd")
    startDate = Format(dstartDate, "yyyy-mm-dd")
    
    If ValidateStockInputs(Symbol, startDate, endDate) Then
        processResult = RetrieveStockHistory(Symbol, startDate, endDate, HistPage)
    End If
    
    ProcessSymbolData = processResult
    
End Function

' Validate Stock Inputs
Private Function ValidateStockInputs(Symbol As String, startDate As String, endDate As String) As Boolean
    ValidateStockInputs = (Len(Symbol) > 0 And IsDate(CDate(startDate)) And IsDate(CDate(endDate)))
    
    If Not ValidateStockInputs Then
        MsgBox "Invalid input for symbol: " & Symbol & ". Please check the symbol and dates.", vbExclamation
    End If
    
End Function

' Retrieve Stock History
Private Function RetrieveStockHistory(Symbol As String, startDate As String, endDate As String, HistPage As Worksheet) As Boolean
    Dim lastRow As Long
    Dim stockDataRange As Range
    Dim numRows As Long
    Dim ticker As String
    Dim colonPos As Integer
    
    HistPage.AutoFilterMode = False
    
        ' Find the last empty row in DataHistory
    lastRow = HistPage.Cells(HistPage.Rows.count, "A").End(xlUp).row + 1
    
    ' Write the STOCKHISTORY formula
    Set stockDataRange = HistPage.Range("A" & lastRow)
    stockDataRange.Formula2R1C1 = "=STOCKHISTORY(""" & Symbol & """, """ & startDate & """, """ & endDate & """,0,0,0,2,3,4,1,5)"
    
    ' Wait for the STOCKHISTORY function to complete
    Application.Wait Now + TimeValue("00:00:02")
    HistPage.Calculate

    ' Check if data was successfully retrieved
    If IsError(HistPage.Cells(lastRow, 1)) Then
        HistPage.Rows(lastRow).ClearContents
        RetrieveStockHistory = False
    Else
        ' Process retrieved data
        numRows = HistPage.Cells(HistPage.Rows.count, "A").End(xlUp).row - lastRow + 1
        If numRows > 0 Then
            colonPos = InStr(Symbol, ":")
            ticker = Trim(Mid(Symbol, colonPos + 1))
            HistPage.Range("G" & lastRow & ":G" & lastRow + numRows - 1).Value = ticker
            ConvertFormulasToValues HistPage
            RetrieveStockHistory = True
        End If
    End If
    
End Function

' Finalize Stock Processing
Private Sub FinalizeStockProcessing()
    Dim HistPage As Worksheet
    Set HistPage = ThisWorkbook.Worksheets("DataHistory")
    
    Call DeleteNARows(HistPage)
    Call BackupALLAndSort
    'Call myConversionRates
    
    If pubNotice Then Call FilterAndReport
    
End Sub
Sub DisplayCompletionMessage(startTime As Double)
    Dim endTime As Double, Finish As Double
    If Not pubNotice Or Not perfTest Then
        Finish = Timer
        MsgBox "Task Completed in: " & Format(Finish - startTime, "0.0") & " seconds,  " & Format((Finish - startTime) / 60, "0.0") & " minutes.", vbExclamation
    End If
End Sub

' Error Handling
Public Sub HandleProcessingError(ByVal procName As String, ByVal ErrorObj As ErrObject)
    Dim ErrorMessage As String

    ' Build the error message with the provided procedure name
    ErrorMessage = "An error occurred in the " & procName & " procedure:" & vbNewLine & _
                   "Error Number: " & ErrorObj.Number & vbNewLine & _
                   "Description: " & ErrorObj.Description & vbNewLine & _
                   "Source: " & ErrorObj.Source
    ' Log the error
    Debug.Print ErrorMessage

    ' Set global stop flag to halt further processing
    gStopMacro = True
    
End Sub

Sub ConvertFormulasToValues(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    
    With ws
        lastRow = .Cells(.Rows.count, "A").End(xlUp).row
        lastCol = .Cells(1, .Columns.count).End(xlToLeft).Column
        .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Value = .Range(.Cells(1, 1), .Cells(lastRow, lastCol)).Value
        .Range("G" & lastRow).NumberFormat = "@"
        .Columns("A:G").AutoFit
    End With
End Sub

Sub WrapWithIFERROR()
Attribute WrapWithIFERROR.VB_ProcData.VB_Invoke_Func = "e\n14"
    Dim cell As Range
    For Each cell In Selection
        If cell.HasFormula Then
            If Not InStr(1, cell.Formula, "IFERROR", vbTextCompare) > 0 Then
                cell.Formula = "=IFERROR(" & Mid(cell.Formula, 2) & ","""")"
            End If
        End If
    Next cell
    
End Sub

Sub BackupALLAndSort()
    Dim wsHistory As Worksheet, wsBackup As Worksheet
    Dim lastRowHistory As Long, lastRowBackup As Long
    Dim combinedRange As Range
    
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False ' Optional for performance
    
    Set wsHistory = ThisWorkbook.Sheets("DataHistory")
    Set wsBackup = ThisWorkbook.Sheets("BackupAll")
        
    wsBackup.Range("A:A").NumberFormat = "yyyy-mm-dd"
    wsBackup.Range("G:G").NumberFormat = "@"
    
    lastRowHistory = wsHistory.Cells(wsHistory.Rows.count, "A").End(xlUp).row
    If lastRowHistory < 2 Then Exit Sub ' No data to copy
    
    lastRowBackup = wsBackup.Cells(wsBackup.Rows.count, "A").End(xlUp).row
    wsHistory.Range("A2:G" & lastRowHistory).Copy wsBackup.Range("A" & lastRowBackup + 1)
    
    lastRowBackup = wsBackup.Cells(wsBackup.Rows.count, "A").End(xlUp).row
    If lastRowBackup <= 1 Then Exit Sub ' No new data
    
    Set combinedRange = wsBackup.Range("A1:G" & lastRowBackup)
    
    ' Sort by Date (A) and Ticker (G)
    With wsBackup.Sort
        .SortFields.Clear
        .SortFields.Add key:=combinedRange.Columns(1), Order:=xlAscending  ' Date
        .SortFields.Add key:=combinedRange.Columns(7), Order:=xlAscending  ' Ticker
        .SetRange combinedRange
        .Header = xlYes
        .Apply
    End With
    
    ' Remove duplicates (Date + Ticker)
    combinedRange.RemoveDuplicates Columns:=Array(1, 7), Header:=xlYes
    
    ' Cleanup N/A rows
    DeleteNARows wsBackup

ErrorHandler:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then HandleProcessingError "BackupAllAndSort", Err
    
End Sub

Sub OneStock()
    Dim stockRange As Range
    Dim Symbol As String, startDate As String, endDate As String
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Symbol = Trim(ThisWorkbook.Sheets("DashBoard").Range("A8").Value)
    
    If (MsgBox("Get data for " & Symbol, vbYesNo) = vbNo) Then End

    endDate = GetPreviousWorkday(Date)
    startDate = GetPreviousWorkday(Date - 366)
          
    ' Create sheet for the Symbol if it doesn't exist
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Symbol)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("DashBoard"))
        ws.Name = Symbol
    End If
    On Error GoTo 0

    ' Clear existing data
    ws.Cells.Clear

    ' Prepare new Sheet
    PrepareHistorySheet ws
    
    Symbol = Trim(ThisWorkbook.Sheets("DashBoard").Range("A8").Value)   ' Fixed: was AK8 (wrong column)
    Symbol = Application.WorksheetFunction.VLookup(Symbol, ThisWorkbook.Sheets("TJX").Range("A:E"), 5, False)
    
    If ValidateStockInputs(Symbol, startDate, endDate) Then
         RetrieveStockHistory Symbol, startDate, endDate, ws
    End If
    
   MsgBox "Done, suggested tickerBackup asap.", vbInformation

ErrorHandler:
    HandleProcessingError "OneStock", Err
    
End Sub

Sub tickerBackup()

Dim lastFromRow As Long, lastToRow As Long
Dim period As Integer
Dim endDate As Date, startDate As Date
Dim ticker As String
Dim visibleRange As Range, combinedRange As Range
Dim wsDash As Worksheet, wsTo As Worksheet, wsFrom As Worksheet

On Error GoTo ErrorHandler
    ticker = ThisWorkbook.Sheets("DashBoard").Range("AF8")
  
    Set wsFrom = ThisWorkbook.Sheets(ticker)
    
    Call ConvertFormulasToValues(wsFrom)
    Call DeleteNARows(wsFrom)
    
    lastFromRow = wsFrom.Cells(wsFrom.Rows.count, "A").End(xlUp).row
    wsFrom.Range("G2:G" & lastFromRow).Value = ticker
    
    Application.CutCopyMode = False
    
    Set wsTo = ThisWorkbook.Sheets("BackupAll")
    lastToRow = wsTo.Cells(wsTo.Rows.count, "A").End(xlUp).row
    
     'Copy data
    wsFrom.Range("A2:G" & lastFromRow).Copy Destination:=wsTo.Range("A" & lastToRow + 1)

    ' Define the range to sort (including the headers)
    Set combinedRange = wsTo.Range("A1:G" & wsTo.Cells(wsTo.Rows.count, "A").End(xlUp).row)

    ' Sort the combined range by ticker (column G) and then by Date (column A)
    combinedRange.Sort key1:=wsTo.Range("G1"), Order1:=xlAscending, Key2:=wsTo.Range("A1"), Order2:=xlAscending, Header:=xlYes
    
    ' Remove duplicates in columns A and g
    combinedRange.RemoveDuplicates Columns:=Array(1, 7), Header:=xlYes
        
    Application.DisplayAlerts = False
    wsFrom.Delete
    Application.DisplayAlerts = True
ErrorHandler:
    HandleProcessingError "TickerBackup", Err
    
 End Sub
Sub SelectedHistorical() ' get uptodate Ticker Data
     
     ' Declare variables
    Dim DashPage As Worksheet
    Dim HistPage As Worksheet
    Dim Symbol As String
    Dim i As Long, lastRow As Long, lastHBrow, numRows As Long, lastCol As Long
    Dim stockDataRange As Range
    Dim dendDate As Date, dstartDate As Date
    Dim endDate As String, startDate As String
     
    ' Set worksheets
    Set DashPage = Sheets("DashBoard")
    Set HistPage = Sheets("DataHistory")
    
    dendDate = GetPreviousWorkday(Date) ' when adding new Stock
    dstartDate = DateAdd("d", -366, dendDate)
    
    endDate = Format(dendDate, "yyyy-mm-dd")
    startDate = Format(dstartDate, "yyyy-mm-dd")
        
    ' Clear contents of A2:H in Data while keeping headers intact
    HistPage.Range("A2:H" & HistPage.Rows.count).ClearContents

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With
    lastRow = DashPage.Cells(DashPage.Rows.count, "A").End(xlUp).row
    
    ' Loop through rows A8 to Value in DashPage
    For i = 8 To lastRow
        
        Symbol = Trim(DashPage.Range("A" & i).Value)
       ' Symbol = Application.WorksheetFunction.VLookup(Symbol, ThisWorkbook.Sheets("TJX").Range("A:E"), 5, False)
        
        If ValidateStockInputs(Symbol, startDate, endDate) Then
             RetrieveStockHistory Symbol, startDate, endDate, HistPage
        End If
        DoEvents
    Next i
    
    With HistPage
        .UsedRange.EntireColumn.AutoFit ' Auto-fit columns for better readability
        .Range("A1:G" & .Cells(.Rows.count, "A").End(xlUp).row).Sort key1:=.Range("G1"), Order1:=xlAscending, Key2:=.Range("A2"), Order2:=xlAscending, Header:=xlYes
    End With
    
    Call DeleteNARows(HistPage)
    Call BackupALLAndSort
    
    With Application
        .CutCopyMode = False
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
    End With
  
  ' Display completion message
    MsgBox "Download and Backup complete. ", vbExclamation
   
End Sub

Sub ClearAllFilters()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
        End If
    Next ws
End Sub

' OPTIMIZATION: Improved ResetApplicationSettings
Sub ResetApplicationSettings()
    ClearAllFilters
    With ThisWorkbook.Sheets("DashBoard")
        .Range("W5").Value = 3
        .Range("W6").Value = 0 'Reset Iterations
        .Range("Y5").Value = ThisWorkbook.Sheets("TJX").Range("E1").Value2
        .Range("Y6").Value = ThisWorkbook.Sheets("TJX").Range("C1").Value2
    End With
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    sigTest = False
    perfTest = False
    pubNotice = False
    gStopMacro = False
End Sub

Sub SaveAndClose()

    Dim wb As Workbook
    On Error Resume Next
    gStopMacro = False
    Application.DisplayAlerts = False
    pubNotice = False
    Set wb = ThisWorkbook
    wb.Save
    wb.Close SaveChanges:=True
    Application.Quit
    
End Sub


