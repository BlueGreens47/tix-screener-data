Attribute VB_Name = "RemoveModules"
' ============================================================
' RemoveModules.bas
' PURPOSE: Remove redundant, stub, and legacy VBA modules
'          from TRADE.xlsm to clean up the VBA project.
'
' HOW TO USE:
'   1. Open TRADE.xlsm in Excel
'   2. Press Alt+F11 to open the VBA Editor (VBE)
'   3. From the VBE menu: File > Import File...
'   4. Select THIS file (RemoveModules.bas) and click Open
'   5. Press F5 or click Run > Run Sub to run RemoveRedundantModules
'   6. Click OK on the confirmation message
'   7. The workbook will be saved automatically
'   8. After completion, delete this module (right-click > Remove)
'
' MODULES REMOVED (18 total):
'   Batch 1 - Hollow stubs (ho* prefix):
'     hoCompleteFilterandReport, hoCompleteCoreProcessing,
'     hoCompleteTrading, hoCompleteWealthEnjoyment,
'     hoFiltering, hoJsonConverter, hoPerformance, hoValidateData
'   Batch 2 - Legacy originals superseded by newer versions:
'     OrigFilteringfromLess, OrigIndicators, origPerformance
'   Batch 3 - ATR and filtering variants from other workbooks:
'     ATRCalculationFromLess, otherATRCalculating,
'     FromOtherATRCalculating, FilteringfromNewBare
'   Batch 4 - Duplicate signal generator with column mapping bug:
'     GenerateSignals
'   Batch 5 - Consolidated duplicate of ALL.bas:
'     AllOne
'   Batch 6 - Broken stub (hardcoded to wrong workbook):
'     Performance
' ============================================================

Option Explicit

Sub RemoveRedundantModules()

    Dim wb As Workbook
    Dim vbProj As Object
    Dim vbComp As Object
    Dim modulesToRemove(17) As String
    Dim i As Integer
    Dim removedCount As Integer
    Dim notFoundList As String
    Dim removedList As String
    Dim msg As String

    ' ---- Define the 18 modules to remove ----
    ' Batch 1: Hollow stubs (ho* prefix)
    modulesToRemove(0) = "hoCompleteFilterandReport"
    modulesToRemove(1) = "hoCompleteCoreProcessing"
    modulesToRemove(2) = "hoCompleteTrading"
    modulesToRemove(3) = "hoCompleteWealthEnjoyment"
    modulesToRemove(4) = "hoFiltering"
    modulesToRemove(5) = "hoJsonConverter"
    modulesToRemove(6) = "hoPerformance"
    modulesToRemove(7) = "hoValidateData"

    ' Batch 2: Legacy originals
    modulesToRemove(8) = "OrigFilteringfromLess"
    modulesToRemove(9) = "OrigIndicators"
    modulesToRemove(10) = "origPerformance"

    ' Batch 3: ATR and filtering variants from other workbooks
    modulesToRemove(11) = "ATRCalculationFromLess"
    modulesToRemove(12) = "otherATRCalculating"
    modulesToRemove(13) = "FromOtherATRCalculating"
    modulesToRemove(14) = "FilteringfromNewBare"

    ' Batch 4: Duplicate signal generator with column mapping bug
    modulesToRemove(15) = "GenerateSignals"

    ' Batch 5: Consolidated duplicate of ALL.bas
    modulesToRemove(16) = "AllOne"

    ' Batch 6: Broken stub (hardcoded to wrong workbook)
    modulesToRemove(17) = "Performance"

    ' ---- Confirm before proceeding ----
    msg = "This will remove 18 redundant VBA modules from:" & vbCrLf & _
          ThisWorkbook.FullName & vbCrLf & vbCrLf & _
          "Modules to remove:" & vbCrLf
    For i = 0 To UBound(modulesToRemove)
        msg = msg & "  - " & modulesToRemove(i) & vbCrLf
    Next i
    msg = msg & vbCrLf & "The workbook will be saved after cleanup." & vbCrLf & _
          "Continue?"

    If MsgBox(msg, vbYesNo + vbQuestion, "Remove Redundant Modules") = vbNo Then
        MsgBox "Operation cancelled. No modules were removed.", vbInformation
        Exit Sub
    End If

    ' ---- Get reference to this workbook's VBA project ----
    Set wb = ThisWorkbook
    Set vbProj = wb.VBProject

    removedCount = 0
    removedList = ""
    notFoundList = ""

    ' ---- Loop and remove each module ----
    For i = 0 To UBound(modulesToRemove)
        Dim targetName As String
        targetName = modulesToRemove(i)

        ' Search for the component by name
        Dim found As Boolean
        found = False

        Dim comp As Object
        For Each comp In vbProj.VBComponents
            If comp.Name = targetName Then
                ' Only remove standard modules (Type 1) and class modules (Type 2)
                ' Do NOT remove worksheets (Type 100) or ThisWorkbook (Type 100)
                If comp.Type = 1 Or comp.Type = 2 Then
                    vbProj.VBComponents.Remove comp
                    removedList = removedList & "  [OK] " & targetName & vbCrLf
                    removedCount = removedCount + 1
                    found = True
                Else
                    notFoundList = notFoundList & "  [SKIP - not a module] " & targetName & vbCrLf
                    found = True
                End If
                Exit For
            End If
        Next comp

        If Not found Then
            notFoundList = notFoundList & "  [NOT FOUND] " & targetName & vbCrLf
        End If
    Next i

    ' ---- Save the workbook ----
    If removedCount > 0 Then
        Application.DisplayAlerts = False
        wb.Save
        Application.DisplayAlerts = True
    End If

    ' ---- Show results ----
    Dim resultMsg As String
    resultMsg = "Module cleanup complete!" & vbCrLf & vbCrLf
    resultMsg = resultMsg & "Removed: " & removedCount & " module(s)" & vbCrLf

    If Len(removedList) > 0 Then
        resultMsg = resultMsg & vbCrLf & "Successfully removed:" & vbCrLf & removedList
    End If

    If Len(notFoundList) > 0 Then
        resultMsg = resultMsg & vbCrLf & "Notes:" & vbCrLf & notFoundList
    End If

    resultMsg = resultMsg & vbCrLf & "The workbook has been saved." & vbCrLf & vbCrLf & _
                "NEXT STEP: Right-click this module (RemoveModules) in the" & vbCrLf & _
                "Project Explorer and select 'Remove RemoveModules' to delete" & vbCrLf & _
                "this cleanup script from the workbook."

    MsgBox resultMsg, vbInformation, "Cleanup Complete"

End Sub

' ---- Helper: List all current module names (for verification) ----
Sub ListAllModules()
    Dim vbProj As Object
    Dim comp As Object
    Dim moduleList As String
    Dim count As Integer

    Set vbProj = ThisWorkbook.VBProject
    count = 0
    moduleList = "VBA Modules in " & ThisWorkbook.Name & ":" & vbCrLf & vbCrLf

    For Each comp In vbProj.VBComponents
        If comp.Type = 1 Or comp.Type = 2 Then  ' Standard or class modules only
            moduleList = moduleList & "  " & comp.Name & vbCrLf
            count = count + 1
        End If
    Next comp

    moduleList = moduleList & vbCrLf & "Total: " & count & " module(s)"
    MsgBox moduleList, vbInformation, "Module List"
End Sub
