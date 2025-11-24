Attribute VB_Name = "Clear"
Option Explicit

' ==========================
' Configuration
' ==========================
Private Const WS_KOV As String = "KOV"
Private Const WS_KOV_MULTI As String = "KOV Multi"
Private Const WS_BATCH As String = "Batch Summary"
Private Const WS_PASTE As String = "Paste Data"
Private Const WS_GRAPHS As String = "Graphs"
Private Const WS_OVERLAYS As String = "Overlays"
Private Const WS_SCRATCH As String = "Scratch"

' Product list anchor
Private Const PRODUCT_LIST_WS As String = "Batch Summary"
Private Const PRODUCT_LIST_ANCHOR As String = "P1"
Private Const DROPDOWN_LAST_ROW As Long = 1000

' Batch Summary layout
Private Const COL_PRODUCT As String = "G"              ' dropdown column
Private Const HEADER_RANGE As String = "A1:G1"         ' header span
Private Const CLEAR_RANGE As String = "A2:G1048576"    ' rows to cin lear (leaves col P intact)

' =====================================
' One-Button Clear Everything (Main)
' =====================================
Public Sub KOV_Clear_All()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim oldCalc As XlCalculation: oldCalc = Application.Calculation
    Dim oldStatus As Variant: oldStatus = Application.StatusBar
    Dim oldShowStatus As Boolean: oldShowStatus = Application.DisplayStatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationManual
    On Error GoTo FINALLY

    SetStatus "Clearing: KOV sheets…"
    ClearOrCreate_KOV wb
    ClearOrCreate_KOVMulti wb

    SetStatus "Resetting Product_List name…"
    Reset_Product_List_Name wb
    
    SetStatus "Clearing: Batch Summary…"
    Reset_BatchSummary wb

    SetStatus "Clearing: Paste Data…"
    Clear_PasteData wb

    SetStatus "Clearing: Graphs & Overlays…"
    Clear_Sheet_ChartsOrContent wb, WS_GRAPHS, "Graphs cleared."
    Clear_Sheet_ChartsOrContent wb, WS_OVERLAYS, "Overlays cleared."

    SetStatus "Clearing: Scratch…"
    Clear_Scratch wb, WS_SCRATCH

    SetStatus "Resetting window flags…"
    Reset_KOV_WindowFlags

FINALLY:
    Application.DisplayAlerts = True
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = oldShowStatus
    Application.StatusBar = oldStatus
End Sub

' ----------------------------- Status bar helper -----------------------------
Private Sub SetStatus(ByVal msg As String)
    Static lastTick As Single
    If Timer - lastTick > 0.15 Then
        Application.StatusBar = msg
        DoEvents
        lastTick = Timer
    End If
End Sub

' ============================== Helpers ==============================

Private Sub ClearOrCreate_KOV(wb As Workbook)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(wb, WS_KOV)
    ClearWorksheetCompletely ws, "Select a product and run KOV."
    ws.Columns("A:L").ColumnWidth = 14
End Sub

Private Sub ClearOrCreate_KOVMulti(wb As Workbook)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(wb, WS_KOV_MULTI)
    ClearWorksheetCompletely ws, "Consolidated KOV (Week)"
    ws.Columns("A:L").ColumnWidth = 14
End Sub

Private Sub Reset_BatchSummary(wb As Workbook)
    Dim ws As Worksheet, rngHdr As Range
    Set ws = GetOrCreateSheet(wb, WS_BATCH)

    ' Clear only what exists
    Dim last As Long
    last = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    If last < 2 Then last = 2
    ws.Range("A2:G" & last).Clear
    
    ' Safer header/CF clearing
    ws.Range(HEADER_RANGE).ClearFormats
    If Not ws.UsedRange Is Nothing Then ws.UsedRange.FormatConditions.Delete

    ' Rebuild headers
    Set rngHdr = ws.Range(HEADER_RANGE)
    rngHdr.value = Array("Tag", "Batch Start", "Batch End", _
                         "Duration (min)", "Duration (hr)", "Status", "Product")
    rngHdr.Font.Bold = True

    ' Formats
    ws.Columns("A:G").ColumnWidth = 18
    ws.Columns("D:E").NumberFormat = "0.00"
    ws.Columns("B:C").NumberFormat = "m/d/yyyy h:mm"

    ' Apply dropdowns to a fixed range, e.g., 2..1000
    With ws.Range(COL_PRODUCT & "2:" & COL_PRODUCT & DROPDOWN_LAST_ROW).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=Product_List"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Product"
        .InputMessage = "Choose a product; leave blank to skip this batch."
        .ErrorTitle = "Pick from list"
        .ErrorMessage = "Use the dropdown list of products."
    End With
End Sub

Private Sub Clear_PasteData(wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(WS_PASTE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ClearWorksheetCompletely ws, ""
End Sub

' Rename + point the Product_List name to your new anchored spill (Batch Summary!P1#)
Private Sub Reset_Product_List_Name(wb As Workbook)
    Dim anchorRef As String
    anchorRef = "='" & PRODUCT_LIST_WS & "'!" & _
                wb.Worksheets(PRODUCT_LIST_WS).Range(PRODUCT_LIST_ANCHOR).Address & "#"

    On Error Resume Next
    wb.Names("Product_List").Delete
    On Error GoTo 0

    wb.Names.Add name:="Product_List", RefersTo:=anchorRef
End Sub

' Clear Scratch helper sheet (if present) without deleting the sheet
Private Sub Clear_Scratch(wb As Workbook, ByVal nm As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ClearWorksheetCompletely ws, ""
    ws.Visible = xlSheetVeryHidden
End Sub

Private Sub ClearWorksheetCompletely(ws As Worksheet, Optional topLeftMsg As String = "")
    Dim i As Long
    Dim wasProtected As Boolean

    On Error Resume Next
    wasProtected = ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios
    If wasProtected Then ws.Unprotect    ' add password if needed
    On Error GoTo 0

    ' Huge speed-ups: avoid extra layout work during deletes
    On Error Resume Next
    ws.DisplayPageBreaks = False
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0

    ' 1) Kill pivots first (they can block/slow clearing)
    On Error Resume Next
    For i = ws.PivotTables.Count To 1 Step -1
        ws.PivotTables(i).TableRange2.Clear
    Next i
    On Error GoTo 0

    ' 2) Unlist tables so clearing is simple (use .Delete if you truly want to remove them)
    On Error Resume Next
    For i = ws.ListObjects.Count To 1 Step -1
        ws.ListObjects(i).Unlist
    Next i
    On Error GoTo 0

    ' 3) Clear Conditional Formats on UsedRange (not entire sheet — much faster)
    On Error Resume Next
    If Not ws.UsedRange Is Nothing Then ws.UsedRange.FormatConditions.Delete
    On Error GoTo 0

    ' 4) Clear UsedRange (values + formats) instead of full .Cells (billions of cells)
    On Error Resume Next
    If Not ws.UsedRange Is Nothing Then ws.UsedRange.Clear
    On Error GoTo 0

    ' 5) Delete charts & shapes **backwards** so the collection doesn’t mutate mid-loop
    On Error Resume Next
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    On Error GoTo 0

    ' 6) Optional A1 message
    If Len(topLeftMsg) > 0 Then ws.Range("A1").value = topLeftMsg

    ' Re-protect if needed
    On Error Resume Next
    If wasProtected Then ws.Protect
    On Error GoTo 0
End Sub

Private Sub DeleteCharts(ws As Worksheet)
    Dim co As ChartObject
    On Error Resume Next
    For Each co In ws.ChartObjects
        co.Delete
    Next co
    On Error GoTo 0
End Sub

Private Function GetOrCreateSheet(wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.name = nm
    End If
    Set GetOrCreateSheet = ws
End Function

Private Sub Reset_KOV_WindowFlags()
    On Error Resume Next
    G_KOV_UseWindow = False
    G_KOV_WindowStart = 0#
    G_KOV_WindowEnd = 0#
    On Error GoTo 0
End Sub

' Clear a target named sheet whether it exists as a Worksheet or a Chart sheet.
' If it is a Chart sheet, delete and recreate a blank Worksheet with the same name.
Private Sub Clear_Sheet_ChartsOrContent(wb As Workbook, ByVal targetName As String, ByVal afterMsg As String)
    Dim ws As Worksheet
    Dim ch As Chart

    ' Case A: worksheet exists
    On Error Resume Next
    Set ws = wb.Worksheets(targetName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        ClearWorksheetCompletely ws, afterMsg
        Exit Sub
    End If

    ' Case B: chart sheet exists
    On Error Resume Next
    Set ch = wb.Charts(targetName)
    On Error GoTo 0
    If Not ch Is Nothing Then
        Application.DisplayAlerts = False
        ch.Delete
        Application.DisplayAlerts = True

        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.name = targetName
        If Len(afterMsg) > 0 Then ws.Range("A1").value = afterMsg
        Exit Sub
    End If

    ' Case C: neither exists -> create a fresh worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.name = targetName
    If Len(afterMsg) > 0 Then ws.Range("A1").value = afterMsg
End Sub

' =========================
' Call this AFTER batch-detection macro populates rows
' =========================
Public Sub Refresh_Product_Dropdowns()
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Worksheets(WS_BATCH)

    ' Determine last used row based on Column A (Tag). Adjust if you prefer a different column.
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then lastRow = 2

    With ws.Range(COL_PRODUCT & "2:" & COL_PRODUCT & lastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=Product_List"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Product"
        .InputMessage = "Choose a product; leave blank to skip this batch."
        .ErrorTitle = "Pick from list"
        .ErrorMessage = "Use the dropdown list of products."
    End With
End Sub


