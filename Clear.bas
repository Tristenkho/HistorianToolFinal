Attribute VB_Name = "Clear"
Option Explicit

'==============================
' One-Button Clear Everything
'==============================
Public Sub KOV_Clear_All()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim oldCalc As XlCalculation: oldCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    On Error GoTo FINALLY

    ' 1) KOV + KOV Multi: wipe contents, formats, charts/shapes
    ClearOrCreate_KOV wb
    ClearOrCreate_KOVMulti wb

    ' 2) Batch Summary: reset headers + product dropdowns (G2:G100)
    Reset_BatchSummary wb

    ' 3) Paste Data: clear all and remove charts/shapes
    Clear_PasteData wb

    ' 3b) Clear the Graphs sheet (handles worksheet or chart-sheet)
    Clear_GraphsSheet wb

    ' 4) UI: reset product picker (B1), rebuild Product_List spill + named range,
    '        reapply validation, remove charts, clear form/ActiveX dropdown selections
    Reset_UI_Picker wb

    ' 5) Reset WeekRunner window flags
    Reset_KOV_WindowFlags

FINALLY:
    Application.DisplayAlerts = True
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'============================== Helpers ==============================

Private Sub ClearOrCreate_KOV(wb As Workbook)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(wb, "KOV")
    ClearWorksheetCompletely ws, "Select a product on UI and run KOV."
    ws.Columns("A:L").ColumnWidth = 14
End Sub

Private Sub ClearOrCreate_KOVMulti(wb As Workbook)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(wb, "KOV Multi")
    ClearWorksheetCompletely ws, "Consolidated KOV (Week)"
    ws.Columns("A:L").ColumnWidth = 14
End Sub

Private Sub Reset_BatchSummary(wb As Workbook)
    Dim ws As Worksheet, rngHdr As Range
    Set ws = GetOrCreateSheet(wb, "Batch Summary")

    With ws
        .Cells.Clear
        .Cells.FormatConditions.Delete
        .Cells.Interior.Pattern = xlNone
        .Cells.Borders.LineStyle = xlNone

        Set rngHdr = .Range("A1:G1")
        rngHdr.value = Array("Tag", "Batch Start", "Batch End", "Duration (min)", "Duration (hr)", "Status", "Product")
        rngHdr.Font.Bold = True

        .Columns("A:G").ColumnWidth = 18
        .Columns("D:E").NumberFormat = "0.00"
        .Columns("B:C").NumberFormat = "m/d/yyyy h:mm"

        ' Reapply dropdown on Product column if Product_List exists (we rebuild it in Reset_UI_Picker)
        With .Range("G2:G100").Validation
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
    End With
End Sub

Private Sub Clear_PasteData(wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("Paste Data")
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0

    ClearWorksheetCompletely ws, ""
End Sub

Private Sub Reset_UI_Picker(wb As Workbook)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(wb, "UI")

    ' Remove only charts; keep other shapes unless you want them gone too
    DeleteCharts ws

    ' Rebuild product spill list in F1 and (re)create named range Product_List
    On Error Resume Next: wb.Names("Product_List").Delete: On Error GoTo 0

    ' Try tblLimits first; fall back to raw "Product Limits" col A if table missing
    On Error Resume Next
    ws.Range("F1").Formula2 = "=IFERROR(SORT(UNIQUE(FILTER(tblLimits[Product],tblLimits[Product]<>""""))),"""")"
    If Len(CStr(ws.Range("F1").Value2)) = 0 Then
        ws.Range("F1").Formula2 = "=LET(src,'Product Limits'!A2:A100000," & _
                                  "IFERROR(SORT(UNIQUE(FILTER(src,src<>""""))),""""))"
    End If
    On Error GoTo 0

    ThisWorkbook.Names.Add name:="Product_List", RefersTo:="=" & ws.name & "!$F$1#"
    ws.Columns("F").Hidden = True

    ' Clear selection in B1 and reapply validation tied to Product_List
    With ws.Range("B1")
        .ClearContents
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="=Product_List"
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = "Select Product"
            .InputMessage = "Choose a product to run KOV."
            .ErrorTitle = "Invalid selection"
            .ErrorMessage = "Pick a product from the list."
        End With
        .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
    End With

    ' If you have Form Controls / ActiveX dropdowns, clear their linked cells/selections
    Clear_Ui_FormAndActiveX_Dropdowns ws
End Sub

Private Sub Clear_Ui_FormAndActiveX_Dropdowns(ws As Worksheet)
    Dim shp As Shape, lc As String
    Dim ole As OLEObject

    ' Forms controls (e.g., Drop Down from Form Controls)
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            On Error Resume Next
            lc = shp.ControlFormat.LinkedCell
            If Len(lc) > 0 Then ws.Parent.Worksheets(ws.name).Range(lc).ClearContents
            On Error GoTo 0
        End If
    Next shp

    ' ActiveX controls (e.g., ComboBox)
    For Each ole In ws.OLEObjects
        On Error Resume Next
        lc = ole.LinkedCell
        If Len(lc) > 0 Then ws.Parent.Worksheets(ws.name).Range(lc).ClearContents
        If TypeName(ole.Object) = "ComboBox" Then ole.Object.ListIndex = -1
        If TypeName(ole.Object) = "DropDown" Then ole.Object.ListIndex = -1
        On Error GoTo 0
    Next ole
End Sub

Private Sub ClearWorksheetCompletely(ws As Worksheet, Optional topLeftMsg As String = "")
    Dim co As ChartObject, shp As Shape
    With ws
        .Cells.Clear
        .Cells.FormatConditions.Delete
        .Cells.Interior.Pattern = xlNone
        .Cells.Borders.LineStyle = xlNone
        On Error Resume Next
        For Each co In .ChartObjects: co.Delete: Next co
        For Each shp In .Shapes: shp.Delete: Next shp
        On Error GoTo 0
        If Len(topLeftMsg) > 0 Then .Range("A1").value = topLeftMsg
    End With
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

Private Sub Clear_GraphsSheet(wb As Workbook)
    Dim ws As Worksheet
    Dim co As ChartObject, shp As Shape

    ' Case A: "Graphs" is a normal worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("Graphs")
    On Error GoTo 0
    If Not ws Is Nothing Then
        With ws
            ' remove charts + shapes
            On Error Resume Next
            For Each co In .ChartObjects: co.Delete: Next co
            For Each shp In .Shapes: shp.Delete: Next shp
            On Error GoTo 0

            ' optional: clear the sheet visuals/content so it's truly blank
            .Cells.Clear
            .Cells.FormatConditions.Delete
            .Cells.Interior.Pattern = xlNone
            .Cells.Borders.LineStyle = xlNone
            .Range("A1").value = "Graphs cleared."
        End With
        Exit Sub
    End If

    ' Case B: "Graphs" is a Chart sheet (not a Worksheet)
    Dim ch As Chart
    On Error Resume Next
    Set ch = wb.Charts("Graphs")
    On Error GoTo 0
    If Not ch Is Nothing Then
        ' deleting a Chart sheet is the only way to "clear" it
        Application.DisplayAlerts = False
        ch.Delete
        Application.DisplayAlerts = False
        ' recreate a blank worksheet named "Graphs" (so your layout stays consistent)
        wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)).name = "Graphs"
    End If
End Sub

Public Sub KOV_ClearSheet(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next

    ' remove any embedded charts/shapes
    Dim co As ChartObject, shp As Shape
    For Each co In ws.ChartObjects: co.Delete: Next co
    For Each shp In ws.Shapes: shp.Delete: Next shp

    ' clear content + formatting
    ws.Cells.Clear
    ws.Cells.FormatConditions.Delete
    ws.Cells.Interior.Pattern = xlNone
    ws.Cells.Borders.LineStyle = xlNone

    ' reset look & friendly message
    ws.Range("A1").value = "Select a product on UI and run KOV."
    ws.Columns("A:L").ColumnWidth = 14

    ' nudge UsedRange to refresh (safe no-op)
    Dim dummy As Range
    Set dummy = ws.UsedRange

    On Error GoTo 0
End Sub


