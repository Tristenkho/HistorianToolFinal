Attribute VB_Name = "GraphByBatch"
Option Explicit

' ===========================
' Module-level configuration
' ===========================
Private Const PASTE_SHEET As String = "Paste Data"
Private Const SUMMARY_SHEET As String = "Batch Summary"
Private Const GRAPHS_SHEET As String = "Graphs"
Private Const OVERLAYS_SHEET As String = "Overlays"
Private Const SCRATCH_SHEET As String = "Scratch"

' Selection behavior
Private Const ALLOW_ROW_CLICK_SELECTION As Boolean = True   ' click any cells on batch rows; map to B:C

' Visual toggles
Private Const USE_COLOR_CYCLE As Boolean = True
Private Const FORCE_COMMON_Y As Boolean = False             ' only for overlay charts
Private Const AUTO_ZOOM As Long = 90

' Line appearance controls (default = muted)
Private Const LINE_WEIGHT_NORMAL As Single = 0.75           ' thin line for default (non-highlighted) batches
Private Const LINE_WEIGHT_HIGHLIGHT As Single = 1.5         ' thicker line for highlighted batches
Private Const LINE_TRANSPARENCY_NORMAL As Double = 0.5      ' semi-transparent baseline
Private Const LINE_TRANSPARENCY_HIGHLIGHT As Double = 0     ' fully opaque highlight

' Chart sizing
Private Const OVERLAY_CHART_W As Single = 420
Private Const OVERLAY_CHART_H As Single = 260
Private Const ROW_CHART_W As Single = 360
Private Const ROW_CHART_H As Single = 220
Private Const COLS_PER_ROW As Long = 3
Private Const GAP_H As Single = 16
Private Const GAP_V As Single = 18
Private Const MARGIN_L As Single = 12
Private Const MARGIN_T As Single = 12

' ===========================
' Public entry points (3)
' ===========================

' 1) Row-of-charts: one row per batch, one chart per selected tag (X = hours since start)
Public Sub Plot_Batch_As_Row_Of_Tag_Charts()
    Dim wsData As Worksheet, wsSum As Worksheet, wsCharts As Worksheet, wsScratch As Worksheet
    Dim tagHeaders As Range, batchSel As Range
    Dim tagCols() As Long, nTags As Long
    Dim firstTime As Double, lastTime As Double
    Dim lastRow As Long, lastCol As Long
    Dim prevCalc As XlCalculation, prevStatus As Variant, prevDisplayStatusBar As Boolean

    On Error GoTo CleanFail
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    prevCalc = Application.Calculation
    prevStatus = Application.StatusBar
    prevDisplayStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationManual

    ' Sheets
    Set wsData = SheetOrFail(PASTE_SHEET): If wsData Is Nothing Then GoTo CleanExit
    Set wsSum = SheetOrFail(SUMMARY_SHEET): If wsSum Is Nothing Then GoTo CleanExit
    Set wsCharts = EnsureSheet(GRAPHS_SHEET, wsData)
    Set wsScratch = EnsureSheet(SCRATCH_SHEET, wsData)
    wsScratch.Cells.Clear
    wsScratch.Visible = xlSheetHidden

    ' Data checks
    lastRow = wsData.Cells(wsData.rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    If lastRow < 3 Or lastCol < 2 Then
        MsgBox "Paste Data: need time in col A and tag headers in row 1 (B:...).", vbExclamation
        GoTo CleanExit
    End If
    If Not GetFirstLastTime(wsData, 1, firstTime, lastTime) Then
        MsgBox "No valid timestamps in Paste Data col A.", vbExclamation
        GoTo CleanExit
    End If

    ' ===== Prompt 1: TAG HEADERS =====
    Application.ScreenUpdating = True
    wsData.Activate
    Set tagHeaders = Application.InputBox( _
        "Select TAG HEADER cells (row 1 on 'Paste Data'). Ctrl+click to multi-select. Do not include A1.", _
        "Pick Tags", Type:=8)
    Application.ScreenUpdating = False
    If tagHeaders Is Nothing Then GoTo CleanExit
    Set tagHeaders = Intersect(tagHeaders, wsData.rows(1))
    If tagHeaders Is Nothing Then
        MsgBox "Pick headers on row 1 of 'Paste Data', starting at B1.", vbExclamation
        GoTo CleanExit
    End If
    nTags = MapHeaderCols(wsData, tagHeaders, tagCols)
    If nTags = 0 Then
        MsgBox "No valid tag headers selected.", vbExclamation
        GoTo CleanExit
    End If

    ' ===== Prompt 2: INCLUDE BATCH ROWS =====
    Dim rows() As Long, nRows As Long
    If ALLOW_ROW_CLICK_SELECTION Then
        Application.ScreenUpdating = True
        wsSum.Activate: Application.GoTo wsSum.Range("A1"), True
        Set batchSel = Application.InputBox( _
            "Click ANY cell on each batch row to INCLUDE. Ctrl+click to multi-select.", _
            "Pick Batches (Included)", Type:=8)
        Application.ScreenUpdating = False
        If batchSel Is Nothing Then GoTo CleanExit
        If Not batchSel.Parent Is wsSum Then
            MsgBox "Please select rows on 'Batch Summary'.", vbExclamation
            GoTo CleanExit
        End If
        CollectRowsFromSelection batchSel, rows, nRows
    Else
        Application.ScreenUpdating = True
        wsSum.Activate: Application.GoTo wsSum.Range("B2"), True
        Set batchSel = Application.InputBox( _
            "Select the B:C cells for all batches to INCLUDE (e.g., B3:C20). You can Ctrl+click multiple.", _
            "Pick Batches (Included)", Type:=8)
        Application.ScreenUpdating = False
        If batchSel Is Nothing Then GoTo CleanExit
        If Not batchSel.Parent Is wsSum Then
            MsgBox "Please select columns B:C on 'Batch Summary'.", vbExclamation
            GoTo CleanExit
        End If
        CollectRowsFromSelection batchSel, rows, nRows
    End If
    If nRows = 0 Then
        MsgBox "No batch rows selected.", vbExclamation
        GoTo CleanExit
    End If

    ' Clear output
    ClearAllCharts wsCharts
    wsCharts.Cells.Clear

    ' Build charts: one row per batch, one chart per selected tag
    Dim b As Long, t As Long, built As Long
    Dim usedStart As Double, usedEnd As Double
    Dim rngX As Range, rngY As Range
    Dim chObj As ChartObject
    Dim leftPos As Single, topPos As Single

    For b = 1 To nRows
        SetStatus "Row-of-charts: batch " & b & " / " & nRows & " …", True

        ' Clamp batch window to data range
        usedStart = GetDateOrDefault(wsSum.Cells(rows(b), 2), firstTime)
        usedEnd = GetDateOrDefault(wsSum.Cells(rows(b), 3), lastTime)
        usedStart = WorksheetFunction.Max(usedStart, firstTime)
        usedEnd = WorksheetFunction.Min(usedEnd, lastTime)
        If usedEnd <= usedStart Then GoTo NextBatchR

        ' Build X for this batch
        Set rngX = BuildScratchTime(wsScratch, wsData, usedStart, usedEnd, lastRow, 1, b)

        ' Create a row of charts (one per tag)
        For t = 1 To nTags
            Set rngY = BuildScratchSeries(wsScratch, wsData, usedStart, usedEnd, lastRow, tagCols(t), b, t)

            leftPos = MARGIN_L + ((t - 1) Mod nTags) * (ROW_CHART_W + GAP_H)
            topPos = MARGIN_T + (b - 1) * (ROW_CHART_H + GAP_V)

            Set chObj = wsCharts.ChartObjects.Add(Left:=leftPos, Top:=topPos, Width:=ROW_CHART_W, Height:=ROW_CHART_H)
            With chObj.Chart
                .ChartType = xlXYScatterLines
                .HasLegend = False
                .HasTitle = True
                .ChartTitle.Text = wsData.Cells(1, tagCols(t)).value & "  |  " & _
                    Format$(CDate(usedStart), "m/d h:mm") & " - " & Format$(CDate(usedEnd), "m/d h:mm") & _
                    "  |  Duration: " & Format$((usedEnd - usedStart) * 24#, "0") & " hr"

                With .SeriesCollection.NewSeries
                    .name = wsData.Cells(1, tagCols(t)).value
                    .XValues = rngX
                    .Values = rngY
                    .MarkerStyle = xlMarkerStyleNone

                    ' Always solid, thin lines (no highlight logic here)
                    .Format.Line.Weight = 1
                    .Format.Line.Transparency = 0
                End With

                With .Axes(xlCategory)
                    .HasTitle = True
                    .AxisTitle.Text = "Hours since batch start"
                    .TickLabels.NumberFormat = "0"
                End With
                With .Axes(xlValue)
                    .HasTitle = True
                    .AxisTitle.Text = "Value"
                End With
            End With
            built = built + 1
        Next t
NextBatchR:
    Next b

    wsScratch.Visible = xlSheetHidden
    wsCharts.Activate
    Application.ScreenUpdating = True
    ActiveWindow.Zoom = AUTO_ZOOM
    If built = 0 Then MsgBox "No charts created. Check selections and data windows.", vbExclamation

CleanExit:
    Application.Calculation = prevCalc
    Application.DisplayStatusBar = prevDisplayStatusBar
    ClearStatus prevStatus
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.Calculation = prevCalc
    Application.DisplayStatusBar = prevDisplayStatusBar
    ClearStatus prevStatus
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' 2) Overlay (Normal X = hours since start); one chart per tag, series = batches
Public Sub Plot_Tags_Overlay_Normal()
    Plot_Tags_Overlay_Core False
End Sub

' 3) Overlay (Scaled X = 0..1 progress); one chart per tag, series = batches
Public Sub Plot_Tags_Overlay_ScaledX()
    Plot_Tags_Overlay_Core True
End Sub

' ===========================
' Overlay core (with highlight subset)
' ===========================
Private Sub Plot_Tags_Overlay_Core(ByVal scaledX As Boolean)
    Dim wsData As Worksheet, wsSum As Worksheet, wsCharts As Worksheet, wsScratch As Worksheet
    Dim tagHeaders As Range, batchSel As Range, hiSel As Range
    Dim tagCols() As Long, nTags As Long
    Dim firstTime As Double, lastTime As Double
    Dim lastRow As Long, lastCol As Long
    Dim prevCalc As XlCalculation, prevStatus As Variant, prevDisplayStatusBar As Boolean

    On Error GoTo CleanFail
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    prevCalc = Application.Calculation
    prevStatus = Application.StatusBar
    prevDisplayStatusBar = Application.DisplayStatusBar
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationManual

    ' Sheets
    Set wsData = SheetOrFail(PASTE_SHEET): If wsData Is Nothing Then GoTo CleanExit
    Set wsSum = SheetOrFail(SUMMARY_SHEET): If wsSum Is Nothing Then GoTo CleanExit
    Set wsCharts = EnsureSheet(OVERLAYS_SHEET, wsData)
    Set wsScratch = EnsureSheet(SCRATCH_SHEET, wsData)
    wsScratch.Cells.Clear
    wsScratch.Visible = xlSheetHidden

    ' Data checks
    lastRow = wsData.Cells(wsData.rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    If lastRow < 3 Or lastCol < 2 Then
        MsgBox "Paste Data: need time in col A and tag headers in row 1 (B:...).", vbExclamation
        GoTo CleanExit
    End If
    If Not GetFirstLastTime(wsData, 1, firstTime, lastTime) Then
        MsgBox "No valid timestamps in Paste Data col A.", vbExclamation
        GoTo CleanExit
    End If

    ' ===== Prompt 1: TAG HEADERS =====
    Application.ScreenUpdating = True
    wsData.Activate
    Set tagHeaders = Application.InputBox( _
        "Select TAG HEADER cells (row 1 on 'Paste Data'). Ctrl+click to multi-select. Do not include A1.", _
        "Pick Tags", Type:=8)
    Application.ScreenUpdating = False
    If tagHeaders Is Nothing Then GoTo CleanExit
    Set tagHeaders = Intersect(tagHeaders, wsData.rows(1))
    If tagHeaders Is Nothing Then
        MsgBox "Pick headers on row 1 of 'Paste Data', starting at B1.", vbExclamation
        GoTo CleanExit
    End If
    nTags = MapHeaderCols(wsData, tagHeaders, tagCols)
    If nTags = 0 Then
        MsgBox "No valid tag headers selected.", vbExclamation
        GoTo CleanExit
    End If

    ' ===== Prompt 2: INCLUDE BATCH ROWS (B) =====
    Dim rows() As Long, nRows As Long
    If ALLOW_ROW_CLICK_SELECTION Then
        Application.ScreenUpdating = True
        wsSum.Activate: Application.GoTo wsSum.Range("A1"), True
        Set batchSel = Application.InputBox( _
            "Click ANY cell on each batch row to INCLUDE. Ctrl+click to multi-select.", _
            "Pick Batches (Included)", Type:=8)
        Application.ScreenUpdating = False
        If batchSel Is Nothing Then GoTo CleanExit
        If Not batchSel.Parent Is wsSum Then
            MsgBox "Please select rows on 'Batch Summary'.", vbExclamation
            GoTo CleanExit
        End If
        CollectRowsFromSelection batchSel, rows, nRows
    Else
        Application.ScreenUpdating = True
        wsSum.Activate: Application.GoTo wsSum.Range("B2"), True
        Set batchSel = Application.InputBox( _
            "Select the B:C cells for all batches to INCLUDE (e.g., B3:C20). You can Ctrl+click multiple.", _
            "Pick Batches (Included)", Type:=8)
        Application.ScreenUpdating = False
        If batchSel Is Nothing Then GoTo CleanExit
        If Not batchSel.Parent Is wsSum Then
            MsgBox "Please select columns B:C on 'Batch Summary'.", vbExclamation
            GoTo CleanExit
        End If
        CollectRowsFromSelection batchSel, rows, nRows
    End If
    If nRows = 0 Then
        MsgBox "No batch rows selected.", vbExclamation
        GoTo CleanExit
    End If

    Dim includedSet As Object: Set includedSet = BuildSetFromArray(rows)

    ' ===== Prompt 3: HIGHLIGHT SUBSET (H) — OPTIONAL =====
    Dim hiSet As Object: Set hiSet = CreateObject("Scripting.Dictionary")
    Application.ScreenUpdating = True
    wsSum.Activate: Application.GoTo wsSum.Range("A1"), True
    On Error Resume Next
    Set hiSel = Application.InputBox( _
        "Optional: click ANY cell on rows to HIGHLIGHT (Ctrl+click). Press Cancel for none.", _
        "Pick Highlighted (Optional)", Type:=8)
    On Error GoTo 0
    Application.ScreenUpdating = False
    If Not hiSel Is Nothing Then
        Set hiSet = BuildRowSetFromSelection(hiSel, rows) ' intersection guard
    End If

    ' Clear output
    ClearAllCharts wsCharts
    wsCharts.Cells.Clear

    ' Build charts: one chart per tag, multiple series (batches)
    Dim t As Long, b As Long
    Dim usedStart As Double, usedEnd As Double
    Dim rngX As Range, rngY As Range
    Dim chObj As ChartObject
    Dim leftPos As Single, topPos As Single
    Dim titleSuffix As String

    titleSuffix = IIf(scaledX, " — Overlay (X scaled 0–1)", " — Overlay (Hours)")

    For t = 1 To nTags
        SetStatus "Overlay: tag " & t & " / " & nTags & " …", True

        leftPos = MARGIN_L + ((t - 1) Mod COLS_PER_ROW) * (OVERLAY_CHART_W + GAP_H)
        topPos = MARGIN_T + ((t - 1) \ COLS_PER_ROW) * (OVERLAY_CHART_H + GAP_V)

        Set chObj = wsCharts.ChartObjects.Add(Left:=leftPos, Top:=topPos, Width:=OVERLAY_CHART_W, Height:=OVERLAY_CHART_H)
        With chObj.Chart
            .ChartType = xlXYScatterLines
            .HasLegend = True
            .HasTitle = True
            .ChartTitle.Text = wsData.Cells(1, tagCols(t)).value & titleSuffix
            .DisplayBlanksAs = xlNotPlotted

            With .Axes(xlCategory)
                .HasTitle = True
                If scaledX Then
                    .AxisTitle.Text = "Normalized batch progress (0–1)"
                    .MinimumScale = 0
                    .MaximumScale = 1
                    .TickLabels.NumberFormat = "0.00"
                Else
                    .AxisTitle.Text = "Hours since batch start"
                    .TickLabels.NumberFormat = "0.0"
                End If
            End With
            With .Axes(xlValue)
                .HasTitle = True
                .AxisTitle.Text = "Value"
            End With
        End With

        ' Add each batch as a series
        For b = 1 To nRows
            SetStatus "Overlay: tag " & t & " / " & nTags & "  —  batch " & b & " / " & nRows & " …"
            DoEvents

            usedStart = GetDateOrDefault(wsSum.Cells(rows(b), 2), firstTime)
            usedEnd = GetDateOrDefault(wsSum.Cells(rows(b), 3), lastTime)
            usedStart = WorksheetFunction.Max(usedStart, firstTime)
            usedEnd = WorksheetFunction.Min(usedEnd, lastTime)
            If usedEnd <= usedStart Then GoTo NextBatchS

            If scaledX Then
                Set rngX = BuildScratchTimeScaled01(wsScratch, wsData, usedStart, usedEnd, lastRow, 1, b)
            Else
                Set rngX = BuildScratchTime(wsScratch, wsData, usedStart, usedEnd, lastRow, 1, b)
            End If
            Set rngY = BuildScratchSeries(wsScratch, wsData, usedStart, usedEnd, lastRow, tagCols(t), b, t)

            With chObj.Chart.SeriesCollection.NewSeries
                .name = Format$(CDate(usedStart), "m/d h:mm")
                .XValues = rngX
                .Values = rngY
                .MarkerStyle = xlMarkerStyleNone

                ' Color per batch (optional)
                If USE_COLOR_CYCLE Then .Format.Line.ForeColor.RGB = PickColor(b)

                ' Style based on highlight set (default = muted)
                ApplySeriesStyle .Format.Line, hiSet.Exists(rows(b))
            End With

NextBatchS:
        Next b

        If FORCE_COMMON_Y Then ApplyCommonYScale chObj.Chart
    Next t

    wsScratch.Visible = xlSheetHidden
    wsCharts.Activate
    Application.ScreenUpdating = True
    ActiveWindow.Zoom = AUTO_ZOOM

CleanExit:
    Application.Calculation = prevCalc
    Application.DisplayStatusBar = prevDisplayStatusBar
    ClearStatus prevStatus
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.Calculation = prevCalc
    Application.DisplayStatusBar = prevDisplayStatusBar
    ClearStatus prevStatus
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ===========================
' Helpers: selection, tags, sets, status
' ===========================

Private Function SheetOrFail(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetOrFail = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If SheetOrFail Is Nothing Then MsgBox "Sheet '" & nm & "' not found.", vbCritical
End Function

Private Function EnsureSheet(ByVal nm As String, ByVal afterSheet As Worksheet) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=afterSheet)
        ws.name = nm
    End If
    Set EnsureSheet = ws
End Function

Private Sub ClearAllCharts(ByVal ws As Worksheet)
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop
End Sub

' Build tag column list from selected headers (multi-area safe; skips col A; de-dupes; preserves left-to-right order)
Private Function MapHeaderCols(ws As Worksheet, hdr As Range, ByRef tagCols() As Long) As Long
    Dim c As Range, tmp() As Long, n As Long, i As Long, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    If hdr Is Nothing Then MapHeaderCols = 0: Exit Function
    ReDim tmp(1 To hdr.Cells.Count)
    For Each c In hdr.Cells
        If c.Row = 1 Then
            If c.Column > 1 Then
                If LenB(ws.Cells(1, c.Column).value) > 0 Then
                    If Not seen.Exists(c.Column) Then
                        n = n + 1
                        tmp(n) = c.Column
                        seen(c.Column) = True
                    End If
                End If
            End If
        End If
    Next c
    If n = 0 Then
        MapHeaderCols = 0
    Else
        ReDim tagCols(1 To n)
        For i = 1 To n: tagCols(i) = tmp(i): Next i
        MapHeaderCols = n
    End If
End Function

' Collect unique row numbers (>1) from selection; preserves click/area order; de-dupes
Private Sub CollectRowsFromSelection(sel As Range, ByRef rowsOut() As Long, ByRef nRows As Long)
    Dim area As Range, rr As Range, dict As Object, k As Variant
    Set dict = CreateObject("Scripting.Dictionary")
    For Each area In sel.Areas
        For Each rr In area.rows
            If rr.Row > 1 Then
                If Not dict.Exists(rr.Row) Then dict.Add rr.Row, rr.Row
            End If
        Next rr
    Next area
    If dict.Count = 0 Then Exit Sub
    ReDim rowsOut(1 To dict.Count)
    nRows = 0
    For Each k In dict.Keys
        nRows = nRows + 1
        rowsOut(nRows) = CLng(k)
    Next k
End Sub

' Build a set (Dictionary) from an array of Longs
Private Function BuildSetFromArray(a() As Long) As Object
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    For i = LBound(a) To UBound(a)
        d(a(i)) = True
    Next i
    Set BuildSetFromArray = d
End Function

' Build a HashSet of row numbers from a Range selection, intersected with the included rows()
Private Function BuildRowSetFromSelection(sel As Range, includedRows() As Long) As Object
    Dim d As Object, ok As Object, area As Range, rr As Range
    Dim i As Long
    Set d = CreateObject("Scripting.Dictionary")
    Set ok = CreateObject("Scripting.Dictionary")
    For i = LBound(includedRows) To UBound(includedRows)
        ok(includedRows(i)) = True
    Next i
    If sel Is Nothing Then
        Set BuildRowSetFromSelection = d
        Exit Function
    End If
    For Each area In sel.Areas
        For Each rr In area.rows
            If rr.Row > 1 Then
                If ok.Exists(rr.Row) Then d(rr.Row) = True ' intersection guard
            End If
        Next rr
    Next area
    Set BuildRowSetFromSelection = d
End Function

' ===========================
' Status Bar helpers
' ===========================
Private Sub SetStatus(ByVal msg As String, Optional ByVal force As Boolean = False)
    Static lastUpdate As Single
    If force Or Timer - lastUpdate > 0.2 Then
        Application.StatusBar = msg
        DoEvents
        lastUpdate = Timer
    End If
End Sub

Private Sub ClearStatus(Optional ByVal prevStatus As Variant)
    If IsMissing(prevStatus) Then
        Application.StatusBar = False
    ElseIf VarType(prevStatus) = vbBoolean And prevStatus = False Then
        Application.StatusBar = False
    Else
        Application.StatusBar = prevStatus
    End If
End Sub

' ===========================
' Scratch builders (R1C1 formulas)
' ===========================
' Scan for first and last valid times in the given column
Private Function GetFirstLastTime(ws As Worksheet, ByVal timeCol As Long, ByRef firstT As Double, ByRef lastT As Double) As Boolean
    Dim r As Long, lr As Long
    lr = ws.Cells(ws.rows.Count, timeCol).End(xlUp).Row
    firstT = 0#: lastT = 0#
    If lr < 2 Then Exit Function

    For r = 2 To lr
        If IsDate(ws.Cells(r, timeCol).value) Then
            firstT = CDbl(CDate(ws.Cells(r, timeCol).value))
            Exit For
        End If
    Next r
    If firstT = 0# Then Exit Function

    For r = lr To 2 Step -1
        If IsDate(ws.Cells(r, timeCol).value) Then
            lastT = CDbl(CDate(ws.Cells(r, timeCol).value))
            Exit For
        End If
    Next r

    GetFirstLastTime = (lastT >= firstT)
End Function

' Return the cell's date as Excel serial if valid; otherwise defaultVal
Private Function GetDateOrDefault(c As Range, ByVal defaultVal As Double) As Double
    Dim v As Variant: v = c.value
    If IsDate(v) Then
        GetDateOrDefault = CDbl(CDate(v))
    Else
        GetDateOrDefault = defaultVal
    End If
End Function

' Build X (hours since start) with NA() outside [startDT, endDT]
Private Function BuildScratchTime(wsScratch As Worksheet, wsData As Worksheet, _
                                  ByVal startDT As Double, ByVal endDT As Double, _
                                  ByVal lastRow As Long, ByVal timeCol As Long, _
                                  ByVal batchIdx As Long) As Range
    Dim c As Long
    c = (batchIdx - 1) * 100 + 1
    wsScratch.Cells(1, c).value = "Hours_since_start_b" & batchIdx
    With wsScratch.Range(wsScratch.Cells(2, c), wsScratch.Cells(lastRow, c))
        .FormulaR1C1 = "=IF(AND('" & wsData.name & "'!R[0]C" & timeCol & ">=" & startDT & "," & _
                               "'" & wsData.name & "'!R[0]C" & timeCol & "<=" & endDT & ")," & _
                               "('" & wsData.name & "'!R[0]C" & timeCol & "-" & startDT & ")*24,NA())"
        .NumberFormat = "0.0"
    End With
    Set BuildScratchTime = wsScratch.Range(wsScratch.Cells(2, c), wsScratch.Cells(lastRow, c))
End Function

' Build Y for one tag with NA() outside [startDT, endDT]
Private Function BuildScratchSeries(wsScratch As Worksheet, wsData As Worksheet, _
                                    ByVal startDT As Double, ByVal endDT As Double, _
                                    ByVal lastRow As Long, ByVal tagCol As Long, _
                                    ByVal batchIdx As Long, ByVal seriesIdx As Long) As Range
    Dim c As Long
    c = (batchIdx - 1) * 100 + 1 + seriesIdx
    wsScratch.Cells(1, c).value = "Y_b" & batchIdx & "_s" & seriesIdx
    With wsScratch.Range(wsScratch.Cells(2, c), wsScratch.Cells(lastRow, c))
        .FormulaR1C1 = "=IF(AND('" & wsData.name & "'!R[0]C1>=" & startDT & "," & _
                               "'" & wsData.name & "'!R[0]C1<=" & endDT & ")," & _
                               "'" & wsData.name & "'!R[0]C" & tagCol & ",NA())"
        .NumberFormat = "0.00"
    End With
    Set BuildScratchSeries = wsScratch.Range(wsScratch.Cells(2, c), wsScratch.Cells(lastRow, c))
End Function

' Build normalized X (0..1) with NA() outside [startDT, endDT]
Private Function BuildScratchTimeScaled01(wsScratch As Worksheet, wsData As Worksheet, _
                                          ByVal startDT As Double, ByVal endDT As Double, _
                                          ByVal lastRow As Long, ByVal timeCol As Long, _
                                          ByVal batchIdx As Long) As Range
    Dim c As Long
    c = (batchIdx - 1) * 300 + 1
    wsScratch.Cells(1, c).value = "Progress_0to1_b" & batchIdx
    With wsScratch.Range(wsScratch.Cells(2, c), wsScratch.Cells(lastRow, c))
        .FormulaR1C1 = "=IF(AND('" & wsData.name & "'!R[0]C" & timeCol & ">=" & startDT & "," & _
                               "'" & wsData.name & "'!R[0]C" & timeCol & "<=" & endDT & ")," & _
                               "IF((" & endDT & "-" & startDT & ")=0,NA(),(" & _
                               "'" & wsData.name & "'!R[0]C" & timeCol & "-" & startDT & ")/(" & endDT & "-" & startDT & "))," & _
                               "NA())"
        .NumberFormat = "0.000"
    End With
    Set BuildScratchTimeScaled01 = wsScratch.Range(wsScratch.Cells(2, c), wsScratch.Cells(lastRow, c))
End Function

' Optional: apply common Y scale by scanning series values
Private Sub ApplyCommonYScale(ByVal ch As Chart)
    Dim s As series, ymin As Double, ymax As Double
    Dim v As Variant, i As Long
    ymin = 1E+99: ymax = -1E+99

    For Each s In ch.SeriesCollection
        v = s.Values
        If IsArray(v) Then
            For i = LBound(v) To UBound(v)
                If Not IsError(v(i)) Then
                    If IsNumeric(v(i)) Then
                        If v(i) < ymin Then ymin = v(i)
                        If v(i) > ymax Then ymax = v(i)
                    End If
                End If
            Next i
        End If
    Next s

    If ymin < 1E+98 And ymax > -1E+98 And ymax > ymin Then
        With ch.Axes(xlValue)
            .MinimumScaleIsAuto = False
            .MaximumScaleIsAuto = False
            .MinimumScale = ymin
            .MaximumScale = ymax
        End With
    End If
End Sub

' High-contrast color cycle
Private Function PickColor(ByVal idx As Long) As Long
    Dim colors As Variant
    colors = Array( _
        RGB(33, 150, 243), RGB(244, 67, 54), RGB(76, 175, 80), _
        RGB(255, 193, 7), RGB(156, 39, 176), RGB(255, 87, 34), _
        RGB(0, 188, 212), RGB(121, 85, 72), RGB(63, 81, 181))
    PickColor = colors((idx - 1) Mod (UBound(colors) + 1))
End Function

' Apply style based on highlight flag (default = muted)
Private Sub ApplySeriesStyle(ByVal ln As LineFormat, ByVal isHighlighted As Boolean)
    On Error Resume Next
    If isHighlighted Then
        ln.Weight = LINE_WEIGHT_HIGHLIGHT
        ln.Transparency = LINE_TRANSPARENCY_HIGHLIGHT
    Else
        ln.Weight = LINE_WEIGHT_NORMAL
        ln.Transparency = LINE_TRANSPARENCY_NORMAL
    End If
    On Error GoTo 0
End Sub


