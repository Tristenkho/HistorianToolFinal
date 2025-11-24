Attribute VB_Name = "Graph"
'==========================
' Graphs Module (3-column grid) — X axis auto-picked by Excel
' X values are hours since first valid timestamp in column A
'==========================
Option Explicit

' Public entry: shows in Macros list
Public Sub PlotAllTags_Grid()
    PlotTagsGrid_Impl ThisWorkbook.Worksheets("Paste Data"), "Graphs", 3
End Sub

'--------------------------
' Private core
'--------------------------
Private Sub PlotTagsGrid_Impl(ByVal wsData As Worksheet, _
                              ByVal graphsSheetName As String, _
                              ByVal colsPerRow As Long)

    Dim wsCharts As Worksheet
    Dim lastCol As Long, lastRow As Long
    Dim timeRange As Range, tagRange As Range
    Dim xVals As Variant

    Dim chartW As Single, chartH As Single
    Dim marginL As Single, marginT As Single, hGap As Single, vGap As Single
    Dim gridRow As Long, gridCol As Long
    Dim chObj As ChartObject
    Dim c As Long, plotted As Long

    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    If wsData Is Nothing Then
        MsgBox "'Paste Data' sheet not found.", vbCritical
        GoTo CleanExit
    End If

    ' Layout settings (tweak to taste)
    chartW = 420
    chartH = 240
    marginL = 18
    marginT = 18
    hGap = 16
    vGap = 16
    If colsPerRow < 1 Then colsPerRow = 3

    ' Find data bounds
    lastRow = wsData.Cells(wsData.rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol < 2 Then
        MsgBox "No data to plot. Expecting time in column A and tags in columns B..", vbExclamation
        GoTo CleanExit
    End If

    Set timeRange = wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastRow, 1))

    ' Build relative-hour X array once for all charts (0 at first timestamp)
    xVals = MakeRelativeHoursArray(timeRange)

    ' Ensure/clear Graphs sheet
    Set wsCharts = EnsureSheet(graphsSheetName, wsData)
    ClearAllCharts wsCharts
    wsCharts.Cells.ClearContents

    ' Loop all tag columns B..last
    plotted = 0
    For c = 2 To lastCol
        ' Skip empty tag columns
        If Application.WorksheetFunction.CountA(wsData.Range(wsData.Cells(2, c), wsData.Cells(lastRow, c))) = 0 Then
            GoTo NextC
        End If

        Set tagRange = wsData.Range(wsData.Cells(2, c), wsData.Cells(lastRow, c))

        gridRow = plotted \ colsPerRow
        gridCol = plotted Mod colsPerRow

        Set chObj = wsCharts.ChartObjects.Add( _
            Left:=marginL + gridCol * (chartW + hGap), _
            Top:=marginT + gridRow * (chartH + vGap), _
            Width:=chartW, Height:=chartH)

        With chObj.Chart
            .ChartType = xlXYScatterLines
            .HasLegend = False

Dim xDec As Variant, yDec As Variant
MakeDecimatedXY timeRange, wsData.Range(wsData.Cells(2, c), wsData.Cells(lastRow, c)), 5000, xDec, yDec

.SeriesCollection.NewSeries
With .SeriesCollection(1)
    .name = wsData.Cells(1, c).value
    .XValues = xDec
    .Values = yDec
    .MarkerStyle = xlMarkerStyleNone
    On Error Resume Next
    .Format.Line.Weight = 0.75
    On Error GoTo 0
End With

            .HasTitle = True
            .ChartTitle.Text = wsData.Cells(1, c).value

            ' X axis: let Excel auto-pick nice whole-hour ticks
            With .Axes(xlCategory)
                .HasTitle = True
                .AxisTitle.Text = "Time (hr)"
                .MinimumScaleIsAuto = True
                .MaximumScaleIsAuto = True
                .MajorUnitIsAuto = True
                .MinorUnitIsAuto = True
                .HasMajorGridlines = False
                .HasMinorGridlines = False
                .TickLabels.NumberFormat = "0"      ' show whole hours
            End With

            ' Y axis
            With .Axes(xlValue)
                .HasTitle = True
                .AxisTitle.Text = "Value"
            End With
        End With

        plotted = plotted + 1
NextC:
    Next c

    wsCharts.Activate
    ActiveWindow.Zoom = 90

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Error while creating charts: " & Err.Description, vbCritical
End Sub

' Build a 1-D array of hours since the first valid timestamp in rng
Private Function MakeRelativeHoursArray(ByVal rng As Range) As Variant
    Dim arr As Variant, outArr() As Double
    Dim n As Long, i As Long
    Dim t0 As Double, t As Double
    Dim foundT0 As Boolean

    arr = rng.Value2          ' 2-D [1..n, 1..1] of Excel serials (days)
    n = UBound(arr, 1)
    ReDim outArr(1 To n)

    ' Find first valid timestamp as t0
    For i = 1 To n
        If IsNumeric(arr(i, 1)) Then
            t0 = CDbl(arr(i, 1))
            foundT0 = True
            Exit For
        End If
    Next i
    If Not foundT0 Then t0 = 0#

    ' Convert each to hours since t0
    For i = 1 To n
        If IsNumeric(arr(i, 1)) Then
            t = CDbl(arr(i, 1))
            outArr(i) = (t - t0) * 24#
        Else
            outArr(i) = IIf(i = 1, 0#, outArr(i - 1))
        End If
    Next i

    MakeRelativeHoursArray = outArr
End Function

' Ensure a sheet exists; if not, create it after the data sheet
Private Function EnsureSheet(ByVal name As String, ByVal afterSheet As Worksheet) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=afterSheet)
        ws.name = name
    End If
    Set EnsureSheet = ws
End Function

' Delete all charts on a sheet
Private Sub ClearAllCharts(ByVal ws As Worksheet)
    Do While ws.ChartObjects.Count > 0
        ws.ChartObjects(1).Delete
    Loop
End Sub

' Return sampled indices from 1..n so that at most maxPts are chosen (always includes first/last)
Private Function SampleIndices(ByVal n As Long, ByVal maxPts As Long) As Variant
    Dim stepN As Double, i As Long, k As Long
    If n <= 0 Then
        SampleIndices = Array()
        Exit Function
    End If
    If n <= maxPts Then
        ReDim outIdx(1 To n) As Long
        For i = 1 To n: outIdx(i) = i: Next i
        SampleIndices = outIdx
        Exit Function
    End If
    stepN = (n - 1) / (maxPts - 1)
    ReDim outIdx(1 To maxPts) As Long
    For k = 0 To maxPts - 1
        outIdx(k + 1) = 1 + CLng(Round(k * stepN, 0))
        If outIdx(k + 1) < 1 Then outIdx(k + 1) = 1
        If outIdx(k + 1) > n Then outIdx(k + 1) = n
    Next k
    SampleIndices = outIdx
End Function

' Build decimated arrays of (hours since t0, tag values), skipping rows with non-time or empty Y
Private Sub MakeDecimatedXY(ByVal timeRng As Range, ByVal valRng As Range, _
                            ByVal maxPts As Long, ByRef xOut As Variant, ByRef yOut As Variant)
    Dim tArr As Variant, vArr As Variant, tmpX() As Double, tmpY() As Double
    Dim n As Long, i As Long, t0 As Double, haveT0 As Boolean, keep() As Long, keepCount As Long
    tArr = timeRng.Value2   ' [1..n,1]
    vArr = valRng.Value2    ' [1..n,1]
    n = UBound(tArr, 1)

    ' find first valid time as t0
    For i = 1 To n
        If IsNumeric(tArr(i, 1)) And tArr(i, 1) > 0 Then
            t0 = CDbl(tArr(i, 1))
            haveT0 = True
            Exit For
        End If
    Next i
    If Not haveT0 Then
        ReDim xOut(1 To 1): ReDim yOut(1 To 1)
        xOut(1) = 0: yOut(1) = CVErr(xlErrNA)
        Exit Sub
    End If

    ' collect rows that have both a time and a numeric Y
    ReDim keep(1 To n)
    For i = 1 To n
        If IsNumeric(tArr(i, 1)) And IsNumeric(vArr(i, 1)) Then
            keepCount = keepCount + 1
            keep(keepCount) = i
        End If
    Next i
    If keepCount = 0 Then
        ReDim xOut(1 To 1): ReDim yOut(1 To 1)
        xOut(1) = 0: yOut(1) = CVErr(xlErrNA)
        Exit Sub
    End If

    ' sample indices down to ~maxPts
    Dim idxs As Variant, m As Long, j As Long
    idxs = SampleIndices(keepCount, maxPts)
    m = UBound(idxs)

    ReDim tmpX(1 To m): ReDim tmpY(1 To m)
    For j = 1 To m
        i = keep(idxs(j))
        tmpX(j) = (CDbl(tArr(i, 1)) - t0) * 24#  ' hours since t0
        tmpY(j) = CDbl(vArr(i, 1))
    Next j

    xOut = tmpX
    yOut = tmpY
End Sub
