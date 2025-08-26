Attribute VB_Name = "GraphAutomation"
'==========================
' Graphs Module (3-column grid) — X axis in hours (0 start)
' Major ticks every 12 hr; minor ticks every 6 hr; no minor gridlines
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
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol < 2 Then
        MsgBox "No data to plot. Expecting time in column A and tags in columns B..", vbExclamation
        GoTo CleanExit
    End If

    Set timeRange = wsData.Range(wsData.Cells(2, 1), wsData.Cells(lastRow, 1))

    ' Build relative-hour X array once for all charts
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

            .SeriesCollection.NewSeries
            With .SeriesCollection(1)
                .name = wsData.Cells(1, c).value
                .XValues = xVals                ' hours since t0 (0 at start)
                .Values = tagRange
                .MarkerStyle = xlMarkerStyleNone
                On Error Resume Next
                .Format.Line.Weight = 0.75
                On Error GoTo 0
            End With

            .HasTitle = True
            .ChartTitle.Text = wsData.Cells(1, c).value

            ' X axis: fixed spacing, clean grid
            With .Axes(xlCategory)
                .HasTitle = True
                .AxisTitle.Text = "Time (hr)"
                .MinimumScale = 0
                .MaximumScaleIsAuto = True
                .MajorUnit = 12                 ' major tick every 12 hr
                .MinorUnit = 6                  ' minor tick every 6 hr
                .HasMajorGridlines = False
                .HasMinorGridlines = False      ' hide minor gridlines (ticks still show)
                .TickLabels.NumberFormat = "0"  ' whole hours
            End With

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

