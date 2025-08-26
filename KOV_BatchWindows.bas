Attribute VB_Name = "KOV_BatchWindows"
Option Explicit

Public Sub Build_R4_BatchSummary_FromFTPT( _
    ByVal productName As String, _
    ByVal ftHeader As String, _
    ByVal ptHeader As String, _
    ByVal startFlow As Double, ByVal startPress As Double, ByVal holdStartMin As Double, _
    ByVal stripPress As Double, ByVal holdEachMin As Double, _
    Optional ByVal tagLabel As String = "R4")

    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Worksheets("Paste Data")
    Dim wsBS As Worksheet
    On Error Resume Next
    Set wsBS = ThisWorkbook.Worksheets("Batch Summary")
    On Error GoTo 0
    If wsBS Is Nothing Then
        Set wsBS = ThisWorkbook.Worksheets.Add(After:=wsD)
        wsBS.name = "Batch Summary"
        wsBS.Range("A1:G1").value = Array("Tag", "Batch Start", "Batch End", "Duration (min)", "Duration (hr)", "Status", "Product")
    End If

    Dim hdr As Object: Set hdr = BuildHeaderIndexAll(wsD) ' must be Public somewhere OR paste minimal versions here
    Dim cT As Long: cT = HeaderCol(hdr, "Time")
    If cT = 0 Then MsgBox "Paste Data missing 'Time'.", vbCritical: Exit Sub

    Dim cFT As Long: cFT = HeaderCol(hdr, ftHeader): If cFT = 0 Then cFT = HeaderCol(hdr, ftHeader & ".Val")
    Dim cPT As Long: cPT = HeaderCol(hdr, ptHeader): If cPT = 0 Then cPT = HeaderCol(hdr, ptHeader & ".Val")
    If cFT = 0 Or cPT = 0 Then
        MsgBox "Flow/Pressure headers not found: " & ftHeader & " / " & ptHeader, vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long: lastRow = wsD.Cells(wsD.Rows.Count, cT).End(xlUp).Row
    If lastRow < 3 Then MsgBox "Not enough rows in Paste Data.", vbExclamation: Exit Sub

    Dim outRow As Long: outRow = wsBS.Cells(wsBS.Rows.Count, 1).End(xlUp).Row + 1
    Dim i As Long, startIdx As Long, sStart As Long, bEnd As Long
    Dim acc As Double, dt As Double, ft As Double, pt As Double

    startIdx = 0: acc = 0
    For i = 3 To lastRow
        dt = (wsD.Cells(i, cT).value - wsD.Cells(i - 1, cT).value) * 24# * 60#
        ft = wsD.Cells(i, cFT).value
        pt = wsD.Cells(i, cPT).value

        ' START: FT>startFlow AND PT>startPress for holdStartMin
        If startIdx = 0 Then
            If ft > startFlow And pt > startPress Then
                acc = acc + Application.Max(0, dt)
                If acc >= holdStartMin Then
                    startIdx = i
                    ' Strip start: PT < stripPress for holdEachMin
                    sStart = FindHoldBelow_FromSheet(wsD, cT, cPT, i + 1, lastRow, stripPress, holdEachMin)
                    ' Batch end: PT > stripPress for holdEachMin
                    If sStart > 0 Then bEnd = FindHoldAbove_FromSheet(wsD, cT, cPT, sStart + 1, lastRow, stripPress, holdEachMin)
                    If bEnd = 0 Then Exit For

                    ' Write row
                    Dim tStart As Variant, tEnd As Variant
                    tStart = wsD.Cells(startIdx, cT).value
                    tEnd = wsD.Cells(bEnd, cT).value
                    wsBS.Cells(outRow, 1).value = tagLabel
                    wsBS.Cells(outRow, 2).value = tStart
                    wsBS.Cells(outRow, 3).value = tEnd
                    wsBS.Cells(outRow, 4).value = DateDiff("n", tStart, tEnd)
                    wsBS.Cells(outRow, 5).value = Round(wsBS.Cells(outRow, 4).value / 60#, 2)
                    wsBS.Cells(outRow, 6).value = "Complete"
                    outRow = outRow + 1

                    ' Continue after this batch
                    i = bEnd + 1
                    startIdx = 0: acc = 0
                End If
            Else
                acc = 0
            End If
        End If
    Next i

End Sub

Public Function FindHoldBelow_FromSheet(ws As Worksheet, cT As Long, cCol As Long, _
    ByVal fromRow As Long, ByVal lastRow As Long, ByVal thresh As Double, ByVal holdMin As Double) As Long
    Dim i As Long, acc As Double, dt As Double, startIdx As Long
    For i = Application.Max(fromRow, 3) To lastRow
        dt = (ws.Cells(i, cT).value - ws.Cells(i - 1, cT).value) * 24# * 60#
        If ws.Cells(i, cCol).value < thresh Then
            If startIdx = 0 Then startIdx = i
            acc = acc + Application.Max(0, dt)
            If acc >= holdMin Then FindHoldBelow_FromSheet = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0
        End If
    Next i
End Function

Public Function FindHoldAbove_FromSheet(ws As Worksheet, cT As Long, cCol As Long, _
    ByVal fromRow As Long, ByVal lastRow As Long, ByVal thresh As Double, ByVal holdMin As Double) As Long
    Dim i As Long, acc As Double, dt As Double, startIdx As Long
    For i = Application.Max(fromRow, 3) To lastRow
        dt = (ws.Cells(i, cT).value - ws.Cells(i - 1, cT).value) * 24# * 60#
        If ws.Cells(i, cCol).value > thresh Then
            If startIdx = 0 Then startIdx = i
            acc = acc + Application.Max(0, dt)
            If acc >= holdMin Then FindHoldAbove_FromSheet = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0
        End If
    Next i
End Function

Sub R4_Build_Week_BatchSummary()
    Dim wsBS As Worksheet
    On Error Resume Next
    Set wsBS = ThisWorkbook.Worksheets("Batch Summary")
    On Error GoTo 0
    If wsBS Is Nothing Then
        Set wsBS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Paste Data"))
        wsBS.name = "Batch Summary"
        wsBS.Range("A1:G1").value = Array("Tag", "Batch Start", "Batch End", _
                                           "Duration (min)", "Duration (hr)", "Status", "Product")
    Else
        wsBS.Cells.ClearContents
        wsBS.Range("A1:G1").value = Array("Tag", "Batch Start", "Batch End", _
                                           "Duration (min)", "Duration (hr)", "Status", "Product")
    End If

    ' Build from R4 FT/PT once; user will fill Product column manually
    Build_R4_BatchSummary_FromFTPT "", "R4_FT_01", "R4_PT_01", _
                                   500, 12, 10, 12, 10, "R4"
    ' ^ productName passed as "" and we commented the write, so Product stays blank
End Sub



