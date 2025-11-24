Attribute VB_Name = "KOV_BatchWindows"
'=== Module: modBatchWindows_R4 ===
Option Explicit

' -------- Entry points for buttons / Macros dialog --------
Public Sub R4_Build()
    ' Append R4 rows; user fills Product later
    Build_R4_BatchSummary_FromFTPT "R4_FT_01", "R4_PT_01", _
                                   500, 12, 10, _
                                   12, 10, _
                                   "R4"
End Sub

Public Sub R4_Run()
    ' Use your existing WeekRunner
    KOV_Run_FromBatchSummary
End Sub

' -------- Core builder (R4 FT/PT only; Product left blank) --------
Public Sub Build_R4_BatchSummary_FromFTPT( _
    ByVal ftHeader As String, _
    ByVal ptHeader As String, _
    ByVal startFlow As Double, ByVal startPress As Double, ByVal holdStartMin As Double, _
    ByVal stripPress As Double, ByVal holdEachMin As Double, _
    Optional ByVal tagLabel As String = "R4")

    Dim wsD As Worksheet, wsBS As Worksheet
    Dim hdr As Object
    Dim cT As Long, cFT As Long, cPT As Long
    Dim lastRow As Long
    Dim i As Long, startIdx As Long, sStart As Long, bEnd As Long
    Dim acc As Double, dt As Double, ft As Double, pt As Double
    Dim tStart As Variant, tEnd As Variant

    Dim holdAcc As Double, startCandIdx As Long
    Dim started As Boolean, startedBeforeData As Boolean
    Dim prevFT As Double, prevPT As Double
    Dim batchStartTime As Variant, batchEndTime As Variant

    Set wsD = ThisWorkbook.Worksheets("Paste Data")
    Set wsBS = EnsureBatchSummary()

    ' Ensure headers exist
    If wsBS.Cells(1, 1).value <> "Tag" Then
        wsBS.Range("A1:G1").value = Array("Tag", "Batch Start", "Batch End", _
                                          "Duration (min)", "Duration (hr)", "Status", "Product")
    End If
    wsBS.Columns(2).NumberFormat = "m/dd/yyyy hh:mm"
    wsBS.Columns(3).NumberFormat = "m/dd/yyyy hh:mm"

    Set hdr = BuildHeaderIndexAll(wsD)
    cT = HeaderCol(hdr, "Time")
    If cT = 0 Then MsgBox "Paste Data is missing header 'Time'.", vbCritical: Exit Sub

    cFT = HeaderCol(hdr, ftHeader): If cFT = 0 Then cFT = HeaderCol(hdr, ftHeader & ".Val")
    cPT = HeaderCol(hdr, ptHeader): If cPT = 0 Then cPT = HeaderCol(hdr, ptHeader & ".Val")
    If cFT = 0 Or cPT = 0 Then
        MsgBox "Flow/Pressure headers not found: " & ftHeader & " / " & ptHeader, vbExclamation
        Exit Sub
    End If

    lastRow = wsD.Cells(wsD.rows.Count, cT).End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "Not enough rows in 'Paste Data'.", vbExclamation
        Exit Sub
    End If
    
    ' Pre-loop initialization like Standard module
    prevFT = wsD.Cells(2, cFT).value
    prevPT = wsD.Cells(2, cPT).value
    started = False
    startedBeforeData = False
    startCandIdx = 0
    holdAcc = 0

    ' If we already begin above both thresholds, remember it
    If prevFT > startFlow And prevPT > startPress Then
        startCandIdx = 2
        holdAcc = 0
        startedBeforeData = True
    End If

    startIdx = 0: acc = 0
    For i = 3 To lastRow
        dt = (wsD.Cells(i, cT).value - wsD.Cells(i - 1, cT).value) * 24# * 60#
        If dt < 0 Then dt = 0   ' clamp negative gaps

        ft = wsD.Cells(i, cFT).value
        pt = wsD.Cells(i, cPT).value

        ' --- START with hold above thresholds ---
        If Not started Then
            If ft > startFlow And pt > startPress Then
                If startCandIdx = 0 And (prevFT <= startFlow Or prevPT <= startPress) Then
                    startCandIdx = i
                    holdAcc = 0
                    startedBeforeData = False
                End If
                holdAcc = holdAcc + dt
                If startCandIdx > 0 And holdAcc >= holdStartMin Then
                    If startedBeforeData Then
                        batchStartTime = "Started before data"
                    Else
                        batchStartTime = wsD.Cells(startCandIdx, cT).value
                    End If
                    started = True
                End If
            Else
                startCandIdx = 0
                holdAcc = 0
            End If
        End If

        ' --- END logic only after strip below AND return above (your original) ---
        If started Then
            ' find strip start (below) and end (above) from current position
            sStart = FindHoldBelow_FromSheet(wsD, cT, cPT, i, lastRow, stripPress, holdEachMin)
            If sStart > 0 Then
                bEnd = FindHoldAbove_FromSheet(wsD, cT, cPT, sStart + 1, lastRow, stripPress, holdEachMin)
                If bEnd > 0 Then
                    batchEndTime = wsD.Cells(bEnd, cT).value

                    ' Append completed batch
                    AppendBatchRow_WithStatus wsBS, tagLabel, batchStartTime, batchEndTime, _
                                              IIf(IsDate(batchStartTime), "Complete", "Partial Start")

                    ' jump index beyond the batch
                    i = bEnd + 1
                    started = False
                    startCandIdx = 0
                    holdAcc = 0
                    startedBeforeData = False
                    ' prime prev values and continue
                    If i <= lastRow Then
                        prevFT = wsD.Cells(i - 1, cFT).value
                        prevPT = wsD.Cells(i - 1, cPT).value
                    End If
                    GoTo ContinueLoop
                Else
                    ' no return above -> keep scanning; tail case handled after loop
                End If
            End If
        End If

        prevFT = ft
        prevPT = pt
ContinueLoop:
    Next i

    wsBS.Columns("A:G").AutoFit
End Sub

' -------- Simple hold finders against a sheet column --------
Public Function FindHoldBelow_FromSheet(ws As Worksheet, cT As Long, cCol As Long, _
    ByVal fromRow As Long, ByVal lastRow As Long, ByVal thresh As Double, ByVal holdMin As Double) As Long

    Dim i As Long, acc As Double, dt As Double, startIdx As Long
    For i = Application.Max(fromRow, 3) To lastRow
        dt = (ws.Cells(i, cT).value - ws.Cells(i - 1, cT).value) * 24# * 60#
        If ws.Cells(i, cCol).value < thresh Then
            If startIdx = 0 Then startIdx = i
            acc = acc + IIf(dt > 0, dt, 0)
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
            acc = acc + IIf(dt > 0, dt, 0)
            If acc >= holdMin Then FindHoldAbove_FromSheet = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0
        End If
    Next i
End Function

' -------- Helpers (append + de-dupe + ensure sheet) --------
Private Function EnsureBatchSummary() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Batch Summary")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Paste Data"))
        ws.name = "Batch Summary"
    End If
    Set EnsureBatchSummary = ws
End Function

Private Sub AppendBatchRowIfNew(ws As Worksheet, ByVal tagLabel As String, _
                                ByVal tStart As Variant, ByVal tEnd As Variant)
    If BatchRowExists(ws, tStart, tEnd) Then Exit Sub

    Dim r As Long: r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).value = tagLabel
    ws.Cells(r, 2).value = tStart
    ws.Cells(r, 3).value = tEnd
    ws.Cells(r, 4).value = DateDiff("n", tStart, tEnd)
    ws.Cells(r, 5).value = Round(ws.Cells(r, 4).value / 60#, 2)
    ws.Cells(r, 6).value = "Complete"
    ' col 7 (Product) intentionally left blank for manual entry
End Sub

Private Function BatchRowExists(ws As Worksheet, ByVal tStart As Variant, ByVal tEnd As Variant) As Boolean
    Dim last As Long: last = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    Dim r As Long
    For r = 2 To last
        If ws.Cells(r, 2).value = tStart And ws.Cells(r, 3).value = tEnd Then
            BatchRowExists = True
            Exit Function
        End If
    Next r
End Function

Private Sub AppendBatchRow_WithStatus(ws As Worksheet, ByVal tagLabel As String, _
                                      ByVal tStart As Variant, ByVal tEnd As Variant, _
                                      ByVal statusText As String)
    ' De-dupe: exact match on values already present
    If BatchRowExists(ws, tStart, tEnd) Then Exit Sub

    Dim r As Long: r = ws.Cells(ws.rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).value = tagLabel
    ws.Cells(r, 2).value = tStart
    ws.Cells(r, 3).value = tEnd

    If IsDate(tStart) And IsDate(tEnd) Then
        ws.Cells(r, 4).value = DateDiff("n", tStart, tEnd)
        ws.Cells(r, 5).value = Round(DateDiff("s", tStart, tEnd) / 3600#, 2)
    Else
        ws.Cells(r, 4).value = ""   ' no numeric duration when text placeholders present
        ws.Cells(r, 5).value = ""
    End If

    ws.Cells(r, 6).value = statusText
End Sub



