Attribute VB_Name = "BatchTime"
Sub ExtractBatchTimesFromWI()
    Dim ws As Worksheet, summaryWs As Worksheet
    Dim lastRow As Long, Col As Long, outputRow As Long
    Dim i As Long
    Dim wiValue As Double, prevValue As Double
    Dim batchStartTime As Variant, batchEndTime As Variant
    Dim started As Boolean

    ' Set your data worksheet
    Set ws = ThisWorkbook.Sheets("Paste Data")

    ' Create or clear the summary sheet
    On Error Resume Next
    Set summaryWs = ThisWorkbook.Sheets("Batch Summary")
    If summaryWs Is Nothing Then
        Set summaryWs = ThisWorkbook.Sheets.Add(After:=ws)
        summaryWs.name = "Batch Summary"
    Else
        summaryWs.Cells.ClearContents
    End If
    On Error GoTo 0

    ' Set headers in the summary sheet
    summaryWs.Range("A1:F1").value = Array("Tag", "Batch Start", "Batch End", "Duration (min)", "Duration (hr)", "Status")
    outputRow = 2

    ' Find last row of data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "Not enough rows in Paste Data.", vbExclamation
        Exit Sub
    End If

    ' Parameters
    Const thresh As Double = 1000
    Const HOLD_MIN As Double = 300

    ' Loop through each column to find WI tags
    For Col = 2 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If InStr(1, ws.Cells(1, Col).value, "WI", vbTextCompare) > 0 Then

            ' --- per-tag state ---
            started = False
            prevValue = ws.Cells(2, Col).value
            batchStartTime = Empty
            batchEndTime = Empty

            ' hold tracking
            Dim holdAcc As Double: holdAcc = 0          ' minutes accumulated >THRESH
            Dim startCandIdx As Long: startCandIdx = 0  ' row index of first >THRESH in current run
            Dim startedBeforeData As Boolean: startedBeforeData = False

            ' >>> Seed candidate if series begins above threshold
            If prevValue > thresh Then
                startCandIdx = 2
                holdAcc = 0
                startedBeforeData = True
            End If

            For i = 3 To lastRow
                wiValue = ws.Cells(i, Col).value

                ' === START DETECTION WITH HOLD ===
                If Not started Then
                    If wiValue > thresh Then
                        If startCandIdx = 0 And prevValue <= thresh Then
                            startCandIdx = i
                            holdAcc = 0
                            startedBeforeData = False
                        End If
                        holdAcc = holdAcc + (ws.Cells(i, 1).value - ws.Cells(i - 1, 1).value) * 24# * 60#
                        If startCandIdx > 0 And holdAcc >= HOLD_MIN Then
                            If startedBeforeData Then
                                batchStartTime = "Started before data"
                            Else
                                batchStartTime = ws.Cells(startCandIdx, 1).value
                            End If
                            started = True
                        End If
                    Else
                        startCandIdx = 0
                        holdAcc = 0
                    End If
                End If

                ' === END DETECTION (unchanged logic) ===
                If started Then
                    If wiValue <= thresh And prevValue > thresh Then
                        batchEndTime = ws.Cells(i, 1).value

                        summaryWs.Cells(outputRow, 1).value = ws.Cells(1, Col).value
                        summaryWs.Cells(outputRow, 3).value = batchEndTime

                        If IsDate(batchStartTime) Then
                            summaryWs.Cells(outputRow, 2).value = batchStartTime
                            summaryWs.Cells(outputRow, 4).value = DateDiff("n", batchStartTime, batchEndTime)
                            summaryWs.Cells(outputRow, 5).value = Round(DateDiff("s", batchStartTime, batchEndTime) / 3600, 2)
                            summaryWs.Cells(outputRow, 6).value = "Complete"
                        Else
                            ' Started before data (or non-date label)
                            summaryWs.Cells(outputRow, 2).value = "Started before data"
                            summaryWs.Cells(outputRow, 4).value = DateDiff("n", ws.Cells(2, 1).value, batchEndTime)
                            summaryWs.Cells(outputRow, 5).value = Round(DateDiff("s", ws.Cells(2, 1).value, batchEndTime) / 3600, 2)
                            summaryWs.Cells(outputRow, 6).value = "Partial Start"
                        End If

                        outputRow = outputRow + 1
                        started = False
                        startCandIdx = 0
                        holdAcc = 0
                        startedBeforeData = False
                    End If
                End If

                prevValue = wiValue
            Next i

            ' Handle batch still running at end of dataset
            If started Then
                summaryWs.Cells(outputRow, 1).value = ws.Cells(1, Col).value
                summaryWs.Cells(outputRow, 3).value = "Ends after data"

                If IsDate(batchStartTime) Then
                    summaryWs.Cells(outputRow, 2).value = batchStartTime
                    summaryWs.Cells(outputRow, 4).value = DateDiff("n", batchStartTime, ws.Cells(lastRow, 1).value)
                    summaryWs.Cells(outputRow, 5).value = Round(DateDiff("s", batchStartTime, ws.Cells(lastRow, 1).value) / 3600, 2)
                    summaryWs.Cells(outputRow, 6).value = "Partial End"
                Else
                    summaryWs.Cells(outputRow, 2).value = "Started before data"
                    summaryWs.Cells(outputRow, 4).value = DateDiff("n", ws.Cells(2, 1).value, ws.Cells(lastRow, 1).value)
                    summaryWs.Cells(outputRow, 5).value = Round(DateDiff("s", ws.Cells(2, 1).value, ws.Cells(lastRow, 1).value) / 3600, 2)
                    summaryWs.Cells(outputRow, 6).value = "Started before data + Ends after data"
                End If

                outputRow = outputRow + 1

            ' >>> Optional: series was above threshold but hold never confirmed before EOF
            ElseIf (startCandIdx > 0 Or prevValue > thresh) Then
                summaryWs.Cells(outputRow, 1).value = ws.Cells(1, Col).value
                summaryWs.Cells(outputRow, 2).value = "Started before data or hold<" & HOLD_MIN & "m not confirmed"
                summaryWs.Cells(outputRow, 3).value = "Ends after data"
                summaryWs.Cells(outputRow, 4).value = DateDiff("n", ws.Cells(2, 1).value, ws.Cells(lastRow, 1).value)
                summaryWs.Cells(outputRow, 5).value = Round(DateDiff("s", ws.Cells(2, 1).value, ws.Cells(lastRow, 1).value) / 3600, 2)
                summaryWs.Cells(outputRow, 6).value = "Unconfirmed (no hold) + Ends after data"
                outputRow = outputRow + 1
            End If

        End If
    Next Col

    MsgBox "Batch times extracted to 'Batch Summary' sheet.", vbInformation
End Sub

Private Function FindHoldBelow_FromSheet(ws As Worksheet, cT As Long, cCol As Long, _
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

Private Function FindHoldAbove_FromSheet(ws As Worksheet, cT As Long, cCol As Long, _
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

