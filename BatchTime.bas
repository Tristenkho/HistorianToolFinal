Attribute VB_Name = "BatchTime"
Option Explicit

' Map any WI header to a simple reactor label
Private Function ReactorLabel(ByVal header As String) As String
    Dim s As String: s = UCase$(Trim$(header))
    s = Replace(s, " ", "")
    ' easy prefixes like R1_WI_01.Val, R2WI, etc.
    If Left$(s, 2) = "R1" Then ReactorLabel = "R1": Exit Function
    If Left$(s, 2) = "R2" Then ReactorLabel = "R2": Exit Function
    If Left$(s, 2) = "R3" Then ReactorLabel = "R3": Exit Function
    If Left$(s, 2) = "R4" Then ReactorLabel = "R4": Exit Function
    ' fallback: look anywhere
    If InStr(s, "R1") > 0 Then ReactorLabel = "R1": Exit Function
    If InStr(s, "R2") > 0 Then ReactorLabel = "R2": Exit Function
    If InStr(s, "R3") > 0 Then ReactorLabel = "R3": Exit Function
    If InStr(s, "R4") > 0 Then ReactorLabel = "R4": Exit Function
    ReactorLabel = header
End Function

Public Sub ExtractBatchTimesFromWI()
    Dim ws As Worksheet, summaryWs As Worksheet
    Dim lastRow As Long, col As Long, outputRow As Long
    Dim i As Long
    Dim wiValue As Double, prevValue As Double
    Dim batchStartTime As Variant, batchEndTime As Variant
    Dim started As Boolean

    Set ws = ThisWorkbook.Sheets("Paste Data")

    ' Ensure/prepare summary sheet WITHOUT clearing existing rows (so R4 stays)
    On Error Resume Next
    Set summaryWs = ThisWorkbook.Sheets("Batch Summary")
    On Error GoTo 0
    If summaryWs Is Nothing Then
        Set summaryWs = ThisWorkbook.Sheets.Add(After:=ws)
        summaryWs.name = "Batch Summary"
    End If
    ' Ensure headers (A..G to match KOV runner expectations; G=Product left blank)
    If summaryWs.Cells(1, 1).value <> "Tag" Then
        summaryWs.Range("A1:G1").value = Array("Tag", "Batch Start", "Batch End", _
                                               "Duration (min)", "Duration (hr)", "Status", "Product")
    End If
    ' Append to first empty row
    outputRow = summaryWs.Cells(summaryWs.rows.Count, 1).End(xlUp).Row + 1

    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "Not enough rows in Paste Data.", vbExclamation
        Exit Sub
    End If

    Const thresh As Double = 1000
    Const HOLD_MIN As Double = 300   ' minutes

    For col = 2 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If InStr(1, ws.Cells(1, col).value, "WI", vbTextCompare) > 0 Then

            started = False
            prevValue = ws.Cells(2, col).value
            batchStartTime = Empty
            batchEndTime = Empty

            Dim holdAcc As Double: holdAcc = 0
            Dim startCandIdx As Long: startCandIdx = 0
            Dim startedBeforeData As Boolean: startedBeforeData = False

            If prevValue > thresh Then
                startCandIdx = 2
                holdAcc = 0
                startedBeforeData = True
            End If

            For i = 3 To lastRow
                wiValue = ws.Cells(i, col).value

                ' --- START with hold ---
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

                ' --- END on falling back below threshold ---
                If started Then
                    If wiValue <= thresh And prevValue > thresh Then
                        batchEndTime = ws.Cells(i, 1).value

                        summaryWs.Cells(outputRow, 1).value = ReactorLabel(ws.Cells(1, col).value) ' << normalize
                        summaryWs.Cells(outputRow, 3).value = batchEndTime

                        If IsDate(batchStartTime) Then
                            summaryWs.Cells(outputRow, 2).value = batchStartTime
                            summaryWs.Cells(outputRow, 4).value = DateDiff("n", batchStartTime, batchEndTime)
                            summaryWs.Cells(outputRow, 5).value = Round(DateDiff("s", batchStartTime, batchEndTime) / 3600, 2)
                            summaryWs.Cells(outputRow, 6).value = "Complete"
                        Else
                            summaryWs.Cells(outputRow, 2).value = "Started before data"
                            summaryWs.Cells(outputRow, 4).value = DateDiff("n", ws.Cells(2, 1).value, batchEndTime)
                            summaryWs.Cells(outputRow, 5).value = Round(DateDiff("s", ws.Cells(2, 1).value, batchEndTime) / 3600, 2)
                            summaryWs.Cells(outputRow, 6).value = "Partial Start"
                        End If

                        summaryWs.Cells(outputRow, 2).NumberFormat = "m/dd/yyyy hh:mm"
                        summaryWs.Cells(outputRow, 3).NumberFormat = "m/dd/yyyy hh:mm"

                        outputRow = outputRow + 1
                        started = False
                        startCandIdx = 0
                        holdAcc = 0
                        startedBeforeData = False
                    End If
                End If

                prevValue = wiValue
            Next i

            ' --- tail cases ---
            If started Then
                summaryWs.Cells(outputRow, 1).value = ReactorLabel(ws.Cells(1, col).value)
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
                summaryWs.Cells(outputRow, 2).NumberFormat = "m/dd/yyyy hh:mm"
                outputRow = outputRow + 1

            ElseIf (startCandIdx > 0 Or prevValue > thresh) Then
                summaryWs.Cells(outputRow, 1).value = ReactorLabel(ws.Cells(1, col).value)
                summaryWs.Cells(outputRow, 2).value = "Started before data or hold<" & HOLD_MIN & "m not confirmed"
                summaryWs.Cells(outputRow, 3).value = "Ends after data"
                summaryWs.Cells(outputRow, 4).value = DateDiff("n", ws.Cells(2, 1).value, ws.Cells(lastRow, 1).value)
                summaryWs.Cells(outputRow, 5).value = Round(DateDiff("s", ws.Cells(2, 1).value, ws.Cells(lastRow, 1).value) / 3600, 2)
                summaryWs.Cells(outputRow, 6).value = "Unconfirmed (no hold) + Ends after data"
                outputRow = outputRow + 1
            End If

        End If
    Next col

    summaryWs.Columns("A:G").AutoFit
    MsgBox "Batch times appended to 'Batch Summary'.", vbInformation
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

