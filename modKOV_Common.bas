Attribute VB_Name = "modKOV_Common"
Option Explicit
Public G_KOV_Silent As Boolean

' Map a header key to a column, tolerant of ".Val" suffix/prefix forms
Public Function HeaderCol(hdr As Object, key As String) As Long
    Dim k As String, base As String
    k = Trim$(key)
    If hdr.Exists(k) Then
        HeaderCol = hdr(k)
        Exit Function
    End If
    If Right$(k, 4) = ".Val" Then
        base = Left$(k, Len(k) - 4)
        If hdr.Exists(base) Then HeaderCol = hdr(base)
    Else
        If hdr.Exists(k & ".Val") Then HeaderCol = hdr(k & ".Val")
    End If
End Function

Public Function CompositeMedian_AndStats( _
    wsData As Worksheet, ByRef t() As Double, hdr As Object, roleTags As Object, _
    ByVal roleKey As String, ByVal n As Long, _
    ByRef nSeries As Long, ByRef MaxDelta As Double, ByRef VarAcross As Double) As Double()

    Dim out() As Double
    nSeries = 0: MaxDelta = 0#: VarAcross = 0#

    If roleTags Is Nothing Or Not roleTags.Exists(roleKey) Then Exit Function
    Dim tags As Collection: Set tags = roleTags(roleKey)
    If tags Is Nothing Or tags.Count = 0 Then Exit Function

    Dim series As New Collection, idx As Long, i As Long
    For idx = 1 To tags.Count
        Dim c As Long: c = HeaderCol(hdr, CStr(tags(idx)))
        If c > 0 Then
            Dim s() As Double: ReDim s(1 To n)
            For i = 1 To n
                s(i) = CDbl(Val(wsData.Cells(i + 1, c).value))
            Next i
            series.Add s
        End If
    Next idx
    If series.Count = 0 Then Exit Function
    nSeries = series.Count

    ReDim out(1 To n)
    Dim arr As Variant, v As Double, haveNonZero As Boolean
    Dim vals() As Double, k As Long, mn As Double, mx As Double
    Dim varSum As Double, varCount As Long

    For i = 1 To n
        haveNonZero = False
        For Each arr In series
            If arr(i) <> 0 Then haveNonZero = True
        Next arr

        k = 0
        ReDim vals(1 To series.Count)
        mn = 0#: mx = 0#
        For Each arr In series
            v = arr(i)
            If haveNonZero And v = 0 Then
                ' skip zeros when any non-zero exists
            Else
                k = k + 1
                vals(k) = v
                If k = 1 Then
                    mn = v: mx = v
                Else
                    mn = IIf(v < mn, v, mn)
                    mx = IIf(v > mx, v, mx)
                End If
            End If
        Next arr

        If k = 0 Then
            out(i) = 0
        Else
            If k < UBound(vals) Then ReDim Preserve vals(1 To k)
            out(i) = Application.WorksheetFunction.Median(vals)
            If (mx - mn) > MaxDelta Then MaxDelta = (mx - mn)
            If k >= 2 Then
                varSum = varSum + PopulationVariance(vals, k)
                varCount = varCount + 1
            End If
        End If
    Next i

    If varCount > 0 Then VarAcross = varSum / varCount Else VarAcross = 0#
    CompositeMedian_AndStats = out
End Function

Private Function PopulationVariance(ByRef a() As Double, ByVal k As Long) As Double
    Dim i As Long, mu As Double, s As Double
    For i = 1 To k
        mu = mu + a(i)
    Next i
    mu = mu / k
    For i = 1 To k
        s = s + (a(i) - mu) * (a(i) - mu)
    Next i
    PopulationVariance = s / k
End Function

Public Function SeriesExists(ByRef a() As Double) As Boolean
    On Error GoTo nope
    SeriesExists = (UBound(a) >= 1)
    Exit Function
nope:
    SeriesExists = False
End Function

Public Function TrimmedMeanTW(ByRef y() As Double, ByRef t() As Double, _
                              ByVal i0 As Long, ByVal i1 As Long, _
                              ByVal trimIn As Double, ByVal trimOut As Double) As Double
    If i1 <= i0 Then Exit Function
    Dim s0 As Double, s1 As Double, i As Long
    s0 = t(i0): s1 = t(i1)
    If trimIn > 0# Then s0 = DateAdd("n", trimIn, s0)
    If trimOut > 0# Then s1 = DateAdd("n", -trimOut, s1)
    Dim num As Double, den As Double, dt As Double, yMid As Double
    For i = i0 + 1 To i1
        dt = MinutesBetween(t(i - 1), t(i))
        If dt < 0 Then dt = 0
        If t(i) <= s0 Or t(i - 1) >= s1 Then
            ' outside trimmed window
        Else
            yMid = (y(i) + y(i - 1)) / 2#
            If t(i) <= s1 And t(i - 1) >= s0 Then
                num = num + dt * yMid
                den = den + dt
            Else
                Dim tA As Double, tB As Double, dt2 As Double
                tA = Application.Max(t(i - 1), s0)
                tB = Application.min(t(i), s1)
                dt2 = MinutesBetween(tA, tB)
                If dt2 > 0 Then
                    num = num + dt2 * yMid
                    den = den + dt2
                End If
            End If
        End If
    Next i
    If den > 0 Then TrimmedMeanTW = num / den
End Function

Public Function TimeWeightedMeanWindow(ByRef y() As Double, ByRef t() As Double, _
                                       ByVal i0 As Long, ByVal i1 As Long, _
                                       ByVal nullZero As Double) As Double
    If i1 <= i0 Then Exit Function
    Dim num As Double, den As Double, dt As Double, i As Long
    For i = i0 + 1 To i1
        dt = MinutesBetween(t(i - 1), t(i))
        If dt < 0 Then dt = 0
        If y(i) <> nullZero And y(i - 1) <> nullZero Then
            num = num + dt * (y(i) + y(i - 1)) / 2#
            den = den + dt
        End If
    Next i
    If den > 0 Then TimeWeightedMeanWindow = num / den
End Function

Public Function MinutesBetween(ByVal t0 As Double, ByVal t1 As Double) As Double
    MinutesBetween = (t1 - t0) * 24# * 60#
End Function

Public Function HoursBetween(ByVal t0 As Double, ByVal t1 As Double) As Double
    HoursBetween = (t1 - t0) * 24#
End Function

Public Function HasLimit(wsLim As Worksheet, ByVal product As String, _
                         ByVal stepName As String, ByVal varKey As String) As Boolean
    Dim vMin As Double, vTV As Double, vMax As Double, vUnits As String, vLabel As String
    Dim hasMin As Boolean, hasTV As Boolean, hasMax As Boolean
    HasLimit = GetLimit_ByStepExact(wsLim, product, stepName, varKey, _
                                    vMin, vTV, vMax, vUnits, vLabel, hasMin, hasTV, hasMax)
End Function

Public Function WriteRow(wsK As Worksheet, ByVal rr As Long, _
                         ByVal stage As String, ByVal tStart As Double, ByVal tEnd As Double, _
                         ByVal metric As String, ByVal measured As Double, _
                         wsL As Worksheet, ByVal product As String, ByVal stepName As String, ByVal varKey As String, _
                         ByVal isTime As Boolean, ByVal notes As String) As Long
    wsK.Cells(rr, 1).value = stage
    wsK.Cells(rr, 2).value = tStart
    wsK.Cells(rr, 3).value = tEnd
    wsK.Cells(rr, 2).NumberFormat = "m/dd/yyyy hh:mm"
    wsK.Cells(rr, 3).NumberFormat = "m/dd/yyyy hh:mm"
    wsK.Cells(rr, 4).value = metric

    If measured <> 0 Then
        If isTime Then
            wsK.Cells(rr, 5).value = Round(measured, 2)   ' time in hours -> 2 decimals
            wsK.Cells(rr, 5).NumberFormat = "0.00"
        Else
            wsK.Cells(rr, 5).value = Round(measured, 1)   ' temps/rates -> 1 decimal
            wsK.Cells(rr, 5).NumberFormat = "0.0"
        End If
    End If

    wsK.Cells(rr, 12).value = notes

    WriteSpecCompare_ByStepExact wsK, rr, product, stepName, varKey, measured, wsL, isTime
    WriteRow = rr + 1
End Function

Private Function GetLimit_ByStepExact( _
    ws As Worksheet, ByVal product As String, _
    ByVal stepName As String, ByVal varKey As String, _
    ByRef vMin As Double, ByRef vTV As Double, ByRef vMax As Double, _
    Optional ByRef vUnits As String, Optional ByRef vLabel As String, _
    Optional ByRef hasMin As Boolean, Optional ByRef hasTV As Boolean, Optional ByRef hasMax As Boolean) As Boolean

    Dim r As Long, lastRow As Long
    Dim p As String, s As String, v As String
    Dim sMin As String, sTV As String, sMax As String

    p = LCase$(Trim$(product))
    s = LCase$(Trim$(stepName))
    v = LCase$(Trim$(varKey))

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If LCase$(Trim$(CStr(ws.Cells(r, 1).value))) = p _
        And LCase$(Trim$(CStr(ws.Cells(r, 2).value))) = s _
        And LCase$(Trim$(CStr(ws.Cells(r, 3).value))) = v Then

            vUnits = CStr(ws.Cells(r, 4).value)
            vLabel = CStr(ws.Cells(r, 8).value)

            sMin = Trim$(CStr(ws.Cells(r, 5).value))
            sTV = Trim$(CStr(ws.Cells(r, 6).value))
            sMax = Trim$(CStr(ws.Cells(r, 7).value))

            hasMin = (Len(sMin) > 0)
            hasTV = (Len(sTV) > 0)
            hasMax = (Len(sMax) > 0)

            If hasMin Then vMin = CDbl(Val(sMin)) Else vMin = 0
            If hasTV Then vTV = CDbl(Val(sTV)) Else vTV = 0
            If hasMax Then vMax = CDbl(Val(sMax)) Else vMax = 0

            GetLimit_ByStepExact = True
            Exit Function
        End If
    Next r
End Function

Private Sub WriteSpecCompare_ByStepExact( _
    ByVal wsOut As Worksheet, ByVal rowOut As Long, _
    ByVal product As String, ByVal stepName As String, ByVal varKey As String, _
    ByVal measured As Double, ByVal wsLim As Worksheet, ByVal isTime As Boolean)

    Dim vMin As Double, vTV As Double, vMax As Double, vUnits As String, vLabel As String
    Dim hasMin As Boolean, hasTV As Boolean, hasMax As Boolean
    Dim pf As String, pass As Boolean, nf As String

    If Not GetLimit_ByStepExact(wsLim, product, stepName, varKey, _
                                vMin, vTV, vMax, vUnits, vLabel, _
                                hasMin, hasTV, hasMax) Then Exit Sub

    ' ----- PASS/FAIL logic -----
    If measured = 0 Then
        pf = ""
    ElseIf hasMin And hasMax Then
        pass = (measured >= vMin And measured <= vMax): pf = IIf(pass, "PASS", "FAIL")
    ElseIf hasMax And Not hasMin Then
        pass = (measured <= vMax): pf = IIf(pass, "PASS", "FAIL")
    ElseIf hasMin And Not hasMax Then
        pass = (measured >= vMin): pf = IIf(pass, "PASS", "FAIL")
    Else
        pf = ""
    End If

    ' ----- Write numbers -----
    wsOut.Cells(rowOut, 6).value = IIf(hasMin, vMin, "") ' Min (F)
    wsOut.Cells(rowOut, 7).value = IIf(hasTV, vTV, "")   ' TV  (G)
    wsOut.Cells(rowOut, 8).value = IIf(hasMax, vMax, "") ' Max (H)
    wsOut.Cells(rowOut, 9).value = pf                    ' Result (I)

    If hasTV And measured <> 0 Then
        wsOut.Cells(rowOut, 10).value = measured - vTV   ' # from TV (J)
    Else
        wsOut.Cells(rowOut, 10).ClearContents
    End If

    wsOut.Cells(rowOut, 11).value = vLabel               ' Label (K)

    ' ----- Number formats (consistent with rounding you added) -----
    nf = IIf(isTime, "0.00", "0.0")          ' hours to 2dp, else 1dp
    If hasMin Then wsOut.Cells(rowOut, 6).NumberFormat = nf
    If hasTV Then wsOut.Cells(rowOut, 7).NumberFormat = nf
    If hasMax Then wsOut.Cells(rowOut, 8).NumberFormat = nf
    If wsOut.Cells(rowOut, 10).value <> "" Then wsOut.Cells(rowOut, 10).NumberFormat = nf

End Sub

Public Sub KOV_ColorizeAllTables(ByVal ws As Worksheet)
    Dim lastRow As Long, r As Long, v As String
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' Clear any prior fill/fonts in key columns
    ws.Range(ws.Cells(2, 9), ws.Cells(lastRow, 9)).Interior.Pattern = xlNone   ' Result (col I)
    ws.Range(ws.Cells(2, 9), ws.Cells(lastRow, 9)).Font.Color = vbBlack
    ws.Range(ws.Cells(2, 11), ws.Cells(lastRow, 11)).Interior.Pattern = xlNone ' Label (col K)

    ' Color PASS/FAIL in Result (I)
    For r = 2 To lastRow
        v = UCase$(Trim$(CStr(ws.Cells(r, 9).value)))
        Select Case v
            Case "PASS"
                ws.Cells(r, 9).Interior.Color = RGB(198, 239, 206)
                ws.Cells(r, 9).Font.Color = RGB(0, 97, 0)
            Case "FAIL"
                ws.Cells(r, 9).Interior.Color = RGB(255, 199, 206)
                ws.Cells(r, 9).Font.Color = RGB(156, 0, 6)
        End Select
    Next r

    ' Shade KOV/AOV in Label (K)
    For r = 2 To lastRow
        v = UCase$(Trim$(CStr(ws.Cells(r, 11).value)))
        Select Case v
            Case "KOV": ws.Cells(r, 11).Interior.Color = RGB(198, 239, 206)
            Case "AOV": ws.Cells(r, 11).Interior.Color = RGB(226, 239, 218)
        End Select
    Next r

    ' Bold header rows of each pasted table (row where col A = "Stage")
    For r = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, 1).value))) = "STAGE" Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)).Font.Bold = True
        End If
    Next r
End Sub

Public Function PrintRoleSummary(ws As Worksheet, ByVal startRow As Long, _
                                 ByVal roleKey As String, roleTags As Object, _
                                 ByVal nUsed As Long, ByVal MaxDelta As Double, ByVal VarAcross As Double, _
                                 Optional ByVal product As String = "") As Long
    Dim tags As Collection: Set tags = roleTags(roleKey)
    If Not tags Is Nothing And tags.Count > 0 Then
        Dim j As Long, s As String
        For j = 1 To tags.Count
            s = s & IIf(Len(s) > 0, ", ", "") & CStr(tags(j))
        Next j

        If startRow = 2 And Len(product) > 0 Then ws.Cells(startRow, 1).value = product
        ws.Cells(startRow, 2).value = roleKey
        ws.Cells(startRow, 3).value = s
        ws.Cells(startRow, 4).value = nUsed
        ws.Cells(startRow, 5).value = Round(MaxDelta, 2)
        ws.Cells(startRow, 6).value = Round(Sqr(VarAcross), 3)

        PrintRoleSummary = startRow + 1
    Else
        PrintRoleSummary = startRow
    End If
End Function

Private Sub KOV_ClearSheet(ws As Worksheet)
    With ws
        .Cells.Clear
        .Cells.FormatConditions.Delete
        .Cells.Interior.Pattern = xlNone
        .Cells.Borders.LineStyle = xlNone
    End With
End Sub

Public Function BuildHeaderIndexAll(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, key$
    For c = 1 To lastCol
        key = Trim$(CStr(ws.Cells(1, c).Value2))
        If Len(key) > 0 Then d(key) = c
    Next c
    Set BuildHeaderIndexAll = d
End Function

Public Function GroupTagsByRole_Explicit(wsMap As Worksheet, ByVal product As String, hdr As Object) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    ' Pre-create common roles (ok if some stay empty)
    Dim roles As Variant, rName As Variant
    roles = Array("TT", "PT", "FT", "CFT", "MFT", "MTT", "PFT", "DFT", "AFT")
    For Each rName In roles
        Set d(rName) = New Collection
    Next

    Dim lastRow As Long: lastRow = wsMap.Cells(wsMap.Rows.Count, 1).End(xlUp).Row
    Dim r As Long, prod$, tag$, role$, tagHeader$

    For r = 2 To lastRow
        prod = Trim$(CStr(wsMap.Cells(r, 1).value))
        If Len(prod) = 0 Or StrComp(prod, product, vbTextCompare) <> 0 Then GoTo nxt

        tag = Trim$(CStr(wsMap.Cells(r, 2).value))
        role = UCase$(Trim$(CStr(wsMap.Cells(r, 3).value)))
        If Len(tag) = 0 Or Len(role) = 0 Then GoTo nxt

        If HeaderCol(hdr, tag) > 0 Then
            tagHeader = tag
        ElseIf HeaderCol(hdr, tag & ".Val") > 0 Then
            tagHeader = tag & ".Val"
        Else
            GoTo nxt
        End If

        If Not d.Exists(role) Then Set d(role) = New Collection
        d(role).Add tagHeader
nxt:
    Next r

    Set GroupTagsByRole_Explicit = d
End Function


Public Sub KOV_Notify(ByVal msg As String)
    If Not G_KOV_Silent Then MsgBox msg, vbInformation
End Sub

