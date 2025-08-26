Attribute VB_Name = "KOV_ASA"
Option Explicit

'======================== SHEETS ========================
Private Const SH_DATA   As String = "Paste Data"
Private Const SH_LIMITS As String = "Product Limits"
Private Const SH_TAGMAP As String = "Tag Map"
Private Const SH_KOV    As String = "KOV"

'======================== PRODUCT =======================
Private Const PRODUCT_NAME As String = "Innospec ASA"     ' R2 recipe

'======================== ROLES (Tag Map roles) =========
' TT  -> Reactor temperature (R2_TT_01 / R2_TT_02)
' PT  -> Reactor pressure   (R2_PT_01 / R2_PT_02)
' MFT -> Maleic flow        (T234_FT_02)
' CFT -> Cooling flow       (R2_E3_FT_01)
' MTT -> Maleic temp        (T234_TT_01 / T234_TT_02)
Private Const ROLE_TT  As String = "TT"
Private Const ROLE_PT  As String = "PT"
Private Const ROLE_MFT As String = "MFT"
Private Const ROLE_CFT As String = "CFT"
Private Const ROLE_MTT As String = "MTT"

'======================== THRESHOLDS ====================
' All holds & trims = 10 min
Private Const HOLD_MIN          As Double = 10
Private Const TRIM_IN_MIN       As Double = 10
Private Const TRIM_OUT_MIN      As Double = 10

' Maleic flow band for charge window
Private Const MFT_BAND_LO       As Double = 100
Private Const MFT_BAND_HI       As Double = 130

' Soak end / Strip start rule (pressure at/below atm)
Private Const PT_ATM            As Double = 14.7

' Strip end rule (CFT baseline + delta)
Private Const CFT_DELTA_RISE    As Double = 150

'=======================================================
'                    PUBLIC ENTRY
'=======================================================
Public Sub KOV_Run_InnospecASA_Main()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsD As Worksheet, wsL As Worksheet, wsM As Worksheet, wsK As Worksheet

    On Error Resume Next
    Set wsD = wb.Worksheets(SH_DATA)
    Set wsL = wb.Worksheets(SH_LIMITS)
    Set wsM = wb.Worksheets(SH_TAGMAP)
    Set wsK = wb.Worksheets(SH_KOV)
    On Error GoTo 0

    If wsD Is Nothing Or wsL Is Nothing Or wsM Is Nothing Then
        MsgBox "Missing sheet(s). Need: Paste Data, Product Limits, Tag Map.", vbCritical
        Exit Sub
    End If
    If wsK Is Nothing Then
        Set wsK = wb.Worksheets.Add(After:=wsD)
        wsK.name = SH_KOV
    End If

    ' ---- headers / time ----
    Dim hdr As Object: Set hdr = BuildHeaderIndexAll(wsD)
    Dim cT As Long: cT = HeaderCol(hdr, "Time")
    If cT = 0 Then
        MsgBox "Missing 'Time' header in Paste Data.", vbCritical
        Exit Sub
    End If

    Dim n As Long, t() As Double
    If Not BuildTimeVector(wsD, cT, t, n) Then
        MsgBox "Time column not recognized as date/time.", vbCritical
        Exit Sub
    End If

    ' ---- restrict to Batch Summary window if provided ----
    Dim i0 As Long, i1 As Long
    ResolveWindowBoundsLocal t, i0, i1

    ' ---- roles -> tags for THIS product ----
    Dim roleTags As Object
    Set roleTags = GroupTagsByRole_Explicit(wsM, PRODUCT_NAME, hdr)

    ' ---- composites + redundancy stats ----
    Dim TT() As Double, pt() As Double, MFT() As Double, cFT() As Double, MTT() As Double
    Dim nTT As Long, nPT As Long, nMFT As Long, nCFT As Long, nMTT As Long
    Dim dTT As Double, dPT As Double, dMFT As Double, dCFT As Double, dMTT As Double
    Dim vTT As Double, vPT As Double, vMFT As Double, vCFT As Double, vMTT As Double

    TT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_TT, n, nTT, dTT, vTT)
    pt = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_PT, n, nPT, dPT, vPT)
    MFT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_MFT, n, nMFT, dMFT, vMFT)
    cFT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_CFT, n, nCFT, dCFT, vCFT)
    MTT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_MTT, n, nMTT, dMTT, vMTT)

    If Not SeriesExists(TT) Or Not SeriesExists(pt) Or Not SeriesExists(MFT) _
       Or Not SeriesExists(cFT) Or Not SeriesExists(MTT) Then
        MsgBox "Required roles missing (TT/PT/MFT/CFT/MTT) for '" & PRODUCT_NAME & "'.", vbCritical
        Exit Sub
    End If

    ' ---- output header ----
    wsK.Cells.ClearContents
    wsK.Range("A1:F1").value = Array("Product", "Role", "Tags used", "Redundancy (N)", "Redundancy (Max)", "Redundancy (StdDev)")
    wsK.Range("A2").value = PRODUCT_NAME
    Dim rr As Long: rr = 2
    rr = PrintRoleSummary(wsK, rr, ROLE_TT, roleTags, nTT, dTT, vTT, PRODUCT_NAME)
    rr = PrintRoleSummary(wsK, rr, ROLE_PT, roleTags, nPT, dPT, vPT)
    rr = PrintRoleSummary(wsK, rr, ROLE_MFT, roleTags, nMFT, dMFT, vMFT)
    rr = PrintRoleSummary(wsK, rr, ROLE_CFT, roleTags, nCFT, dCFT, vCFT)
    rr = PrintRoleSummary(wsK, rr, ROLE_MTT, roleTags, nMTT, dMTT, vMTT)

    rr = rr + 2
    wsK.Rows(rr - 1).RowHeight = 8
    wsK.Range("A" & rr & ":L" & rr).value = Array("Stage", "Start Time", "End Time", _
                                                  "Metric", "Value", "Min", "TV", "Max", _
                                                  "Result", "# from TV", "Label", "Notes")
    wsK.Range("A" & rr & ":L" & rr).Font.Bold = True
    rr = rr + 1

    '==================== DETECTION ====================

    ' ---- Maleic charge window (MFT band) ----
    Dim iM_Start As Long, iM_End As Long
    iM_Start = Find_MaleicStart_InBandRange(MFT, t, MFT_BAND_LO, MFT_BAND_HI, HOLD_MIN, i0, i1)
    If iM_Start > 0 Then
        iM_End = Find_MaleicEnd_OutOfBandRange(MFT, t, MFT_BAND_LO, MFT_BAND_HI, HOLD_MIN, iM_Start + 1, i1)
    End If

    If iM_Start > 0 And iM_End > iM_Start Then
        ' ---- Charge metrics ----
        Dim mttCharge As Double, ttCharge As Double, maleicH As Double
        mttCharge = TrimmedMeanTW(MTT, t, iM_Start, iM_End, TRIM_IN_MIN, TRIM_OUT_MIN)   ' KOV: Charge Temp = MTT
        ttCharge = TrimmedMeanTW(TT, t, iM_Start, iM_End, TRIM_IN_MIN, TRIM_OUT_MIN)     ' optional diagnostic
        maleicH = HoursBetween(t(iM_Start), t(iM_End))

        If HasLimit(wsL, PRODUCT_NAME, "Maleic Charge", "Charge Temperature") Then
            rr = WriteRow(wsK, rr, "Maleic Charge", t(iM_Start), t(iM_End), _
                          "Charge Temperature (F)", Round(mttCharge, 1), _
                          wsL, PRODUCT_NAME, "Maleic Charge", "Charge Temperature", False, _
                          "MFT in 100–130 for 10m; end out-of-band 10m. Charge Temp = MTT TW-mean (trim 10/10).")
        End If

        If HasLimit(wsL, PRODUCT_NAME, "Maleic Charge", "Time") Then
            rr = WriteRow(wsK, rr, "Maleic Charge", t(iM_Start), t(iM_End), _
                          "Time (h)", Round(maleicH, 2), _
                          wsL, PRODUCT_NAME, "Maleic Charge", "Time", True, _
                          "Duration from first in-band hold to first sustained out-of-band hold.")
        End If

        ' ---- Soak (Start = maleic end; End = PT = 14.7 for 10m) ----
        Dim iSoak_Start As Long, iSoak_End As Long
        iSoak_Start = iM_End
        iSoak_End = Find_FirstHold_Single_Range(pt, t, "<=", PT_ATM, HOLD_MIN, iSoak_Start + 1, i1)

        If iSoak_End > iSoak_Start Then
            Dim soakT As Double, soakH As Double
            soakT = TrimmedMeanTW(TT, t, iSoak_Start, iSoak_End, TRIM_IN_MIN, TRIM_OUT_MIN)
            soakH = HoursBetween(t(iSoak_Start), t(iSoak_End))

            If HasLimit(wsL, PRODUCT_NAME, "Soak", "Temperature") Then
                rr = WriteRow(wsK, rr, "Soak", t(iSoak_Start), t(iSoak_End), _
                              "Temperature (F)", Round(soakT, 1), _
                              wsL, PRODUCT_NAME, "Soak", "Temperature", False, _
                              "Start=Maleic end; End=PT =14.7 psia for 10m. TT mean (trim 10/10).")
            End If
            If HasLimit(wsL, PRODUCT_NAME, "Soak", "Time") Then
                rr = WriteRow(wsK, rr, "Soak", t(iSoak_Start), t(iSoak_End), _
                              "Time (h)", Round(soakH, 2), _
                              wsL, PRODUCT_NAME, "Soak", "Time", True, _
                              "Window between maleic end and PT <14.7 psia.")
            End If

            ' ---- Strip (Start = same PT hold; End = CFT = baseline+150 for 10m) ----
            Dim iStrip_Start As Long, iStrip_End As Long
            iStrip_Start = iSoak_End
            iStrip_End = Find_CFT_DeltaRiseHold_FromBaseline_Range(cFT, t, iStrip_Start, CFT_DELTA_RISE, HOLD_MIN, i1)

            If iStrip_End = 0 And G_KOV_UseWindow Then
                ' If not found, let it end at the window end (graceful fallback)
                iStrip_End = i1
            End If

            If iStrip_End > iStrip_Start Then
                Dim stripT As Double, stripH As Double
                stripT = TrimmedMeanTW(TT, t, iStrip_Start, iStrip_End, TRIM_IN_MIN, TRIM_OUT_MIN)
                stripH = HoursBetween(t(iStrip_Start), t(iStrip_End))

                If HasLimit(wsL, PRODUCT_NAME, "Strip", "Temperature") Then
                    rr = WriteRow(wsK, rr, "Strip", t(iStrip_Start), t(iStrip_End), _
                                  "Temperature (F)", Round(stripT, 1), _
                                  wsL, PRODUCT_NAME, "Strip", "Temperature", False, _
                                  "Start=PT =14.7 psia for 10m; End=CFT = (start+150) for 10m. TT TW-mean (trim 10/10).")
                End If
                If HasLimit(wsL, PRODUCT_NAME, "Strip", "Time") Then
                    rr = WriteRow(wsK, rr, "Strip", t(iStrip_Start), t(iStrip_End), _
                                  "Time (h)", Round(stripH, 2), _
                                  wsL, PRODUCT_NAME, "Strip", "Time", True, _
                                  "From PT hold to CFT delta-rise hold.")
                End If
            End If
        Else
            MsgBox "Soak end (PT =14.7 psia for 10m) not found within window.", vbExclamation
        End If
    Else
        MsgBox "Maleic window not found (MFT 100–130 for 10m, then out-of-band for 10m).", vbExclamation
    End If

    KOV_ColorizeAllTables wsK
    wsK.Columns("A:L").AutoFit
    KOV_Notify "KOV complete for '" & PRODUCT_NAME & "'."
End Sub

'=======================================================
'                HEADERS / TIME / TAG MAP
'=======================================================
Private Function BuildHeaderIndexAll(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, key$
    For c = 1 To lastCol
        key = Trim$(CStr(ws.Cells(1, c).Value2))
        If Len(key) > 0 Then d(key) = c
    Next c
    Set BuildHeaderIndexAll = d
End Function

Private Function HeaderCol(hdr As Object, key As String) As Long
    Dim k$: k = Trim$(key)
    If hdr.Exists(k) Then HeaderCol = hdr(k)
End Function

Private Function BuildTimeVector(ws As Worksheet, ByVal cTime As Long, ByRef t() As Double, ByRef n As Long) As Boolean
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, cTime).End(xlUp).Row
    If lastRow < 3 Then Exit Function
    n = lastRow - 1
    ReDim t(1 To n)

    Dim i As Long, v As Variant
    For i = 1 To n
        v = ws.Cells(i + 1, cTime).value
        If IsDate(v) Then
            t(i) = CDbl(CDate(v))
        ElseIf IsNumeric(v) Then
            t(i) = CDbl(v)
        Else
            t(i) = 0
        End If
    Next i

    Dim tot As Double
    For i = 2 To n
        Dim dt As Double: dt = MinutesBetween(t(i - 1), t(i))
        If dt > 0 Then tot = tot + dt
    Next i
    BuildTimeVector = (tot > 0.5)
End Function

Private Function GroupTagsByRole_Explicit(wsMap As Worksheet, ByVal product As String, hdr As Object) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    Set d(ROLE_TT) = New Collection
    Set d(ROLE_PT) = New Collection
    Set d(ROLE_MFT) = New Collection
    Set d(ROLE_MTT) = New Collection
    Set d(ROLE_CFT) = New Collection

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

        If d.Exists(role) Then d(role).Add tagHeader
nxt:
    Next r

    Set GroupTagsByRole_Explicit = d
End Function

'=======================================================
'                WINDOW RESOLUTION
'=======================================================
Private Sub ResolveWindowBoundsLocal(ByRef t() As Double, ByRef i0 As Long, ByRef i1 As Long)
    ' Uses global window flags set by WeekRunner
    On Error Resume Next
    If G_KOV_UseWindow Then
        Dim s As Double, e As Double
        s = G_KOV_WindowStart: e = G_KOV_WindowEnd
        If e > 0 And e < s Then e = s

        Dim n As Long: n = UBound(t)
        Dim i As Long
        i0 = 1: i1 = n
        For i = 1 To n
            If t(i) >= s Then i0 = Application.Max(1, i - 1): Exit For
        Next i
        For i = n To 1 Step -1
            If t(i) <= e Or e = 0 Then i1 = i: Exit For
        Next i
        If i1 < i0 Then
            i0 = 1: i1 = n
        End If
    Else
        i0 = 1: i1 = UBound(t)
    End If
End Sub

'=======================================================
'                DETECTION (local helpers)
'=======================================================
Private Function Find_MaleicStart_InBandRange( _
    ByRef v() As Double, ByRef t() As Double, _
    ByVal lo As Double, ByVal hi As Double, _
    ByVal holdMin As Double, ByVal i0 As Long, ByVal i1 As Long) As Long

    Dim i As Long, acc As Double, dt As Double, startIdx As Long: startIdx = 0
    For i = Application.Max(i0, 2) To i1
        dt = MinutesBetween(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If v(i) >= lo And v(i) <= hi Then
            If startIdx = 0 Then startIdx = i
            acc = acc + dt
            If acc >= holdMin Then Find_MaleicStart_InBandRange = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0
        End If
    Next i
End Function

Private Function Find_MaleicEnd_OutOfBandRange( _
    ByRef v() As Double, ByRef t() As Double, _
    ByVal lo As Double, ByVal hi As Double, _
    ByVal holdMin As Double, ByVal fromIdx As Long, ByVal i1 As Long) As Long

    If fromIdx <= 0 Then Exit Function
    Dim i As Long, acc As Double, dt As Double, startIdx As Long: startIdx = 0
    For i = Application.Max(fromIdx, 2) To i1
        dt = MinutesBetween(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If Not (v(i) >= lo And v(i) <= hi) Then
            If startIdx = 0 Then startIdx = i
            acc = acc + dt
            If acc >= holdMin Then Find_MaleicEnd_OutOfBandRange = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0
        End If
    Next i
End Function

Private Function Find_FirstHold_Single_Range( _
    ByRef v() As Double, ByRef t() As Double, _
    ByVal op As String, ByVal thresh As Double, _
    ByVal holdMin As Double, ByVal fromIdx As Long, ByVal i1 As Long) As Long

    Dim i As Long, acc As Double, dt As Double, startIdx As Long: startIdx = 0
    For i = Application.Max(fromIdx, 2) To i1
        dt = MinutesBetween(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If (op = "<=" And v(i) <= thresh) Or (op = ">=" And v(i) >= thresh) Then
            If startIdx = 0 Then startIdx = i
            acc = acc + dt
            If acc >= holdMin Then Find_FirstHold_Single_Range = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0
        End If
    Next i
End Function

Private Function Find_CFT_DeltaRiseHold_FromBaseline_Range( _
    ByRef v() As Double, ByRef t() As Double, _
    ByVal startIdx As Long, ByVal deltaReq As Double, _
    ByVal holdMin As Double, ByVal i1 As Long) As Long

    If startIdx <= 0 Then Exit Function
    Dim base As Double: base = v(startIdx)
    Dim i As Long, dt As Double, acc As Double, firstHit As Long: firstHit = 0

    For i = Application.Max(startIdx + 1, 2) To i1
        dt = MinutesBetween(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If v(i) >= base + deltaReq Then
            If firstHit = 0 Then firstHit = i
            acc = acc + dt
            If acc >= holdMin Then Find_CFT_DeltaRiseHold_FromBaseline_Range = firstHit: Exit Function
        Else
            firstHit = 0: acc = 0
        End If
    Next i
End Function


