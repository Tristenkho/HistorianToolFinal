Attribute VB_Name = "KOV_OLI"
Option Explicit

'======================== SHEETS ========================
Private Const SH_DATA   As String = "Paste Data"
Private Const SH_LIMITS As String = "Product Limits"
Private Const SH_TAGMAP As String = "Tag Map"
Private Const SH_KOV    As String = "KOV"

'======================== ROLES =========================
Private Const ROLE_TT As String = "TT"   ' R4_TT_02
Private Const ROLE_PT As String = "PT"   ' R4_PT_01
Private Const ROLE_FT As String = "FT"   ' R4_FT_01

'======================== THRESHOLDS ====================
' batch/est/strip detection
Private Const FT_START As Double = 500
Private Const PT_ATM   As Double = 12
Private Const HOLD10   As Double = 10
Private Const EST_TT_MIN As Double = 356  ' esterification est. floor

' trims
Private Const EST_TRIM_IN  As Double = 10
Private Const EST_TRIM_OUT As Double = 10
Private Const STRIP_TRIM   As Double = 30  ' for TT min/max in strip

'=======================================================
'                  PUBLIC ENTRYPOINTS
'=======================================================
Public Sub KOV_Run_InnospecOLI9000M_Main()
    Run_R4_OLI "Innospec OLI 9000M", True
End Sub

Public Sub KOV_Run_InnospecOLI9200LN_Main()
    Run_R4_OLI "Innospec OLI 9200LN", True
End Sub

'=======================================================
'            MAIN ENGINE (shared for both OLI)
'=======================================================
Private Sub Run_R4_OLI(ByVal PRODUCT_NAME As String, ByVal writeStripTTminmax As Boolean)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsD As Worksheet: Set wsD = wb.Worksheets(SH_DATA)
    Dim wsL As Worksheet: Set wsL = wb.Worksheets(SH_LIMITS)
    Dim wsM As Worksheet: Set wsM = wb.Worksheets(SH_TAGMAP)
    Dim wsK As Worksheet
    On Error Resume Next
    Set wsK = wb.Worksheets(SH_KOV)
    On Error GoTo 0
    If wsK Is Nothing Then Set wsK = wb.Worksheets.Add(After:=wsD): wsK.name = SH_KOV

    ' ---- headers / time ----
    Dim hdr As Object: Set hdr = BuildHeaderIndexAll(wsD)
    Dim cT As Long: cT = HeaderCol(hdr, "Time")
    If cT = 0 Then MsgBox "Missing 'Time' header in Paste Data.", vbCritical: Exit Sub
    Dim t() As Double, n As Long
    If Not BuildTimeVector(wsD, cT, t, n) Then MsgBox "Time column not recognized.", vbCritical: Exit Sub

    ' ---- window (from WeekRunner, if set) ----
    Dim i0 As Long, i1 As Long: ResolveWindowBoundsLocal t, i0, i1

    ' ---- roles -> tags for THIS product ----
    Dim roleTags As Object: Set roleTags = GroupTagsByRole_Explicit(wsM, PRODUCT_NAME, hdr)

    ' ---- composites ----
    Dim TT() As Double, pt() As Double, ft() As Double
    Dim nTT As Long, nPT As Long, nFT As Long
    Dim dTT As Double, dPT As Double, dFT As Double
    Dim vTT As Double, vPT As Double, vFT As Double

    TT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_TT, n, nTT, dTT, vTT)
    pt = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_PT, n, nPT, dPT, vPT)
    ft = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_FT, n, nFT, dFT, vFT)

    If Not SeriesExists(TT) Or Not SeriesExists(pt) Or Not SeriesExists(ft) Then
        MsgBox "Required roles missing (TT/PT/FT) for '" & PRODUCT_NAME & "'.", vbCritical
        Exit Sub
    End If

    ' ---- output header ----
    wsK.Cells.ClearContents
    wsK.Range("A1:F1").value = Array("Product", "Role", "Tags used", "Redundancy (N)", "Redundancy (Max)", "Redundancy (StdDev)")
    wsK.Range("A2").value = PRODUCT_NAME
    Dim rr As Long: rr = 2
    rr = PrintRoleSummary(wsK, rr, ROLE_TT, roleTags, nTT, dTT, vTT, PRODUCT_NAME)
    rr = PrintRoleSummary(wsK, rr, ROLE_PT, roleTags, nPT, dPT, vPT)
    rr = PrintRoleSummary(wsK, rr, ROLE_FT, roleTags, nFT, dFT, vFT)

    rr = rr + 2
    wsK.Rows(rr - 1).RowHeight = 8
    wsK.Range("A" & rr & ":L" & rr).value = Array("Stage", "Start Time", "End Time", "Metric", "Value", "Min", "TV", "Max", "Result", "# from TV", "Label", "Notes")
    wsK.Range("A" & rr & ":L" & rr).Font.Bold = True
    rr = rr + 1

    '==================== DETECTION ====================

    ' ---- Batch start: FT>500 AND PT>12 for 10m ----
    Dim iB_Start As Long
    iB_Start = Find_Hold_BothAbove(ft, pt, t, FT_START, PT_ATM, HOLD10, i0, i1)

    ' ---- Ester start: TT>356 for 10m (after batch start) ----
    Dim iE_Start As Long
    If iB_Start > 0 Then iE_Start = Find_Hold_Above(TT, t, EST_TT_MIN, HOLD10, iB_Start + 1, i1)

    ' ---- Ester end / Strip start: PT<12 for 10m (after ester start) ----
    Dim iS_Start As Long
    If iE_Start > 0 Then iS_Start = Find_Hold_Below(pt, t, PT_ATM, HOLD10, iE_Start + 1, i1)

    ' ---- Strip end / Batch end: PT>12 for 10m ----
    Dim iB_End As Long
    If iS_Start > 0 Then iB_End = Find_Hold_Above(pt, t, PT_ATM, HOLD10, iS_Start + 1, i1)
    If iB_End = 0 And G_KOV_UseWindow Then iB_End = i1

    '==================== METRICS ======================

    ' ---- Esterification temp (TW-mean with 10/10 trims) ----
    If iB_Start > 0 And iE_Start > 0 And iS_Start > iE_Start Then
        Dim estT As Double: estT = TrimmedMeanTW(TT, t, iE_Start, iS_Start, EST_TRIM_IN, EST_TRIM_OUT)
        rr = WriteRowOrNoLimit(wsK, rr, "Esterification", t(iE_Start), t(iS_Start), _
                               "Temperature (F)", Round(estT, 1), _
                               wsL, PRODUCT_NAME, "Esterification", "Temperature", False, _
                               "Start: TT>356(10m) after batch start; End: PT<12(10m). TT TW-mean (trim 10/10).")
    Else
        rr = WriteNoLimitRow(wsK, rr, "Esterification", "", "", _
                             "Temperature (F)", "", _
                             "Window not found (check TT>356 hold or PT<12 hold).")
    End If

    ' ---- Strip stage: PT(min), TT min/max with 30-min trims (product-specific) ----
    If iS_Start > 0 And iB_End > iS_Start Then
        ' PT min over the entire strip stage
        Dim ptMin As Double: ptMin = SeriesMinInRange(pt, iS_Start, iB_End)
        ' PT(min) AOV row (both products have this in limits)
        rr = WriteRowOrNoLimit(wsK, rr, "Strip", t(iS_Start), t(iB_End), _
                               "Pressure (min) (psia)", Round(ptMin, 2), _
                               wsL, PRODUCT_NAME, "Strip", "Pressure (min)", False, _
                               "Min PT during strip (PT<12?PT>12).")

        If writeStripTTminmax Then
            ' For 9000M: TT min/max with 30-min trims on strip window
            Dim iT0 As Long, iT1 As Long
            TrimWindowByMinutes t, iS_Start, iB_End, STRIP_TRIM, STRIP_TRIM, iT0, iT1
            If iT0 > 0 And iT1 > iT0 Then
                Dim ttMin As Double, ttMax As Double
                ttMin = SeriesMinInRange(TT, iT0, iT1)
                ttMax = SeriesMaxInRange(TT, iT0, iT1)

                rr = WriteRowOrNoLimit(wsK, rr, "Strip", t(iS_Start), t(iB_End), _
                                       "Temperature (min) (F)", Round(ttMin, 1), _
                                       wsL, PRODUCT_NAME, "Strip", "Temperature (min)", False, _
                                       "TT min in strip with ±30 min trims.")
                rr = WriteRowOrNoLimit(wsK, rr, "Strip", t(iS_Start), t(iB_End), _
                                       "Temperature (max) (F)", Round(ttMax, 1), _
                                       wsL, PRODUCT_NAME, "Strip", "Temperature (max)", False, _
                                       "TT max in strip with ±30 min trims.")
            Else
                rr = WriteNoLimitRow(wsK, rr, "Strip", "", "", _
                                     "Temperature (min/max) (F)", "", _
                                     "Trim left too short for 30/30 min; widen window.")
            End If
        End If
    Else
        rr = WriteNoLimitRow(wsK, rr, "Strip", "", "", _
                             "Pressure (min) (psia)", "", _
                             "Strip window not found (PT<12(10m) then PT>12(10m)).")
    End If

    KOV_ColorizeAllTables wsK
    wsK.Columns("A:L").AutoFit
    KOV_Notify "KOV complete for '" & PRODUCT_NAME & "'."
End Sub

'=======================================================
'                     LOCAL HELPERS
'   (kept here so this module works standalone)
'   Uses your shared: CompositeMedian_AndStats, SeriesExists,
'   PrintRoleSummary, HasLimit, WriteRow, TrimmedMeanTW,
'   KOV_ColorizeAllTables.
'=======================================================

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

Private Sub ResolveWindowBoundsLocal(ByRef t() As Double, ByRef i0 As Long, ByRef i1 As Long)
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
        If i1 < i0 Then i0 = 1: i1 = n
    Else
        i0 = 1: i1 = UBound(t)
    End If
End Sub

' both-above hold (FT>FT_START AND PT>PT_ATM for holdMin)
Private Function Find_Hold_BothAbove(ByRef ft() As Double, ByRef pt() As Double, ByRef t() As Double, _
    ByVal ftThresh As Double, ByVal ptThresh As Double, ByVal holdMin As Double, _
    ByVal i0 As Long, ByVal i1 As Long) As Long

    Dim i As Long, acc As Double, dt As Double, startIdx As Long
    startIdx = 0: acc = 0
    For i = Application.Max(i0, 2) To i1
        dt = MinutesBetween(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If (ft(i) > ftThresh) And (pt(i) > ptThresh) Then
            If startIdx = 0 Then startIdx = i
            acc = acc + dt
            If acc >= holdMin Then Find_Hold_BothAbove = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0
        End If
    Next i
End Function

Private Function Find_Hold_Above(ByRef v() As Double, ByRef t() As Double, _
    ByVal thresh As Double, ByVal holdMin As Double, ByVal i0 As Long, ByVal i1 As Long) As Long

    Dim i As Long, acc As Double, dt As Double, startIdx As Long
    startIdx = 0: acc = 0
    For i = Application.Max(i0, 2) To i1
        dt = MinutesBetween(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If v(i) > thresh Then
            If startIdx = 0 Then startIdx = i
            acc = acc + dt
            If acc >= holdMin Then Find_Hold_Above = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0
        End If
    Next i
End Function

Private Function Find_Hold_Below(ByRef v() As Double, ByRef t() As Double, _
    ByVal thresh As Double, ByVal holdMin As Double, ByVal i0 As Long, ByVal i1 As Long) As Long

    Dim i As Long, acc As Double, dt As Double, startIdx As Long
    startIdx = 0: acc = 0
    For i = Application.Max(i0, 2) To i1
        dt = MinutesBetween(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If v(i) < thresh Then
            If startIdx = 0 Then startIdx = i
            acc = acc + dt
            If acc >= holdMin Then Find_Hold_Below = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0
        End If
    Next i
End Function

Private Function SeriesMinInRange(ByRef v() As Double, ByVal iStart As Long, ByVal iEnd As Long) As Double
    If iStart <= 0 Or iEnd <= 0 Or iEnd < iStart Then Exit Function
    Dim i As Long, m As Double: m = v(iStart)
    For i = iStart + 1 To iEnd
        If v(i) < m Then m = v(i)
    Next i
    SeriesMinInRange = m
End Function

Private Function SeriesMaxInRange(ByRef v() As Double, ByVal iStart As Long, ByVal iEnd As Long) As Double
    If iStart <= 0 Or iEnd <= 0 Or iEnd < iStart Then Exit Function
    Dim i As Long, m As Double: m = v(iStart)
    For i = iStart + 1 To iEnd
        If v(i) > m Then m = v(i)
    Next i
    SeriesMaxInRange = m
End Function

Private Sub TrimWindowByMinutes(ByRef t() As Double, ByVal iStart As Long, ByVal iEnd As Long, _
    ByVal trimIn As Double, ByVal trimOut As Double, ByRef iT0 As Long, ByRef iT1 As Long)

    ' move start forward by trimIn minutes; end backward by trimOut minutes
    Dim targetStart As Double: targetStart = t(iStart) + trimIn / (24# * 60#)
    Dim targetEnd As Double:   targetEnd = t(iEnd) - trimOut / (24# * 60#)

    Dim i As Long
    iT0 = iStart
    For i = iStart To iEnd
        If t(i) >= targetStart Then iT0 = i: Exit For
    Next i

    iT1 = iEnd
    For i = iEnd To iStart Step -1
        If t(i) <= targetEnd Then iT1 = i: Exit For
    Next i

    If iT1 <= iT0 Then iT0 = 0: iT1 = 0
End Sub

' ---- "Always show" wrappers ----
Private Function WriteNoLimitRow(ws As Worksheet, ByVal rr As Long, _
    ByVal stage As String, ByVal tStart As Variant, ByVal tEnd As Variant, _
    ByVal metric As String, ByVal value As Variant, ByVal notes As String) As Long

    ws.Cells(rr, 1).value = stage
    If Not IsEmpty(tStart) Then ws.Cells(rr, 2).value = tStart
    If Not IsEmpty(tEnd) Then ws.Cells(rr, 3).value = tEnd
    ws.Cells(rr, 4).value = metric
    If Not IsEmpty(value) Then ws.Cells(rr, 5).value = value
    ws.Cells(rr, 9).value = "No limit"
    ws.Cells(rr, 11).value = "Info"
    ws.Cells(rr, 12).value = notes
    WriteNoLimitRow = rr + 1
End Function

Private Function WriteRowOrNoLimit(ws As Worksheet, ByVal rr As Long, _
    ByVal stage As String, ByVal tStart As Variant, ByVal tEnd As Variant, _
    ByVal metric As String, ByVal value As Variant, _
    wsL As Worksheet, ByVal product As String, ByVal section As String, ByVal varLabel As String, _
    ByVal isTime As Boolean, ByVal notes As String) As Long

    If IsEmpty(value) Then
        WriteRowOrNoLimit = WriteNoLimitRow(ws, rr, stage, tStart, tEnd, metric, "", notes & " [no data]")
    ElseIf HasLimit(wsL, product, section, varLabel) Then
        WriteRowOrNoLimit = WriteRow(ws, rr, stage, tStart, tEnd, metric, value, _
                                     wsL, product, section, varLabel, isTime, notes)
    Else
        WriteRowOrNoLimit = WriteNoLimitRow(ws, rr, stage, tStart, tEnd, metric, value, notes & " [no limit]")
    End If
End Function

