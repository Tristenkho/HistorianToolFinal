Attribute VB_Name = "KOV_C9283_11658"
Option Explicit

'======================== SHEETS ========================
Private Const SH_DATA   As String = "Paste Data"
Private Const SH_LIMITS As String = "Product Limits"
Private Const SH_TAGMAP As String = "Tag Map"
Private Const SH_KOV    As String = "KOV"

'======================== ROLES =========================
Private Const ROLE_TT  As String = "TT"   ' Reactor temperature
Private Const ROLE_PFT As String = "PFT"  ' PAM flow (R3_FT_11)        - C9283
Private Const ROLE_DFT As String = "DFT"  ' DMAPA flow (R3_FT_11)      - 116.58
Private Const ROLE_AFT As String = "AFT"  ' APP SN100 flow (R3_FT_03)  - C9283
Private Const ROLE_CFT As String = "CFT"  ' Cooling flow (R3_E3_FT_01) - 116.58

'======================== THRESHOLDS ====================
' Charge window (both products)
Private Const FLOW_START_THRESH     As Double = 40   ' >40 for 10 min
Private Const FLOW_START_HOLD_MIN   As Double = 10
Private Const FLOW_END_THRESH       As Double = 30   ' <30 for 60 min
Private Const FLOW_END_HOLD_MIN     As Double = 60

' Strip ends per product
Private Const AFT_STRIP_THRESH      As Double = 150  ' C9283: AFT >=150 for 10 min
Private Const AFT_STRIP_HOLD_MIN    As Double = 10

Private Const CFT_STRIP_THRESH      As Double = 150  ' 116.58: CFT >=150 for 10 min
Private Const CFT_STRIP_HOLD_MIN    As Double = 10

' Trim for TW averages
Private Const TRIM_IN_MIN           As Double = 10
Private Const TRIM_OUT_MIN          As Double = 10

'=======================================================
'               PUBLIC ENTRYPOINTS
'=======================================================
Public Sub KOV_Run_Lubrizol11658_Main()
    Run_11658 "Lubrizol 116.58"
End Sub

Public Sub KOV_Run_InfineumC9283_Main()
    Run_C9283 "Infineum C9283"
End Sub

'=======================================================
'                 LUBRIZOL 116.58 ENGINE
' Limits expected:
'   DMAPA Charge / Temperature (start)   °F   200 / 215 / 230   (KOV)
'   DMAPA Charge / Temperature           °F   200 / 215 / 250   (KOV)   [avg]
'   Strip        / Temperature (max)     °F   295 / 300 / 305   (AOV)   [max]
'=======================================================
Private Sub Run_11658(ByVal PRODUCT_NAME As String)
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
    Dim cT As Long: cT = HeaderCol(hdr, "Time"): If cT = 0 Then MsgBox "Missing 'Time' in Paste Data.", vbCritical: Exit Sub
    Dim t() As Double, n As Long
    If Not BuildTimeVector(wsD, cT, t, n) Then MsgBox "Time column not recognized.", vbCritical: Exit Sub

    ' ---- window (from WeekRunner globals if set) ----
    Dim i0 As Long, i1 As Long: ResolveWindowBoundsLocal t, i0, i1

    ' ---- roles -> tags ----
    Dim roleTags As Object: Set roleTags = GroupTagsByRole_Explicit(wsM, PRODUCT_NAME, hdr)

    ' ---- composites ----
    Dim TT() As Double, dFT() As Double, cFT() As Double
    Dim nTT As Long, nDFT As Long, nCFT As Long
    Dim dTT As Double, dDFT As Double, dCFT As Double
    Dim vTT As Double, vDFT As Double, vCFT As Double

    TT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_TT, n, nTT, dTT, vTT)
    dFT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_DFT, n, nDFT, dDFT, vDFT)
    cFT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_CFT, n, nCFT, dCFT, vCFT)

    If Not SeriesExists(TT) Or Not SeriesExists(dFT) Or Not SeriesExists(cFT) Then
        MsgBox "Missing TT/DFT/CFT for '" & PRODUCT_NAME & "'.", vbCritical
        Exit Sub
    End If

    ' ---- output header ----
    wsK.Cells.ClearContents
    wsK.Range("A1:F1").value = Array("Product", "Role", "Tags found in Paste Data", "N", "Max?", "StdDev")
    wsK.Range("A2").value = PRODUCT_NAME
    Dim rr As Long: rr = 2
    rr = PrintRoleSummary(wsK, rr, ROLE_TT, roleTags, nTT, dTT, vTT, PRODUCT_NAME)
    rr = PrintRoleSummary(wsK, rr, ROLE_DFT, roleTags, nDFT, dDFT, vDFT)
    rr = PrintRoleSummary(wsK, rr, ROLE_CFT, roleTags, nCFT, dCFT, vCFT)

    rr = rr + 2
    wsK.Rows(rr - 1).RowHeight = 8
    wsK.Range("A" & rr & ":L" & rr).value = Array("Stage", "Start Time", "End Time", "Metric", "Value", "Min", "TV", "Max", "Result", "# from TV", "Label", "Notes")
    wsK.Range("A" & rr & ":L" & rr).Font.Bold = True
    rr = rr + 1

    '==================== DETECTION ====================

    ' ---- DMAPA Charge via DFT holds ----
    Dim iC_Start As Long, iC_End As Long
    iC_Start = Find_Hold_Above(dFT, t, FLOW_START_THRESH, FLOW_START_HOLD_MIN, i0, i1)
    If iC_Start > 0 Then iC_End = Find_Hold_Below(dFT, t, FLOW_END_THRESH, FLOW_END_HOLD_MIN, iC_Start + 1, i1)
    If iC_End = 0 And G_KOV_UseWindow Then iC_End = i1

    If iC_Start > 0 And iC_End > iC_Start Then
        Dim startT As Double: startT = TT(iC_Start)
        Dim chargeTavg As Double: chargeTavg = TrimmedMeanTW(TT, t, iC_Start, iC_End, TRIM_IN_MIN, TRIM_OUT_MIN)

        rr = WriteRowOrNoLimit(wsK, rr, "DMAPA Charge", t(iC_Start), t(iC_End), _
                               "Temperature (start) (F)", Round(startT, 1), _
                               wsL, PRODUCT_NAME, "DMAPA Charge", "Temperature (start)", False, _
                               "Charge start: DFT>40(10m); start TT at index.")

        rr = WriteRowOrNoLimit(wsK, rr, "DMAPA Charge", t(iC_Start), t(iC_End), _
                               "Temperature (F)", Round(chargeTavg, 1), _
                               wsL, PRODUCT_NAME, "DMAPA Charge", "Temperature", False, _
                               "Charge end: DFT<30(60m). Metric=TT TW-mean (trim 10/10).")

        ' ---- Strip (start = charge end; end = CFT>=150 for 10m; metric = MAX TT) ----
        Dim iS_Start As Long, iS_End As Long
        iS_Start = iC_End
        If iS_Start > 0 Then
            iS_End = Find_Hold_Above(cFT, t, CFT_STRIP_THRESH, CFT_STRIP_HOLD_MIN, iS_Start + 1, i1)
            If iS_End = 0 And G_KOV_UseWindow Then iS_End = i1
        End If

        If iS_Start > 0 And iS_End > iS_Start Then
            Dim stripTmax As Double: stripTmax = SeriesMaxInRange(TT, iS_Start, iS_End)
            rr = WriteRowOrNoLimit(wsK, rr, "Strip", t(iS_Start), t(iS_End), _
                                   "Temperature (max) (F)", Round(stripTmax, 1), _
                                   wsL, PRODUCT_NAME, "Strip", "Temperature (max)", False, _
                                   "Strip start=charge end; end=CFT>=150(10m). Metric=MAX TT in strip.")
        Else
            rr = WriteNoLimitRow(wsK, rr, "Strip", "", "", _
                                 "Temperature (max) (F)", "", _
                                 "Strip window not found (CFT>=150(10m)).")
        End If

    Else
        rr = WriteNoLimitRow(wsK, rr, "DMAPA Charge", "", "", _
                             "Temperature (F)", "", _
                             "Charge window not found (DFT>40(10m) then DFT<30(60m)).")
    End If

    KOV_ColorizeAllTables wsK
    wsK.Columns("A:L").AutoFit
    KOV_Notify "KOV complete for '" & PRODUCT_NAME & "'."
End Sub

'=======================================================
'                 INFINEUM C9283 ENGINE
' Limits expected:
'   PAM Charge / Temperature      °F   293 / 336 / 338   (AOV) [avg]
'   PAM Charge / Time             h    2 / 3 / 4         (KOV)
'   Strip     / Temperature       °F   293 / 336 / 338   (KOV) [avg]
'=======================================================
Private Sub Run_C9283(ByVal PRODUCT_NAME As String)
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
    Dim cT As Long: cT = HeaderCol(hdr, "Time"): If cT = 0 Then MsgBox "Missing 'Time' in Paste Data.", vbCritical: Exit Sub
    Dim t() As Double, n As Long
    If Not BuildTimeVector(wsD, cT, t, n) Then MsgBox "Time column not recognized.", vbCritical: Exit Sub

    ' ---- window ----
    Dim i0 As Long, i1 As Long: ResolveWindowBoundsLocal t, i0, i1

    ' ---- roles -> tags ----
    Dim roleTags As Object: Set roleTags = GroupTagsByRole_Explicit(wsM, PRODUCT_NAME, hdr)

    ' ---- composites ----
    Dim TT() As Double, PFT() As Double, AFT() As Double
    Dim nTT As Long, nPFT As Long, nAFT As Long
    Dim dTT As Double, dPFT As Double, dAFT As Double
    Dim vTT As Double, vPFT As Double, vAFT As Double

    TT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_TT, n, nTT, dTT, vTT)
    PFT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_PFT, n, nPFT, dPFT, vPFT)
    AFT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_AFT, n, nAFT, dAFT, vAFT)

    If Not SeriesExists(TT) Or Not SeriesExists(PFT) Or Not SeriesExists(AFT) Then
        MsgBox "Missing TT/PFT/AFT for '" & PRODUCT_NAME & "'.", vbCritical
        Exit Sub
    End If

    ' ---- output header ----
    wsK.Cells.ClearContents
    wsK.Range("A1:F1").value = Array("Product", "Role", "Tags found in Paste Data", "N", "Max?", "StdDev")
    wsK.Range("A2").value = PRODUCT_NAME
    Dim rr As Long: rr = 2
    rr = PrintRoleSummary(wsK, rr, ROLE_TT, roleTags, nTT, dTT, vTT, PRODUCT_NAME)
    rr = PrintRoleSummary(wsK, rr, ROLE_PFT, roleTags, nPFT, dPFT, vPFT)
    rr = PrintRoleSummary(wsK, rr, ROLE_AFT, roleTags, nAFT, dAFT, vAFT)

    rr = rr + 2
    wsK.Rows(rr - 1).RowHeight = 8
    wsK.Range("A" & rr & ":L" & rr).value = Array("Stage", "Start Time", "End Time", "Metric", "Value", "Min", "TV", "Max", "Result", "# from TV", "Label", "Notes")
    wsK.Range("A" & rr & ":L" & rr).Font.Bold = True
    rr = rr + 1

    '==================== DETECTION ====================

    ' ---- PAM Charge via PFT holds ----
    Dim iC_Start As Long, iC_End As Long
    iC_Start = Find_Hold_Above(PFT, t, FLOW_START_THRESH, FLOW_START_HOLD_MIN, i0, i1)
    If iC_Start > 0 Then iC_End = Find_Hold_Below(PFT, t, FLOW_END_THRESH, FLOW_END_HOLD_MIN, iC_Start + 1, i1)
    If iC_End = 0 And G_KOV_UseWindow Then iC_End = i1

    If iC_Start > 0 And iC_End > iC_Start Then
        Dim chargeT As Double: chargeT = TrimmedMeanTW(TT, t, iC_Start, iC_End, TRIM_IN_MIN, TRIM_OUT_MIN)
        Dim chargeH As Double: chargeH = HoursBetween(t(iC_Start), t(iC_End))

        rr = WriteRowOrNoLimit(wsK, rr, "PAM Charge", t(iC_Start), t(iC_End), _
                               "Temperature (F)", Round(chargeT, 1), _
                               wsL, PRODUCT_NAME, "PAM Charge", "Temperature", False, _
                               "Charge start: PFT>40(10m); end: PFT<30(60m). TT TW-mean (trim 10/10).")

        rr = WriteRowOrNoLimit(wsK, rr, "PAM Charge", t(iC_Start), t(iC_End), _
                               "Time (h)", Round(chargeH, 2), _
                               wsL, PRODUCT_NAME, "PAM Charge", "Time", True, _
                               "Duration of PFT charge window.")

        ' ---- Strip (start = charge end; end = AFT>=150 for 10m; metric = AVG TT) ----
        Dim iS_Start As Long, iS_End As Long
        iS_Start = iC_End
        If iS_Start > 0 Then
            iS_End = Find_Hold_Above(AFT, t, AFT_STRIP_THRESH, AFT_STRIP_HOLD_MIN, iS_Start + 1, i1)
            If iS_End = 0 And G_KOV_UseWindow Then iS_End = i1
        End If

        If iS_Start > 0 And iS_End > iS_Start Then
            Dim stripT As Double: stripT = TrimmedMeanTW(TT, t, iS_Start, iS_End, TRIM_IN_MIN, TRIM_OUT_MIN)
            rr = WriteRowOrNoLimit(wsK, rr, "Strip", t(iS_Start), t(iS_End), _
                                   "Temperature (F)", Round(stripT, 1), _
                                   wsL, PRODUCT_NAME, "Strip", "Temperature", False, _
                                   "Strip start=charge end; end=AFT>=150(10m). TT TW-mean (trim 10/10).")
        Else
            rr = WriteNoLimitRow(wsK, rr, "Strip", "", "", _
                                 "Temperature (F)", "", _
                                 "Strip window not found (AFT>=150(10m)).")
        End If

    Else
        rr = WriteNoLimitRow(wsK, rr, "PAM Charge", "", "", _
                             "Temperature (F)", "", _
                             "Charge window not found (PFT>40(10m) then PFT<30(60m)).")
    End If

    KOV_ColorizeAllTables wsK
    wsK.Columns("A:L").AutoFit
    KOV_Notify "KOV complete for '" & PRODUCT_NAME & "'."
End Sub

'=======================================================
'                      HELPERS (LOCAL)
'  These are local so this module works standalone.
'  It still relies on your shared functions:
'   CompositeMedian_AndStats, SeriesExists, PrintRoleSummary,
'   HasLimit, WriteRow, TrimmedMeanTW, KOV_ColorizeAllTables.
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

Private Sub ResolveWindowBoundsLocal(ByRef t() As Double, ByRef i0 As Long, ByRef i1 As Long)
    On Error Resume Next
    Dim useW As Boolean: useW = G_KOV_UseWindow
    Dim s As Double: s = G_KOV_WindowStart
    Dim e As Double: e = G_KOV_WindowEnd
    On Error GoTo 0

    If useW Then
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

Private Function GroupTagsByRole_Explicit(wsMap As Worksheet, ByVal product As String, hdr As Object) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    ' Only roles needed in this module
    Set d(ROLE_TT) = New Collection
    Set d(ROLE_PFT) = New Collection
    Set d(ROLE_DFT) = New Collection
    Set d(ROLE_AFT) = New Collection
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

Private Function Find_Hold_Above(ByRef v() As Double, ByRef t() As Double, _
    ByVal thresh As Double, ByVal holdMin As Double, ByVal i0 As Long, ByVal i1 As Long) As Long
    Dim i As Long, acc As Double, dt As Double, startIdx As Long
    startIdx = 0
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
    ByVal thresh As Double, ByVal holdMin As Double, ByVal fromIdx As Long, ByVal i1 As Long) As Long
    If fromIdx <= 0 Then Exit Function
    Dim i As Long, acc As Double, dt As Double, startIdx As Long
    startIdx = 0
    For i = Application.Max(fromIdx, 2) To i1
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

Private Function MinutesBetween(ByVal t0 As Double, ByVal t1 As Double) As Double
    MinutesBetween = (t1 - t0) * 24# * 60#
End Function

Private Function SeriesMaxInRange(ByRef v() As Double, ByVal iStart As Long, ByVal iEnd As Long) As Double
    Dim i As Long, m As Double
    If iStart <= 0 Or iEnd <= 0 Or iEnd < iStart Then Exit Function
    m = v(iStart)
    For i = iStart + 1 To iEnd
        If v(i) > m Then m = v(i)
    Next i
    SeriesMaxInRange = m
End Function

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

