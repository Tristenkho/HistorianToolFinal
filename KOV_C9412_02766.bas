Attribute VB_Name = "KOV_C9412_02766"
Option Explicit

'======================== SHEETS ========================
Private Const SH_DATA   As String = "Paste Data"
Private Const SH_LIMITS As String = "Product Limits"
Private Const SH_TAGMAP As String = "Tag Map"
Private Const SH_KOV    As String = "KOV"

'======================== ROLES ========================
Private Const ROLE_TT  As String = "TT"     ' Reactor temperature
Private Const ROLE_PT  As String = "PT"     ' Reactor pressure (psia)
Private Const ROLE_PFT As String = "PFT"    ' PAM flow transmitter (e.g., R3_FT_11)
Private Const ROLE_CFT As String = "CFT"    ' Cooling flow transmitter (e.g., R3_E3_FT_01)

'======================== DETECTION THRESHOLDS =========
Private Const PFT_START_THRESH     As Double = 30   ' >30 for 10 min
Private Const PFT_START_HOLD_MIN   As Double = 10
Private Const PFT_END_THRESH       As Double = 30   ' <30 for 60 min
Private Const PFT_END_HOLD_MIN     As Double = 60

Private Const PT_STRIP_THRESH      As Double = 12   ' <=12 psia for 10 min
Private Const PT_STRIP_HOLD_MIN    As Double = 10

Private Const CFT_STRIP_THRESH     As Double = 150  ' >=150 for 5 min
Private Const CFT_STRIP_HOLD_MIN   As Double = 5

'======================== SHEET NAMES ==================
' (Assumes these globals already exist; if not, replace with literals)
' SH_DATA   = "Paste Data"
' SH_LIMITS = "Product Limits"
' SH_TAGMAP = "Tag Map"
' SH_KOV    = "KOV"

'=======================================================
'               PUBLIC ENTRIES (DISPATCH TARGETS)
'=======================================================
Public Sub KOV_Run_InfineumC9412_Main()
    KOV_Run_PAM_Generic "Infineum C9412"
End Sub

Public Sub KOV_Run_Lubrizol02766_Main()
    KOV_Run_PAM_Generic "Lubrizol 0276.6"
End Sub

'=======================================================
'                   GENERIC RUNNER
'=======================================================
Private Sub KOV_Run_PAM_Generic(ByVal PRODUCT_NAME As String)
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
    Dim TT() As Double, pt() As Double, PFT() As Double, cFT() As Double
    Dim nTT As Long, nPT As Long, nPFT As Long, nCFT As Long
    Dim dTT As Double, dPT As Double, dPFT As Double, dCFT As Double
    Dim vTT As Double, vPT As Double, vPFT As Double, vCFT As Double

    TT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_TT, n, nTT, dTT, vTT)
    pt = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_PT, n, nPT, dPT, vPT)
    PFT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_PFT, n, nPFT, dPFT, vPFT)
    cFT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_CFT, n, nCFT, dCFT, vCFT)

    If Not SeriesExists(TT) Or Not SeriesExists(pt) Or Not SeriesExists(PFT) Or Not SeriesExists(cFT) Then
        MsgBox "Required roles missing (TT/PT/PFT/CFT) for '" & PRODUCT_NAME & "'.", vbCritical
        Exit Sub
    End If

    ' ---- output header ----
    wsK.Cells.ClearContents
    wsK.Range("A1:F1").value = Array("Product", "Role", "Tags used", "N", "Max", "StdDev")
    wsK.Range("A2").value = PRODUCT_NAME
    Dim rr As Long: rr = 2
    rr = PrintRoleSummary(wsK, rr, ROLE_TT, roleTags, nTT, dTT, vTT, PRODUCT_NAME)
    rr = PrintRoleSummary(wsK, rr, ROLE_PT, roleTags, nPT, dPT, vPT)
    rr = PrintRoleSummary(wsK, rr, ROLE_PFT, roleTags, nPFT, dPFT, vPFT)
    rr = PrintRoleSummary(wsK, rr, ROLE_CFT, roleTags, nCFT, dCFT, vCFT)

    rr = rr + 2
    wsK.rows(rr - 1).RowHeight = 8
    wsK.Range("A" & rr & ":L" & rr).value = Array("Stage", "Start Time", "End Time", _
                                                  "Metric", "Value", "Min", "TV", "Max", _
                                                  "Result", "# from TV", "Label", "Notes")
    wsK.Range("A" & rr & ":L" & rr).Font.Bold = True
    rr = rr + 1

    '==================== DETECTION ====================

    ' ---- PAM Charge window via PFT holds ----
    Dim iC_Start As Long, iC_End As Long
    iC_Start = Find_Hold_Above(PFT, t, PFT_START_THRESH, PFT_START_HOLD_MIN, i0, i1)
    If iC_Start > 0 Then
        iC_End = Find_Hold_Below(PFT, t, PFT_END_THRESH, PFT_END_HOLD_MIN, iC_Start + 1, i1)
    End If

    If iC_Start > 0 And iC_End > iC_Start Then
        Dim chargeH As Double, startTemp As Double, endTemp As Double
        chargeH = HoursBetween(t(iC_Start), t(iC_End))
        startTemp = TT(iC_Start)
        endTemp = TT(iC_End)

        ' PAM Charge: Temperature (start)
        If HasLimit(wsL, PRODUCT_NAME, "PAM Charge", "Temperature (start)") Then
            rr = WriteRow(wsK, rr, "PAM Charge", t(iC_Start), t(iC_End), _
                          "Temperature (start) (F)", Round(startTemp, 1), _
                          wsL, PRODUCT_NAME, "PAM Charge", "Temperature (start)", False, _
                          "Start when PFT>30 for 10m; TT at start index.")
        End If

        ' PAM Charge: Time (h)
        If HasLimit(wsL, PRODUCT_NAME, "PAM Charge", "Time") Then
            rr = WriteRow(wsK, rr, "PAM Charge", t(iC_Start), t(iC_End), _
                          "Time (h)", Round(chargeH, 2), _
                          wsL, PRODUCT_NAME, "PAM Charge", "Time", True, _
                          "Duration between PFT>30(10m) and PFT<30(60m).")
        End If

        ' PAM Charge: Temperature (end)
        If HasLimit(wsL, PRODUCT_NAME, "PAM Charge", "Temperature (end)") Then
            rr = WriteRow(wsK, rr, "PAM Charge", t(iC_Start), t(iC_End), _
                          "Temperature (end) (F)", Round(endTemp, 1), _
                          wsL, PRODUCT_NAME, "PAM Charge", "Temperature (end)", False, _
                          "End when PFT<30 for 60m; TT at end index.")
        End If

        ' ---- Strip: Start PT<=12 (10m), End CFT>=150 (5m) ----
        Dim iS_Start As Long, iS_End As Long
        iS_Start = Find_FirstHold_Single_Range(pt, t, "<=", PT_STRIP_THRESH, PT_STRIP_HOLD_MIN, iC_End + 1, i1)
        If iS_Start > 0 Then
            iS_End = Find_FirstHold_Single_Range(cFT, t, ">=", CFT_STRIP_THRESH, CFT_STRIP_HOLD_MIN, iS_Start + 1, i1)
            If iS_End = 0 And G_KOV_UseWindow Then iS_End = i1 ' graceful fallback
        End If

        If iS_Start > 0 And iS_End > iS_Start Then
            Dim stripH As Double, stripT As Double
            stripH = HoursBetween(t(iS_Start), t(iS_End))
            stripT = TrimmedMeanTW(TT, t, iS_Start, iS_End, 10, 10)

            ' Strip: Temperature (°F)
            If HasLimit(wsL, PRODUCT_NAME, "Strip", "Temperature") Then
                rr = WriteRow(wsK, rr, "Strip", t(iS_Start), t(iS_End), _
                              "Temperature (F)", Round(stripT, 1), _
                              wsL, PRODUCT_NAME, "Strip", "Temperature", False, _
                              "Start=PT<=12(10m); End=CFT>=150(5m); TT TW-mean (trim 10/10).")
            End If

            ' Strip: Time (h) -- only present for some products (e.g., C9412)
            If HasLimit(wsL, PRODUCT_NAME, "Strip", "Time") Then
                rr = WriteRow(wsK, rr, "Strip", t(iS_Start), t(iS_End), _
                              "Time (h)", Round(stripH, 2), _
                              wsL, PRODUCT_NAME, "Strip", "Time", True, _
                              "Between PT hold and CFT hold.")
            End If

        Else
            MsgBox "Strip window not found (PT<=12(10m) then CFT>=150(5m)).", vbExclamation
        End If

    Else
        MsgBox "PAM Charge window not found (PFT>30(10m) then PFT<30(60m)).", vbExclamation
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
    Dim lastRow As Long: lastRow = ws.Cells(ws.rows.Count, cTime).End(xlUp).Row
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
    Set d(ROLE_PFT) = New Collection
    Set d(ROLE_CFT) = New Collection

    Dim lastRow As Long: lastRow = wsMap.Cells(wsMap.rows.Count, 1).End(xlUp).Row
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
'                DETECTION HELPERS
'=======================================================
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

Private Function Find_FirstHold_Single_Range( _
    ByRef v() As Double, ByRef t() As Double, _
    ByVal op As String, ByVal thresh As Double, _
    ByVal holdMin As Double, ByVal fromIdx As Long, ByVal i1 As Long) As Long

    Dim i As Long, acc As Double, dt As Double, startIdx As Long
    startIdx = 0
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


