Attribute VB_Name = "KOV_19858"
Option Explicit

'======================== SHEETS ========================
Private Const SH_DATA   As String = "Paste Data"
Private Const SH_LIMITS As String = "Product Limits"
Private Const SH_TAGMAP As String = "Tag Map"
Private Const SH_KOV    As String = "KOV"

'======================== PRODUCT =======================
Private Const PRODUCT_NAME As String = "Lubrizol 198.58"

'======================== ROLES (Tag Map roles) =========
'   TT  -> Reactor temperature (e.g., R1_TT_01 / R1_TT_02)
'   MFT -> Maleic charge flow      (e.g., T234_FT_01)
'   MTT -> Maleic tank temperature (e.g., T234_TT_01/02)
'   CFT -> Cooler/cooling flow     (e.g., R1_E3_FT_01)
Private Const ROLE_TT  As String = "TT"
Private Const ROLE_MFT As String = "MFT"
Private Const ROLE_MTT As String = "MTT"
Private Const ROLE_CFT As String = "CFT"

'======================== THRESHOLDS (EDIT HERE) =========
' ---- Maleic Charge window (based on MFT) ----
Private Const MFT_BAND_LO      As Double = 95
Private Const MFT_BAND_HI      As Double = 135
Private Const MFT_INBAND_HOLD  As Double = 10
Private Const MFT_OUTBAND_HOLD As Double = 10

' ---- Trimming (minutes removed from ends for temperature means) ----
Private Const TRIM_IN_MIN      As Double = 10
Private Const TRIM_OUT_MIN     As Double = 10

' ---- Soak end rule (based on CFT rise from soak start) ----
Private Const SOAK_CFT_DELTA   As Double = 150
Private Const SOAK_HOLD_MIN    As Double = 10

'=======================================================
'                    PUBLIC ENTRY
'=======================================================
Public Sub KOV_Run_Lubrizol19858_Main()
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

    ' ---- Build header index & time vector ----
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
    
    ' ---- limit to the Batch Summary window (if set) ----
    Dim i0 As Long, i1 As Long
    ResolveWindowBounds t, i0, i1

    ' ---- Group tags by explicit roles for THIS product ----
    Dim roleTags As Object
    Set roleTags = GroupTagsByRole_Explicit(wsM, PRODUCT_NAME, hdr)

    ' ---- Build composites (median across redundant tags) ----
    Dim TT() As Double, MFT() As Double, MTT() As Double, cFT() As Double
    Dim nTT As Long, nMFT As Long, nMTT As Long, nCFT As Long
    Dim dTT As Double, dMFT As Double, dMTT As Double, dCFT As Double
    Dim vTT As Double, vMFT As Double, vMTT As Double, vCFT As Double

    TT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_TT, n, nTT, dTT, vTT)
    MFT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_MFT, n, nMFT, dMFT, vMFT)
    MTT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_MTT, n, nMTT, dMTT, vMTT)
    cFT = CompositeMedian_AndStats(wsD, t, hdr, roleTags, ROLE_CFT, n, nCFT, dCFT, vCFT)

    If Not SeriesExists(TT) Or Not SeriesExists(MFT) Or Not SeriesExists(MTT) Then
        MsgBox "Required roles missing (TT/MFT/MTT). Check Tag Map and Paste Data headers ('.Val' accepted).", vbCritical
        Exit Sub
    End If

    ' ---- Clear KOV & redundancy block ----
    wsK.Cells.ClearContents
    wsK.Range("A1:F1").value = Array("Product", "Role", "Tags used", "Redundancy (N)", "Redundancy (Max)", "Redundancy (StdDev)")
    wsK.Range("A2").value = PRODUCT_NAME
    Dim rr As Long: rr = 2
    rr = PrintRoleSummary(wsK, rr, ROLE_TT, roleTags, nTT, dTT, vTT, PRODUCT_NAME)
    rr = PrintRoleSummary(wsK, rr, ROLE_MFT, roleTags, nMFT, dMFT, vMFT)
    rr = PrintRoleSummary(wsK, rr, ROLE_MTT, roleTags, nMTT, dMTT, vMTT)
    rr = PrintRoleSummary(wsK, rr, ROLE_CFT, roleTags, nCFT, dCFT, vCFT)

    rr = rr + 2
    wsK.Rows(rr - 1).RowHeight = 8
    wsK.Range("A" & rr & ":L" & rr).value = Array("Stage", "Start Time", "End Time", "Metric", "Value", "Min", "TV", "Max", "Result", "# from TV", "Label", "Notes")
    wsK.Range("A" & rr & ":L" & rr).Font.Bold = True
    rr = rr + 1

' -------------------- MALEIC CHARGE --------------------
If SeriesExists(MFT) Then
    Dim iM_Start As Long, iM_End As Long

    ' find first in-band hold, then first sustained out-of-band hold
    iM_Start = FirstHold_InBand_Range(MFT, MFT_BAND_LO, MFT_BAND_HI, t, MFT_INBAND_HOLD, i0, i1)
    If iM_Start > 0 Then
        iM_End = FirstHold_OutOfBand_Range(MFT, MFT_BAND_LO, MFT_BAND_HI, t, iM_Start + 1, MFT_OUTBAND_HOLD, i1)
    End If

    If iM_Start > 0 And iM_End > iM_Start Then
        ' reactor TT mean during charge (trimmed)
        Dim ttCharge As Double
        ttCharge = TrimmedMeanTW(TT, t, iM_Start, iM_End, TRIM_IN_MIN, TRIM_OUT_MIN)

        ' MTT mean during charge (no trim)
        Dim mttCharge As Double
        mttCharge = TimeWeightedMeanWindow(MTT, t, iM_Start, iM_End, 0#)

        ' MFT mean during charge (no trim) – "Rate"
        Dim rateCharge As Double
        rateCharge = TimeWeightedMeanWindow(MFT, t, iM_Start, iM_End, 0#)

        ' Output rows (guarded by limits existence)
        If HasLimit(wsL, PRODUCT_NAME, "Maleic Charge", "Temperature") Then
            rr = WriteRow(wsK, rr, "Maleic Charge", t(iM_Start), t(iM_End), _
                          "Reactor Temperature (F)", Round(ttCharge, 1), _
                          wsL, PRODUCT_NAME, "Maleic Charge", "Temperature", False, _
                          "TT mean (trim " & TRIM_IN_MIN & "/" & TRIM_OUT_MIN & "m); window = MFT in [" & MFT_BAND_LO & "–" & MFT_BAND_HI & _
                          "] = " & MFT_INBAND_HOLD & "m; end when out-of-band = " & MFT_OUTBAND_HOLD & "m.")
        End If

        If HasLimit(wsL, PRODUCT_NAME, "Maleic Charge", "Charge Temperature") Then
            rr = WriteRow(wsK, rr, "Maleic Charge", t(iM_Start), t(iM_End), _
                          "Charge Temperature (F)", Round(mttCharge, 1), _
                          wsL, PRODUCT_NAME, "Maleic Charge", "Charge Temperature", False, _
                          "MTT mean over Maleic window.")
        End If

        If HasLimit(wsL, PRODUCT_NAME, "Maleic Charge", "Rate") Then
            rr = WriteRow(wsK, rr, "Maleic Charge", t(iM_Start), t(iM_End), _
                          "Rate (lb/min)", Round(rateCharge, 1), _
                          wsL, PRODUCT_NAME, "Maleic Charge", "Rate", False, _
                          "MFT mean over Maleic window.")
        End If

        '=======================================================
        '                         SOAK
        '  Start = Maleic end; End = CFT rises by +? and holds
        '=======================================================
        Dim iSoak_Start As Long, iSoak_End As Long
        iSoak_Start = iM_End

        If SeriesExists(cFT) Then
            iSoak_End = FirstHold_CFT_DeltaRise_Range( _
                            cFT, t, iSoak_Start, SOAK_CFT_DELTA, SOAK_HOLD_MIN, i1)
        End If

        If iSoak_Start > 0 And iSoak_End > iSoak_Start Then
            Dim soakH As Double, soakT As Double
            soakH = HoursBetween(t(iSoak_Start), t(iSoak_End))
            soakT = TrimmedMeanTW(TT, t, iSoak_Start, iSoak_End, TRIM_IN_MIN, TRIM_OUT_MIN)

            If HasLimit(wsL, PRODUCT_NAME, "Soak", "Temperature") Then
                rr = WriteRow(wsK, rr, "Soak", t(iSoak_Start), t(iSoak_End), _
                              "Temperature (F)", Round(soakT, 1), _
                              wsL, PRODUCT_NAME, "Soak", "Temperature", False, _
                              "Start=Maleic end; End when CFT = base+" & SOAK_CFT_DELTA & _
                              " for = " & SOAK_HOLD_MIN & "m; trim " & TRIM_IN_MIN & "/" & TRIM_OUT_MIN & "m.")
            End If

            If HasLimit(wsL, PRODUCT_NAME, "Soak", "Time") Then
                rr = WriteRow(wsK, rr, "Soak", t(iSoak_Start), t(iSoak_End), _
                              "Time (h)", Round(soakH, 2), _
                              wsL, PRODUCT_NAME, "Soak", "Time", True, _
                              "Hours from Maleic end to CFT = base+" & SOAK_CFT_DELTA & " held " & SOAK_HOLD_MIN & "m.")
            End If
        Else
            MsgBox "Soak END not found (need CFT rise +" & SOAK_CFT_DELTA & _
                   " for " & SOAK_HOLD_MIN & " min).", vbExclamation
        End If

    Else
        MsgBox "Maleic window not found (MFT " & MFT_BAND_LO & "–" & MFT_BAND_HI & _
               " for " & MFT_INBAND_HOLD & " min).", vbExclamation
    End If

Else
    MsgBox "Maleic flow role (MFT) missing. Check Tag Map / Paste Data.", vbExclamation
End If

KOV_ColorizeAllTables wsK
wsK.Columns("A:L").AutoFit
KOV_Notify "KOV complete for '" & PRODUCT_NAME & "'."
End Sub

'=======================================================
'                      HELPERS (product-specific only)
'=======================================================

'---- Tag Map (explicit roles, accepts tag or tag+".Val") ----
Private Function GroupTagsByRole_Explicit(wsMap As Worksheet, ByVal product As String, hdr As Object) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set d(ROLE_TT) = New Collection
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

'---- Headers / time vector ----
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


