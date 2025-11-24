Attribute VB_Name = "KOV_C9402_C9411"
Option Explicit

'======================== SHEET NAMES ========================
Private Const SH_UI     As String = "UI"
Private Const SH_DATA   As String = "Paste Data"
Private Const SH_TAGMAP As String = "Tag Map"
Private Const SH_SOAK   As String = "KOV Soak"
Private Const SH_STRIP  As String = "KOV Strip"
Private Const SH_LIMITS As String = "Product Limits"
Private Const SH_KOV    As String = "KOV"

'======================== ROLE KEYS ==========================
Private Const ROLE_TT   As String = "TT"
Private Const ROLE_PT   As String = "PT"
Private Const ROLE_FT   As String = "FT"
Private Const ROLE_E3FT As String = "E3FT"

'======================== CONFIG TYPES =======================
Private Type SoakCfg
    FT_Lo As Double
    FT_Hi As Double
    inHoldMin As Double
    FT_Cross As Double
    outHoldMin As Double
    peakWinMin As Double
    fallFrac As Double
    FallHoldMin As Double
    trimIn As Double
    trimOut As Double
End Type

Private Type StripCfg
    PT_Thresh As Double
    StartHoldMin As Double
    DeltaFromStart As Double
    EndHoldMin As Double
    trimIn As Double
    trimOut As Double
End Type

'======================== PUBLIC ENTRY =======================
Public Sub KOV_Run_v2_Main()
    KOV_Run_Impl
End Sub

'======================== MAIN IMPL ==========================
Private Sub KOV_Run_Impl()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsUI As Worksheet, wsData As Worksheet, wsMap As Worksheet
    Dim wsSoak As Worksheet, wsStrip As Worksheet, wsLim As Worksheet, wsK As Worksheet

    On Error Resume Next
    Set wsUI = wb.Worksheets(SH_UI)
    Set wsData = wb.Worksheets(SH_DATA)
    Set wsMap = wb.Worksheets(SH_TAGMAP)
    Set wsSoak = wb.Worksheets(SH_SOAK)
    Set wsStrip = wb.Worksheets(SH_STRIP)
    Set wsLim = wb.Worksheets(SH_LIMITS)
    Set wsK = wb.Worksheets(SH_KOV)
    On Error GoTo 0

If wsData Is Nothing Or wsMap Is Nothing _
   Or wsSoak Is Nothing Or wsStrip Is Nothing Or wsLim Is Nothing Then
    MsgBox "Missing sheet(s). Need: Paste Data, Tag Map, KOV Soak, KOV Strip, Product Limits.", vbCritical
    Exit Sub
End If
' wsUI is allowed to be Nothing now

    If wsK Is Nothing Then
        Set wsK = wb.Worksheets.Add(After:=wsData)
        wsK.name = SH_KOV
    End If

Dim product As String

If Len(G_SELECTED_PRODUCT) > 0 Then
    ' KOV Multi / windowed runs set this
    product = G_SELECTED_PRODUCT

ElseIf Not wsUI Is Nothing Then
    ' Optional: still support the old single-product UI flow
    product = Trim$(CStr(wsUI.Range("B1").value))

Else
    MsgBox "No product specified. Run from KOV Multi (which sets G_SELECTED_PRODUCT) " & _
           "or create a UI sheet with the product in B1.", vbExclamation
    Exit Sub
End If

If Len(product) = 0 Then
    MsgBox "No product specified. Please set G_SELECTED_PRODUCT (via KOV Multi) " & _
           "or enter a product in UI!B1.", vbExclamation
    Exit Sub
End If

    Dim oldCalc As XlCalculation: oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CLEAN_FAIL

    '---- headers ----
    Dim hdr As Object: Set hdr = BuildHeaderIndex(wsData)
    Dim cTime As Long: cTime = HeaderCol(hdr, "Time")
    If cTime = 0 Then Err.Raise 5, , "Missing 'Time' header in Paste Data."

    '---- tag map ? roles ----
    Dim roleTags As Object: Set roleTags = GroupTagsByRole(wsMap, product, hdr)

    '---- time vector ----
    Dim lastRow As Long: lastRow = wsData.Cells(wsData.rows.Count, 1).End(xlUp).Row
    If lastRow < 3 Then Err.Raise 5, , "No data in 'Paste Data'."
    Dim n As Long: n = lastRow - 1
    Dim t() As Double: ReDim t(1 To n)
    Dim i As Long
    For i = 1 To n: t(i) = wsData.Cells(i + 1, cTime).Value2: Next i

    '---- composites + redundancy QA ----
    Dim TT() As Double, pt() As Double, ft() As Double, E3FT() As Double
    Dim nTT As Long, nPT As Long, nFT As Long, nE3 As Long
    Dim dTT As Double, dPT As Double, dFT As Double, dE3 As Double
    Dim vTT As Double, vPT As Double, vFT As Double, vE3 As Double
    
    TT = CompositeMedian_AndStats(wsData, t, hdr, roleTags, ROLE_TT, n, nTT, dTT, vTT)
    pt = CompositeMedian_AndStats(wsData, t, hdr, roleTags, ROLE_PT, n, nPT, dPT, vPT)
    ft = CompositeMedian_AndStats(wsData, t, hdr, roleTags, ROLE_FT, n, nFT, dFT, vFT)
    E3FT = CompositeMedian_AndStats(wsData, t, hdr, roleTags, ROLE_E3FT, n, nE3, dE3, vE3)

    If Not SeriesExists(TT) Or Not SeriesExists(pt) Or Not SeriesExists(ft) Or Not SeriesExists(E3FT) Then
        Err.Raise 5, , "Required roles missing for '" & product & "'. Need TT/PT/FT/E3FT present in Paste Data."
    End If

    '---- load configs ----
    Dim sc As SoakCfg, rc As StripCfg
    If Not LoadSoakCfg(wsSoak, product, sc) Then Err.Raise 5, , "Product '" & product & "' not found in KOV Soak."
    If Not LoadStripCfg(wsStrip, product, rc) Then Err.Raise 5, , "Product '" & product & "' not found in KOV Strip."

    '==== DETECTION ====
    Dim SoakStartIdx As Long, SoakEndCommitIdx As Long, SoakEndIdx As Long
    Dim StripStartIdx As Long, StripEndIdx As Long

    SoakStartIdx = SoakStart_MaleicThenCross(ft, t, sc.FT_Lo, sc.FT_Hi, sc.inHoldMin, sc.FT_Cross, sc.outHoldMin)

    If SoakStartIdx > 0 Then
        SoakEndIdx = SoakEnd_FallFromPeak_Backtrack(pt, t, SoakStartIdx, sc.peakWinMin, sc.fallFrac, sc.FallHoldMin, SoakEndCommitIdx)
    End If

    If SoakEndIdx > 0 Then
        StripStartIdx = FirstHoldBelow(pt, t, SoakEndIdx + 1, rc.PT_Thresh, rc.StartHoldMin)
    End If

    If StripStartIdx > 0 Then
        StripEndIdx = StripEnd_E3FT_DeltaHold(E3FT, t, StripStartIdx, rc.DeltaFromStart, rc.EndHoldMin)
    End If

    '==== METRICS ====
    Dim soakH As Double, soakT As Double, stripH As Double, stripT As Double
    If SoakStartIdx > 0 And SoakEndIdx > SoakStartIdx Then
        soakH = HoursBetween(t(SoakStartIdx), t(SoakEndIdx))
        soakT = TrimmedMeanTW(TT, t, SoakStartIdx, SoakEndIdx, sc.trimIn, sc.trimOut)
    End If
    If StripStartIdx > 0 And StripEndIdx > StripStartIdx Then
        stripH = HoursBetween(t(StripStartIdx), t(StripEndIdx))
        stripT = TrimmedMeanTW(TT, t, StripStartIdx, StripEndIdx, rc.trimIn, rc.trimOut)
    End If

    '==== OUTPUT ====
    WriteKOV wsK, product, roleTags, _
             nTT, dTT, vTT, _
             nPT, dPT, vPT, _
             nFT, dFT, vFT, _
             nE3, dE3, vE3, _
             t, SoakStartIdx, SoakEndIdx, SoakEndCommitIdx, soakH, soakT, _
             StripStartIdx, StripEndIdx, stripH, stripT, sc, rc

CLEAN_OK:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    KOV_Notify "KOV complete for '" & product & "'."
    Exit Sub

CLEAN_FAIL:
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation
End Sub

'======================== CONFIG LOADERS =====================
Private Function LoadSoakCfg(ws As Worksheet, ByVal product As String, ByRef c As SoakCfg) As Boolean
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If StrComp(CStr(ws.Cells(r, 1).value), product, vbTextCompare) = 0 Then
            c.FT_Lo = CDbl(val(ws.Cells(r, 2).value))
            c.FT_Hi = CDbl(val(ws.Cells(r, 3).value))
            c.inHoldMin = CDbl(val(ws.Cells(r, 4).value))
            c.FT_Cross = CDbl(val(ws.Cells(r, 5).value))
            c.outHoldMin = CDbl(val(ws.Cells(r, 6).value))
            c.peakWinMin = CDbl(val(ws.Cells(r, 7).value))
            c.fallFrac = CDbl(val(ws.Cells(r, 8).value))
            c.FallHoldMin = CDbl(val(ws.Cells(r, 9).value))
            c.trimIn = CDbl(val(ws.Cells(r, 10).value))
            c.trimOut = CDbl(val(ws.Cells(r, 11).value))
            LoadSoakCfg = True
            Exit Function
        End If
    Next r
End Function

Private Function LoadStripCfg(ws As Worksheet, ByVal product As String, ByRef c As StripCfg) As Boolean
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If StrComp(CStr(ws.Cells(r, 1).value), product, vbTextCompare) = 0 Then
            c.PT_Thresh = CDbl(val(ws.Cells(r, 2).value))
            c.StartHoldMin = CDbl(val(ws.Cells(r, 3).value))
            c.DeltaFromStart = CDbl(val(ws.Cells(r, 4).value))
            c.EndHoldMin = CDbl(val(ws.Cells(r, 5).value))
            c.trimIn = CDbl(val(ws.Cells(r, 6).value))
            c.trimOut = CDbl(val(ws.Cells(r, 7).value))
            LoadStripCfg = True
            Exit Function
        End If
    Next r
End Function

'======================== TAGS & COMPOSITES ==================
Private Function BuildHeaderIndex(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 'TextCompare
    Dim c As Range
    For Each c In ws.rows(1).Cells
        If Len(c.Value2) = 0 Then Exit For
        d(CStr(c.Value2)) = c.Column
    Next c
    Set BuildHeaderIndex = d
End Function

Private Function HeaderCol(hdr As Object, key As String) As Long
    If hdr.Exists(key) Then HeaderCol = hdr(key): Exit Function
    Dim base As String
    If Right$(key, 4) = ".Val" Then
        base = Left$(key, Len(key) - 4)
        If hdr.Exists(base) Then HeaderCol = hdr(base): Exit Function
    Else
        If hdr.Exists(key & ".Val") Then HeaderCol = hdr(key & ".Val"): Exit Function
    End If
End Function

Private Function GroupTagsByRole(wsMap As Worksheet, ByVal product As String, hdr As Object) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1  ' TextCompare

    Set d(ROLE_TT) = New Collection
    Set d(ROLE_PT) = New Collection
    Set d(ROLE_FT) = New Collection
    Set d(ROLE_E3FT) = New Collection

    Dim lastRow As Long: lastRow = wsMap.Cells(wsMap.rows.Count, 1).End(xlUp).Row
    Dim r As Long, prod$, tag$, role$

    For r = 2 To lastRow
        prod = Trim$(CStr(wsMap.Cells(r, 1).value))
        If Len(prod) = 0 Then GoTo NextR
        If StrComp(prod, product, vbTextCompare) <> 0 Then GoTo NextR

        tag = Trim$(CStr(wsMap.Cells(r, 2).value))
        If Len(tag) = 0 Then GoTo NextR
        If HeaderCol(hdr, tag) = 0 Then GoTo NextR  ' require exact tag present in Paste Data

        role = InferRole(tag)
        If d.Exists(role) Then d(role).Add tag

NextR:
    Next r

    Set GroupTagsByRole = d
End Function

Private Function InferRole(ByVal tag As String) As String
    Dim s$: s = UCase$(tag)
    If InStr(s, "E3") > 0 And InStr(s, "FT") > 0 Then InferRole = ROLE_E3FT: Exit Function
    If InStr(s, "TT") > 0 Then InferRole = ROLE_TT: Exit Function
    If InStr(s, "PT") > 0 Then InferRole = ROLE_PT: Exit Function
    If InStr(s, "FT") > 0 Then InferRole = ROLE_FT: Exit Function
    InferRole = ROLE_FT
End Function

'======================== SPECS & OUTPUT =====================
Private Sub WriteKOV(ByVal wsOut As Worksheet, ByVal product As String, roleTags As Object, _
                     ByVal nTT As Long, ByVal dTT As Double, ByVal vTT As Double, _
                     ByVal nPT As Long, ByVal dPT As Double, ByVal vPT As Double, _
                     ByVal nFT As Long, ByVal dFT As Double, ByVal vFT As Double, _
                     ByVal nE3 As Long, ByVal dE3 As Double, ByVal vE3 As Double, _
                     ByRef t() As Double, ByVal SoakStartIdx As Long, ByVal SoakEndIdx As Long, ByVal SoakEndCommitIdx As Long, _
                     ByVal soakH As Double, ByVal soakT As Double, _
                     ByVal StripStartIdx As Long, ByVal StripEndIdx As Long, _
                     ByVal stripH As Double, ByVal stripT As Double, _
                     ByRef sc As SoakCfg, ByRef rc As StripCfg)

    Dim rr As Long, notes As String, metricName As String, metricUnits As String

    With wsOut
        .Cells.ClearContents
        ' Top summary block
        .Range("A1:F1").value = Array("Product", "Role", "Tags used", _
                              "N", "Max", "StdDev")
        .Range("A2").value = product
        rr = 2
        rr = PrintRoleSummary(wsOut, rr, "TT", roleTags, nTT, dTT, vTT, product)
        rr = PrintRoleSummary(wsOut, rr, "PT", roleTags, nPT, dPT, vPT)
        rr = PrintRoleSummary(wsOut, rr, "FT", roleTags, nFT, dFT, vFT)
        rr = PrintRoleSummary(wsOut, rr, "E3FT", roleTags, nE3, dE3, vE3)
        rr = rr + 1

        ' Results header (A..L)
        .Range("A" & rr & ":L" & rr).value = Array("Stage", "Start Time", "End Time", _
                                                   "Metric", "Value", _
                                                   "Min", "TV", "Max", "Result", "# from TV", "Label", "Notes")
        .Range("A" & rr & ":L" & rr).Font.Bold = True
        rr = rr + 1

        ' ===== SOAK =====
        If SoakStartIdx > 0 Then
            ' Soak Temperature
            If Not LimitVarUnits(product, "Soak", "Temperature", metricName, metricUnits) Then
                metricName = "Temperature": metricUnits = "F"
            End If
            .Cells(rr, 1).value = "Soak"
            .Cells(rr, 2).value = IIf(SoakStartIdx > 0, t(SoakStartIdx), "")
            .Cells(rr, 3).value = IIf(SoakEndIdx > 0, t(SoakEndIdx), "")
            .Cells(rr, 2).NumberFormat = "m/dd/yyyy hh:mm": .Cells(rr, 3).NumberFormat = "m/dd/yyyy hh:mm"
            .Cells(rr, 4).value = metricName & " (" & metricUnits & ")"
            If soakT > 0 Then .Cells(rr, 5).value = Round(soakT, 1)

            notes = "SoakStart: FT in [" & sc.FT_Lo & "–" & sc.FT_Hi & "] for " & sc.inHoldMin & "m, " & _
                    "then FT > " & sc.FT_Cross & " for " & sc.outHoldMin & "m. " & _
                    "SoakEnd: PT = peak" & Format(sc.fallFrac, "0%") & " for " & sc.FallHoldMin & "m; " & _
                    "backtracked to peak within " & sc.peakWinMin & "m window. " & _
                    "Trim " & sc.trimIn & "/" & sc.trimOut & "m."
            .Cells(rr, 12).value = notes

            WriteSpecCompare wsOut, rr, product, soakT, "Soak", "Temperature"
            rr = rr + 1

            ' Soak Time
            If Not LimitVarUnits(product, "Soak", "Time", metricName, metricUnits) Then
                metricName = "Time": metricUnits = "h"
            End If
            .Cells(rr, 1).value = "Soak"
            .Cells(rr, 2).value = IIf(SoakStartIdx > 0, t(SoakStartIdx), "")
            .Cells(rr, 3).value = IIf(SoakEndIdx > 0, t(SoakEndIdx), "")
            .Cells(rr, 2).NumberFormat = "m/dd/yyyy hh:mm": .Cells(rr, 3).NumberFormat = "m/dd/yyyy hh:mm"
            .Cells(rr, 4).value = metricName & " (" & metricUnits & ")"
            If soakH > 0 Then .Cells(rr, 5).value = Round(soakH, 2)
            .Cells(rr, 12).value = "Duration from SoakStart to SoakEnd."
            WriteSpecCompare wsOut, rr, product, soakH, "Soak", "Time"
            rr = rr + 1
        End If

        ' ===== STRIP =====
        If StripStartIdx > 0 Then
            ' Strip Temperature
            If Not LimitVarUnits(product, "Strip", "Temperature", metricName, metricUnits) Then
                metricName = "Temperature": metricUnits = "F"
            End If
            .Cells(rr, 1).value = "Strip"
            .Cells(rr, 2).value = t(StripStartIdx)
            .Cells(rr, 3).value = IIf(StripEndIdx > 0, t(StripEndIdx), "")
            .Cells(rr, 2).NumberFormat = "m/dd/yyyy hh:mm": .Cells(rr, 3).NumberFormat = "m/dd/yyyy hh:mm"
            .Cells(rr, 4).value = metricName & " (" & metricUnits & ")"
            If stripT > 0 Then .Cells(rr, 5).value = Round(stripT, 1)

            notes = "StripStart: PT = " & rc.PT_Thresh & " for " & rc.StartHoldMin & "m. " & _
                    "StripEnd: E3FT = E0 + " & rc.DeltaFromStart & " for " & rc.EndHoldMin & "m. " & _
                    "Trim " & rc.trimIn & "/" & rc.trimOut & "m."
            .Cells(rr, 12).value = notes

            WriteSpecCompare wsOut, rr, product, stripT, "Strip", "Temperature"
            rr = rr + 1

            ' Strip Time (calc only)
            .Cells(rr, 1).value = "Strip"
            .Cells(rr, 2).value = t(StripStartIdx)
            .Cells(rr, 3).value = IIf(StripEndIdx > 0, t(StripEndIdx), "")
            .Cells(rr, 2).NumberFormat = "m/dd/yyyy hh:mm": .Cells(rr, 3).NumberFormat = "m/dd/yyyy hh:mm"
            .Cells(rr, 4).value = "Time (h)"
            If stripH > 0 Then .Cells(rr, 5).value = Round(stripH, 2)
            .Cells(rr, 12).value = "Duration from StripStart to StripEnd."
            rr = rr + 1
        End If

        Dim wsK As Worksheet
        Set wsK = ThisWorkbook.Worksheets("KOV")

        KOV_ColorizeAllTables wsK
        .Columns("A:L").AutoFit
    End With
End Sub

Private Function LimitVarUnits(ByVal product As String, ByVal stepName As String, ByVal varKey As String, _
                               ByRef outVar As String, ByRef outUnits As String) As Boolean
    Dim ws As Worksheet, r As Long, lastRow As Long
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(SH_LIMITS): On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If StrComp(CStr(ws.Cells(r, 1).value), product, vbTextCompare) = 0 _
        And StrComp(CStr(ws.Cells(r, 2).value), stepName, vbTextCompare) = 0 _
        And InStr(1, LCase$(CStr(ws.Cells(r, 3).value)), LCase$(varKey)) > 0 Then
            outVar = CStr(ws.Cells(r, 3).value)
            outUnits = CStr(ws.Cells(r, 4).value)
            LimitVarUnits = (Len(outVar) > 0)
            Exit Function
        End If
    Next r
End Function

Private Sub WriteSpecCompare(ByVal wsOut As Worksheet, ByVal rowOut As Long, _
                             ByVal product As String, ByVal measured As Double, _
                             ByVal stepName As String, ByVal varKey As String)

    Dim vMin As Double, vTV As Double, vMax As Double, vUnits As String, vLabel As String
    Dim ok As Boolean
    ok = GetLimit(product, stepName, varKey, vMin, vTV, vMax, vUnits, vLabel)

    With wsOut
        If ok Then
            .Cells(rowOut, "F").value = vMin
            .Cells(rowOut, "G").value = vTV
            .Cells(rowOut, "H").value = vMax

            Dim pf As String
            If measured = 0 Then
                pf = ""
            ElseIf measured >= vMin And measured <= vMax Then
                pf = "PASS"
            Else
                pf = "FAIL"
            End If
            .Cells(rowOut, "I").value = pf
            If measured <> 0 Then .Cells(rowOut, "J").value = Round(measured - vTV, 2)
            .Cells(rowOut, "K").value = vLabel

            Dim nf$: nf = IIf(varKey = "Time", "0.00", "0")
            .Cells(rowOut, "F").NumberFormat = nf
            .Cells(rowOut, "G").NumberFormat = nf
            .Cells(rowOut, "H").NumberFormat = nf
            .Cells(rowOut, "J").NumberFormat = nf

            If pf = "PASS" Then .Cells(rowOut, "I").Interior.Color = RGB(198, 239, 206)
            If pf = "FAIL" Then .Cells(rowOut, "I").Interior.Color = RGB(255, 199, 206)

            If UCase$(vLabel) = "KOV" Then .Cells(rowOut, "K").Interior.Color = RGB(0, 255, 0)
            If UCase$(vLabel) = "AOV" Then .Cells(rowOut, "K").Interior.Color = RGB(230, 230, 230)
        End If
    End With
End Sub

Private Function GetLimit(ByVal product As String, ByVal stepName As String, ByVal varKey As String, _
                          ByRef vMin As Double, ByRef vTV As Double, ByRef vMax As Double, _
                          Optional ByRef vUnits As String, Optional ByRef vLabel As String) As Boolean
    Dim ws As Worksheet, r As Long, lastRow As Long
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(SH_LIMITS): On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If StrComp(CStr(ws.Cells(r, 1).value), product, vbTextCompare) = 0 _
        And StrComp(CStr(ws.Cells(r, 2).value), stepName, vbTextCompare) = 0 _
        And InStr(1, LCase$(CStr(ws.Cells(r, 3).value)), LCase$(varKey)) > 0 Then
            vUnits = CStr(ws.Cells(r, 4).value)
            vMin = CDbl(val(ws.Cells(r, 5).value))
            vTV = CDbl(val(ws.Cells(r, 6).value))
            vMax = CDbl(val(ws.Cells(r, 7).value))
            vLabel = CStr(ws.Cells(r, 8).value)
            GetLimit = True
            Exit Function
        End If
    Next r
End Function

