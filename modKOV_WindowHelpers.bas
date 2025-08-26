Attribute VB_Name = "modKOV_WindowHelpers"
Option Explicit

' ----------------- public: window bounds from globals -----------------
Public Sub ResolveWindowBounds(ByRef t() As Double, ByRef i0 As Long, ByRef i1 As Long)
    i0 = 2: i1 = UBound(t)
    If G_KOV_UseWindow Then
        Dim a As Long, b As Long
        a = FindIndexGE(t, G_KOV_WindowStart)
        b = FindIndexLE(t, G_KOV_WindowEnd)
        If a > 0 Then i0 = a
        If b > 0 Then i1 = b
        If i1 <= i0 Then i0 = 2: i1 = UBound(t) ' fallback to full run
    End If
End Sub

' ----------------- public: generic range-limited detectors ------------
Public Function FirstHold_InBand_Range(ByRef y() As Double, ByVal lo As Double, ByVal hi As Double, _
                                       ByRef t() As Double, ByVal holdMin As Double, _
                                       ByVal iStart As Long, ByVal iEnd As Long) As Long
    Dim i As Long, dt As Double, acc As Double, startIdx As Long
    If iStart < 2 Then iStart = 2
    If iEnd <= iStart Then Exit Function
    startIdx = 0: acc = 0#
    For i = iStart To iEnd
        dt = MinutesBetween_KOVH(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If y(i) >= lo And y(i) <= hi Then
            If startIdx = 0 Then startIdx = i
            acc = acc + dt
            If acc >= holdMin Then FirstHold_InBand_Range = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0#
        End If
    Next i
End Function

Public Function FirstHold_OutOfBand_Range(ByRef y() As Double, ByVal lo As Double, ByVal hi As Double, _
                                          ByRef t() As Double, ByVal iStart As Long, ByVal holdMin As Double, _
                                          ByVal iEnd As Long) As Long
    Dim i As Long, dt As Double, acc As Double, startIdx As Long
    If iStart < 2 Then iStart = 2
    If iEnd <= iStart Then Exit Function
    startIdx = 0: acc = 0#
    For i = iStart To iEnd
        dt = MinutesBetween_KOVH(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If (y(i) <= lo) Or (y(i) >= hi) Then
            If startIdx = 0 Then startIdx = i
            acc = acc + dt
            If acc >= holdMin Then FirstHold_OutOfBand_Range = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0#
        End If
    Next i
End Function

Public Function FirstHold_Single_Range(ByRef y() As Double, ByVal op As String, ByVal th As Double, _
                                       ByRef t() As Double, ByVal iStart As Long, ByVal holdMin As Double, _
                                       ByVal iEnd As Long) As Long
    Dim i As Long, dt As Double, acc As Double, startIdx As Long, ok As Boolean
    If iStart < 2 Then iStart = 2
    If iEnd <= iStart Then Exit Function
    startIdx = 0: acc = 0#
    For i = iStart To iEnd
        dt = MinutesBetween_KOVH(t(i - 1), t(i)): If dt < 0 Then dt = 0
        ok = (op = "<=" And y(i) <= th) Or (op = ">=" And y(i) >= th)
        If ok Then
            If startIdx = 0 Then startIdx = i
            acc = acc + dt
            If acc >= holdMin Then FirstHold_Single_Range = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0#
        End If
    Next i
End Function

Public Function FirstHold_Dual_Range(ByRef a() As Double, ByVal opA As String, ByVal thA As Double, _
                                     ByRef b() As Double, ByVal opB As String, ByVal thB As Double, _
                                     ByRef t() As Double, ByVal holdMin As Double, _
                                     ByVal iStart As Long, ByVal iEnd As Long) As Long
    Dim i As Long, dt As Double, acc As Double, startIdx As Long, okA As Boolean, okB As Boolean
    If iStart < 2 Then iStart = 2
    If iEnd <= iStart Then Exit Function
    startIdx = 0: acc = 0#
    For i = iStart To iEnd
        dt = MinutesBetween_KOVH(t(i - 1), t(i)): If dt < 0 Then dt = 0
        okA = (opA = "<=" And a(i) <= thA) Or (opA = ">=" And a(i) >= thA)
        okB = (opB = "<=" And b(i) <= thB) Or (opB = ">=" And b(i) >= thB)
        If okA And okB Then
            If startIdx = 0 Then startIdx = i
            acc = acc + dt
            If acc >= holdMin Then FirstHold_Dual_Range = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0#
        End If
    Next i
End Function

Public Function FirstHold_CFT_DeltaRise_Range(ByRef cFT() As Double, ByRef t() As Double, _
                                              ByVal iStartIdx As Long, ByVal deltaReq As Double, _
                                              ByVal holdMin As Double, ByVal iEnd As Long) As Long
    If iStartIdx <= 0 Then Exit Function
    If iEnd <= iStartIdx Then Exit Function
    Dim base As Double: base = cFT(iStartIdx)
    Dim i As Long, dt As Double, acc As Double, startIdx As Long
    startIdx = 0: acc = 0#
    For i = iStartIdx + 1 To iEnd
        dt = MinutesBetween_KOVH(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If cFT(i) >= base + deltaReq Then
            If startIdx = 0 Then startIdx = i
            acc = acc + dt
            If acc >= holdMin Then FirstHold_CFT_DeltaRise_Range = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0#
        End If
    Next i
End Function

' ----------------- public: v2 (C9402/C9411) range variants -------------
Public Function SoakStart_MaleicThenCross_Range(ByRef ft() As Double, ByRef t() As Double, _
    ByVal lo As Double, ByVal hi As Double, ByVal inHoldMin As Double, _
    ByVal cross As Double, ByVal outHoldMin As Double, ByVal iStart As Long, ByVal iEnd As Long) As Long

    Dim i As Long, dt As Double, inAcc As Double, outAcc As Double, eligible As Boolean
    Dim i0 As Long, startIdx As Long
    If iStart < 2 Then iStart = 2
    If iEnd <= iStart Then Exit Function

    ' 1) in-band accumulation
    inAcc = 0#: eligible = False
    For i = iStart + 1 To iEnd
        dt = MinutesBetween_KOVH(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If ft(i) >= lo And ft(i) <= hi Then inAcc = inAcc + dt
        If Not eligible And inAcc >= inHoldMin Then eligible = True: Exit For
    Next i
    If Not eligible Then Exit Function
    i0 = i

    ' 2) FT > Cross with hold; mark start of hold
    startIdx = 0: outAcc = 0#
    For i = i0 To iEnd
        dt = MinutesBetween_KOVH(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If ft(i) > cross Then
            If startIdx = 0 Then startIdx = i
            outAcc = outAcc + dt
            If outAcc >= outHoldMin Then SoakStart_MaleicThenCross_Range = startIdx: Exit Function
        Else
            startIdx = 0: outAcc = 0#
        End If
    Next i
End Function

Public Function SoakEnd_FallFromPeak_Backtrack_Range(ByRef pt() As Double, ByRef t() As Double, _
    ByVal SoakStartIdx As Long, ByVal peakWinMin As Double, ByVal fallFrac As Double, _
    ByVal holdMin As Double, ByRef CommitIdx As Long, ByVal iEnd As Long) As Long

    Dim i As Long, dt As Double, holdAcc As Double, peak As Double
    CommitIdx = 0: holdAcc = 0#
    If SoakStartIdx <= 0 Then Exit Function
    If iEnd <= SoakStartIdx Then Exit Function

    For i = Application.Max(SoakStartIdx + 1, 3) To iEnd
        peak = RollingPeakLastMinutes_Range(pt, t, i, peakWinMin, SoakStartIdx)
        If peak > 0 And pt(i) <= peak * fallFrac Then
            dt = MinutesBetween_KOVH(t(i - 1), t(i)): If dt < 0 Then dt = 0
            holdAcc = holdAcc + dt
            If holdAcc >= holdMin Then CommitIdx = i: Exit For
        Else
            holdAcc = 0#
        End If
    Next i
    If CommitIdx = 0 Then Exit Function

    SoakEnd_FallFromPeak_Backtrack_Range = PeakIndexLastMinutes_Range(pt, t, CommitIdx - 1, peakWinMin, SoakStartIdx)
End Function

Public Function FirstHoldBelow_Range(ByRef v() As Double, ByRef t() As Double, _
    ByVal iStart As Long, ByVal th As Double, ByVal holdMin As Double, ByVal iEnd As Long) As Long

    Dim i As Long, dt As Double, acc As Double, startIdx As Long
    If iStart < 2 Then iStart = 2
    If iEnd <= iStart Then Exit Function
    startIdx = 0: acc = 0#
    For i = iStart To iEnd
        dt = MinutesBetween_KOVH(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If v(i) <= th Then
            If startIdx = 0 Then startIdx = i
            acc = acc + dt
            If acc >= holdMin Then FirstHoldBelow_Range = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0#
        End If
    Next i
End Function

Public Function StripEnd_E3FT_DeltaHold_Range(ByRef E3FT() As Double, ByRef t() As Double, _
    ByVal StripStartIdx As Long, ByVal deltaReq As Double, ByVal holdMin As Double, ByVal iEnd As Long) As Long

    If StripStartIdx <= 0 Then Exit Function
    If iEnd <= StripStartIdx Then Exit Function
    Dim base As Double: base = E3FT(StripStartIdx)
    Dim i As Long, dt As Double, acc As Double, startIdx As Long
    startIdx = 0: acc = 0#
    For i = StripStartIdx + 1 To iEnd
        dt = MinutesBetween_KOVH(t(i - 1), t(i)): If dt < 0 Then dt = 0
        If E3FT(i) >= base + deltaReq Then
            If startIdx = 0 Then startIdx = i
            acc = acc + dt
            If acc >= holdMin Then StripEnd_E3FT_DeltaHold_Range = startIdx: Exit Function
        Else
            startIdx = 0: acc = 0#
        End If
    Next i
End Function

' ----------------- private helpers --------------------
Private Function FindIndexGE(ByRef t() As Double, ByVal ts As Double) As Long
    Dim i As Long
    For i = 2 To UBound(t)
        If t(i) >= ts Then FindIndexGE = i: Exit Function
    Next i
End Function

Private Function FindIndexLE(ByRef t() As Double, ByVal ts As Double) As Long
    Dim i As Long
    For i = UBound(t) To 2 Step -1
        If t(i) <= ts Then FindIndexLE = i: Exit Function
    Next i
End Function

Private Function RollingPeakLastMinutes_Range(ByRef v() As Double, ByRef t() As Double, _
    ByVal iNow As Long, ByVal lookMin As Double, ByVal lowerIdx As Long) As Double

    Dim back As Double, i As Long, seg As Double, pk As Double, first As Boolean
    If iNow < 2 Then RollingPeakLastMinutes_Range = v(1): Exit Function
    back = 0#: first = True: pk = v(iNow)
    For i = iNow To Application.Max(lowerIdx + 1, 2) Step -1
        seg = MinutesBetween_KOVH(t(i - 1), t(i)): If seg <= 0 Then seg = 1
        back = back + seg
        If first Or v(i) > pk Then pk = v(i): first = False
        If back >= lookMin Then Exit For
    Next i
    RollingPeakLastMinutes_Range = pk
End Function

Private Function PeakIndexLastMinutes_Range(ByRef v() As Double, ByRef t() As Double, _
    ByVal iNow As Long, ByVal lookMin As Double, ByVal lowerIdx As Long) As Long

    Dim back As Double, i As Long, seg As Double, pk As Double, pki As Long, first As Boolean
    If iNow < 2 Then PeakIndexLastMinutes_Range = 1: Exit Function
    back = 0#: first = True: pk = v(iNow): pki = iNow
    For i = iNow To Application.Max(lowerIdx + 1, 2) Step -1
        seg = MinutesBetween_KOVH(t(i - 1), t(i)): If seg <= 0 Then seg = 1
        back = back + seg
        If first Or v(i) > pk Then pk = v(i): pki = i: first = False
        If back >= lookMin Then Exit For
    Next i
    PeakIndexLastMinutes_Range = pki
End Function

Private Function MinutesBetween_KOVH(ByVal t0 As Double, ByVal t1 As Double) As Double
    MinutesBetween_KOVH = (t1 - t0) * 24# * 60#
End Function

'==================== Back-compat shims (auto-apply window) ====================

Public Function FirstHold_InBand( _
    ByRef y() As Double, ByVal lo As Double, ByVal hi As Double, _
    ByRef t() As Double, ByVal holdMin As Double) As Long

    Dim i0 As Long, i1 As Long
    ResolveWindowBounds t, i0, i1
    FirstHold_InBand = FirstHold_InBand_Range(y, lo, hi, t, holdMin, i0, i1)
End Function

Public Function FirstHold_OutOfBand( _
    ByRef y() As Double, ByVal lo As Double, ByVal hi As Double, _
    ByRef t() As Double, ByVal iStart As Long, ByVal holdMin As Double) As Long

    Dim i0 As Long, i1 As Long
    ResolveWindowBounds t, i0, i1
    If iStart < i0 Then iStart = i0
    FirstHold_OutOfBand = FirstHold_OutOfBand_Range(y, lo, hi, t, iStart, holdMin, i1)
End Function

Public Function FirstHold_Single( _
    ByRef y() As Double, ByVal op As String, ByVal th As Double, _
    ByRef t() As Double, ByVal iStart As Long, ByVal holdMin As Double) As Long

    Dim i0 As Long, i1 As Long
    ResolveWindowBounds t, i0, i1
    If iStart < i0 Then iStart = i0
    FirstHold_Single = FirstHold_Single_Range(y, op, th, t, iStart, holdMin, i1)
End Function

Public Function FirstHold_Dual( _
    ByRef a() As Double, ByVal opA As String, ByVal thA As Double, _
    ByRef b() As Double, ByVal opB As String, ByVal thB As Double, _
    ByRef t() As Double, ByVal holdMin As Double, _
    ByVal iStart As Long) As Long

    Dim i0 As Long, i1 As Long
    ResolveWindowBounds t, i0, i1
    If iStart < i0 Then iStart = i0
    FirstHold_Dual = FirstHold_Dual_Range(a, opA, thA, b, opB, thB, t, holdMin, iStart, i1)
End Function

Public Function FirstHold_CFT_DeltaRise( _
    ByRef cFT() As Double, ByRef t() As Double, _
    ByVal iStartIdx As Long, ByVal deltaReq As Double, ByVal holdMin As Double) As Long

    Dim i0 As Long, i1 As Long
    ResolveWindowBounds t, i0, i1
    If iStartIdx < i0 Then iStartIdx = i0
    FirstHold_CFT_DeltaRise = FirstHold_CFT_DeltaRise_Range(cFT, t, iStartIdx, deltaReq, holdMin, i1)
End Function

' ---- v2 patterns (C9402 / C9411) ----
Public Function SoakStart_MaleicThenCross( _
    ByRef ft() As Double, ByRef t() As Double, _
    ByVal lo As Double, ByVal hi As Double, ByVal inHoldMin As Double, _
    ByVal cross As Double, ByVal outHoldMin As Double) As Long

    Dim i0 As Long, i1 As Long
    ResolveWindowBounds t, i0, i1
    SoakStart_MaleicThenCross = SoakStart_MaleicThenCross_Range(ft, t, lo, hi, inHoldMin, cross, outHoldMin, i0, i1)
End Function

Public Function SoakEnd_FallFromPeak_Backtrack( _
    ByRef pt() As Double, ByRef t() As Double, _
    ByVal SoakStartIdx As Long, _
    ByVal peakWinMin As Double, ByVal fallFrac As Double, ByVal holdMin As Double, _
    ByRef CommitIdx As Long) As Long

    Dim i0 As Long, i1 As Long
    ResolveWindowBounds t, i0, i1
    SoakEnd_FallFromPeak_Backtrack = SoakEnd_FallFromPeak_Backtrack_Range( _
        pt, t, SoakStartIdx, peakWinMin, fallFrac, holdMin, CommitIdx, i1)
End Function

Public Function FirstHoldBelow( _
    ByRef v() As Double, ByRef t() As Double, _
    ByVal iStart As Long, ByVal th As Double, ByVal holdMin As Double) As Long

    Dim i0 As Long, i1 As Long
    ResolveWindowBounds t, i0, i1
    If iStart < i0 Then iStart = i0
    FirstHoldBelow = FirstHoldBelow_Range(v, t, iStart, th, holdMin, i1)
End Function

Public Function StripEnd_E3FT_DeltaHold( _
    ByRef E3FT() As Double, ByRef t() As Double, _
    ByVal StripStartIdx As Long, ByVal deltaReq As Double, ByVal holdMin As Double) As Long

    Dim i0 As Long, i1 As Long
    ResolveWindowBounds t, i0, i1
    If StripStartIdx < i0 Then StripStartIdx = i0
    StripEnd_E3FT_DeltaHold = StripEnd_E3FT_DeltaHold_Range(E3FT, t, StripStartIdx, deltaReq, holdMin, i1)
End Function

