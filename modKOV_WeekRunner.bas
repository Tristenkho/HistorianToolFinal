Attribute VB_Name = "modKOV_WeekRunner"
Public Sub KOV_Run_FromBatchSummary()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsBS As Worksheet, wsK As Worksheet, wsKM As Worksheet

    On Error Resume Next
    Set wsBS = wb.Worksheets("Batch Summary")
    Set wsK = wb.Worksheets("KOV")
    On Error GoTo 0

    ' Show fatal errors (not silenced)
    If wsBS Is Nothing Then
        MsgBox "Batch Summary not found.", vbExclamation
        Exit Sub
    End If
    If wsK Is Nothing Then
        MsgBox "KOV sheet not found.", vbExclamation
        Exit Sub
    End If

    ' Create/clear KOV Multi
    On Error Resume Next
    Set wsKM = wb.Worksheets("KOV Multi")
    On Error GoTo 0
    If wsKM Is Nothing Then
        Set wsKM = wb.Worksheets.Add(After:=wsK)
        wsKM.name = "KOV Multi"
    Else
        wsKM.Cells.Clear
    End If
    wsKM.Range("A1").value = "Consolidated KOV (Week)"
    wsKM.Range("A1").Font.Bold = True
    Dim outRow As Long: outRow = 3

    Dim lastRow As Long: lastRow = wsBS.Cells(wsBS.Rows.Count, 1).End(xlUp).Row
    Dim r As Long

    ' Silence per-batch popups only during the loop
    Dim wasSilent As Boolean: wasSilent = G_KOV_Silent
    G_KOV_Silent = True
    Application.ScreenUpdating = False

    For r = 2 To lastRow
        Dim tag As String, prod As String
        Dim st As Variant, en As Variant

        tag = Trim$(CStr(wsBS.Cells(r, 1).value))   ' Tag
        st = wsBS.Cells(r, 2).value                 ' Batch Start
        en = wsBS.Cells(r, 3).value                 ' Batch End
        prod = Trim$(CStr(wsBS.Cells(r, 7).value))  ' Product (col G)

        If Len(prod) = 0 Then GoTo nextR
        If Not IsDate(st) Or Not IsDate(en) Then GoTo nextR

        ' Local window (also used for header text)
        Dim winStart As Double, winEnd As Double
        winStart = CDbl(st) - (1# / 24#)    ' 1h pre-start buffer
        winEnd = CDbl(en)

        ' Set globals for product runners
        KOV_SetWindow winStart, winEnd
        G_SELECTED_PRODUCT = prod

        ' Clear KOV sheet before each run
        With wsK
            .Cells.Clear
            .Cells.FormatConditions.Delete
            .Cells.Interior.Pattern = xlNone
            .Cells.Borders.LineStyle = xlNone
        End With

        ' Dispatch to product KOV
        Dim tgt As String
        Select Case UCase$(Replace(prod, " ", ""))
            Case "LUBRIZOL198.58", "198.58": tgt = "'" & wb.name & "'!KOV_Run_Lubrizol19858_Main"
            Case "INFINEUMC9242", "C9242":   tgt = "'" & wb.name & "'!KOV_Run_InfineumC9242_Main"
            Case "INFINEUMC9402", "C9402":   tgt = "'" & wb.name & "'!KOV_Run_v2_Main"
            Case "INFINEUMC9411", "C9411":   tgt = "'" & wb.name & "'!KOV_Run_v2_Main"
            Case "INNOSPECASA", "ASA":       tgt = "'" & wb.name & "'!KOV_Run_InnospecASA_Main"
            Case "INFINEUMC9412", "C9412":   tgt = "'" & wb.name & "'!KOV_Run_InfineumC9412_Main"
            Case "LUBRIZOL0276.6", "0276.6": tgt = "'" & wb.name & "'!KOV_Run_Lubrizol02766_Main"
            Case "INFINEUMC9283", "C9283":   tgt = "'" & wb.name & "'!KOV_Run_InfineumC9283_Main"
            Case "LUBRIZOL116.58", "116.58": tgt = "'" & wb.name & "'!KOV_Run_Lubrizol11658_Main"
            Case "INNOSPECOLI9000M", "OLI9000M":   tgt = "'" & wb.name & "'!KOV_Run_InnospecOLI9000M_Main"
            Case "INNOSPECOLI9200LN", "OLI9200LN": tgt = "'" & wb.name & "'!KOV_Run_InnospecOLI9200LN_Main"
            Case Else: tgt = ""
        End Select
        If Len(tgt) > 0 Then
            Application.Run tgt
        End If

afterRun:
        ' Copy latest KOV into consolidated sheet
        Dim used As Range
        Set used = wsK.UsedRange

        wsKM.Cells(outRow, 1).value = "Row " & r & " | " & prod & _
            " | Window: " & Format(winStart, "m/d/yyyy hh:mm") & " – " & _
            Format(winEnd, "m/d/yyyy hh:mm") & IIf(Len(tag) > 0, " | Tag: " & tag, "")
        wsKM.Cells(outRow, 1).Font.Bold = True
        outRow = outRow + 1

        If Not used Is Nothing Then
            If Application.WorksheetFunction.CountA(used) > 0 Then
                used.Copy wsKM.Cells(outRow, 1)
                outRow = outRow + used.Rows.Count + 2
                Application.CutCopyMode = False
            Else
                outRow = outRow + 1
            End If
        End If

        ' Reset globals for next row
        G_SELECTED_PRODUCT = vbNullString
        KOV_ClearWindow

nextR:
    Next r

    KOV_ColorizeAllTables wsKM
    wsKM.Columns("A:L").AutoFit

    ' restore UI state and notify once
    Application.ScreenUpdating = True
    G_KOV_Silent = wasSilent
    KOV_Notify "KOV Multi complete (see 'KOV Multi')."
End Sub

Sub R4_Run_KOV_For_Week()
    KOV_Run_FromBatchSummary
End Sub

