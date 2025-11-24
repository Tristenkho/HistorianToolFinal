Attribute VB_Name = "modKOV_WeekRunner"
Option Explicit

Public Sub KOV_Run_FromBatchSummary()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsBS As Worksheet, wsK As Worksheet, wsKM As Worksheet

    On Error Resume Next
    Set wsBS = wb.Worksheets("Batch Summary")
    Set wsK = wb.Worksheets("KOV")
    On Error GoTo 0
    If wsBS Is Nothing Then MsgBox "Batch Summary not found.", vbExclamation: Exit Sub
    If wsK Is Nothing Then MsgBox "KOV sheet not found.", vbExclamation: Exit Sub

    ' Create/clear consolidated output
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

    Dim lastRow As Long: lastRow = wsBS.Cells(wsBS.rows.Count, 1).End(xlUp).Row
    Dim r As Long

    ' Silence per-batch to avoid popups, show progress in status bar
    Dim wasSilent As Boolean: wasSilent = G_KOV_Silent
    G_KOV_Silent = True
    Application.ScreenUpdating = False

    For r = 2 To lastRow
        Dim tag As String, prod As String
        Dim st As Variant, en As Variant

        tag = Trim$(CStr(wsBS.Cells(r, 1).value))
        st = wsBS.Cells(r, 2).value
        en = wsBS.Cells(r, 3).value
        prod = Trim$(CStr(wsBS.Cells(r, 7).value))

        If Len(prod) = 0 Or Not IsDate(st) Or Not IsDate(en) Then GoTo NextR

        Application.StatusBar = "KOV running " & prod & " (row " & r & ")..."

        ' Set window per batch
        KOV_SetWindow CDbl(st) - (1# / 24#), CDbl(en)
        G_SELECTED_PRODUCT = prod

        ' Clear KOV sheet used area
        With wsK
            .Cells.Clear
            .Cells.FormatConditions.Delete
            .Cells.Interior.Pattern = xlNone
            .Cells.Borders.LineStyle = xlNone
        End With

        ' Dispatch (errors trapped per-row)
        On Error Resume Next
        Application.Run DispatchTargetFor(prod)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextR
        End If
        On Error GoTo 0

        ' Copy latest KOV into consolidated
        Dim used As Range
        Set used = wsK.UsedRange

        wsKM.Cells(outRow, 1).value = "Row " & r & " | " & prod & _
            " | Window: " & Format(CDbl(st) - (1# / 24#), "m/d/yyyy hh:mm") & " – " & _
            Format(CDbl(en), "m/d/yyyy hh:mm") & IIf(Len(tag) > 0, " | Tag: " & tag, "")
        wsKM.Cells(outRow, 1).Font.Bold = True
        outRow = outRow + 1

        If Not used Is Nothing Then
            If Application.WorksheetFunction.CountA(used) > 0 Then
                used.Copy wsKM.Cells(outRow, 1)
                outRow = outRow + used.rows.Count + 2
                Application.CutCopyMode = False
            Else
                outRow = outRow + 1
            End If
        End If

        ' Reset for next
        G_SELECTED_PRODUCT = vbNullString
        KOV_ClearWindow

NextR:
    Next r

    KOV_ColorizeAllTables wsKM
    wsKM.Columns("A:L").AutoFit

    ' restore UI and notify once
    Application.ScreenUpdating = True
    Application.StatusBar = False
    G_KOV_Silent = wasSilent
    KOV_Notify "KOV Multi complete (see 'KOV Multi')."
End Sub

' Small helper to build the Application.Run target string for a product name
Private Function DispatchTargetFor(ByVal prod As String) As String
    Dim key As String
    key = UCase$(Replace(Replace(prod, " ", ""), ".", "")) ' normalize e.g., "116.58"

    Dim wbName As String: wbName = ThisWorkbook.name
    Select Case key
        Case "LUBRIZOL19858", "19858":          DispatchTargetFor = "'" & wbName & "'!KOV_Run_Lubrizol19858_Main"
        Case "INFINEUMC9242", "C9242":          DispatchTargetFor = "'" & wbName & "'!KOV_Run_InfineumC9242_Main"
        Case "INFINEUMC9402", "C9402":          DispatchTargetFor = "'" & wbName & "'!KOV_Run_v2_Main"
        Case "INFINEUMC9411", "C9411":          DispatchTargetFor = "'" & wbName & "'!KOV_Run_v2_Main"
        Case "INNOSPECASA", "ASA":              DispatchTargetFor = "'" & wbName & "'!KOV_Run_InnospecASA_Main"
        Case "INFINEUMC9412", "C9412":          DispatchTargetFor = "'" & wbName & "'!KOV_Run_InfineumC9412_Main"
        Case "LUBRIZOL02766", "02766":          DispatchTargetFor = "'" & wbName & "'!KOV_Run_Lubrizol02766_Main"
        Case "INFINEUMC9283", "C9283":          DispatchTargetFor = "'" & wbName & "'!KOV_Run_InfineumC9283_Main"
        Case "LUBRIZOL11658", "11658":          DispatchTargetFor = "'" & wbName & "'!KOV_Run_Lubrizol11658_Main"
        Case "INNOSPECOLI9000M", "OLI9000M":    DispatchTargetFor = "'" & wbName & "'!KOV_Run_InnospecOLI9000M_Main"
        Case "INNOSPECOLI9200LN", "OLI9200LN":  DispatchTargetFor = "'" & wbName & "'!KOV_Run_InnospecOLI9200LN_Main"
        Case Else
            DispatchTargetFor = "'" & wbName & "'!KOV_Run_v2_Main" ' safe default
    End Select
End Function


