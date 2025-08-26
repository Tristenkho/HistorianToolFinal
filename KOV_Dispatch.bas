Attribute VB_Name = "KOV_Dispatch"
Option Explicit

Public Sub KOV_Run_Dispatch()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ui As Worksheet: Set ui = wb.Worksheets("UI")

    Dim prod As String, p As String, tgt As String
    prod = Trim$(CStr(ui.Range("B1").value))
    If Len(prod) = 0 Then MsgBox "Pick a Product in UI!B1.", vbExclamation: Exit Sub
    p = UCase$(Application.WorksheetFunction.Trim(prod))

    Select Case p
        Case "INFINEUM C9242", "C9242", "INFINEUMC9242"
            tgt = "'" & wb.name & "'!KOV_Run_InfineumC9242_Main"

        Case "INFINEUM C9402", "C9402", "INFINEUMC9402"
            tgt = "'" & wb.name & "'!KOV_Run_v2_Main"

        Case "INFINEUM C9411", "C9411", "INFINEUMC9411"
            tgt = "'" & wb.name & "'!KOV_Run_v2_Main"
            
        Case "LUBRIZOL 198.58", "198.58", "LUBRIZOL198.58"
            tgt = "'" & wb.name & "'!KOV_Run_Lubrizol19858_Main"
        
        Case "INNOSPEC ASA", "ASA"
            tgt = "'" & wb.name & "'!KOV_Run_InnospecASA_Main"
            
        Case "INFINEUM C9412", "C9412", "INFINEUMC9412"
            tgt = "'" & wb.name & "'!KOV_Run_InfineumC9412_Main"
        
        Case "LUBRIZOL 0276.6", "0276.6", "LUBRIZOL0276.6"
            tgt = "'" & wb.name & "'!KOV_Run_Lubrizol02766_Main"
        
        Case "INFINEUM C9283", "C9283", "INFINEUMC9283"
            tgt = "'" & wb.name & "'!KOV_Run_InfineumC9283_Main"

        Case "LUBRIZOL 116.58", "116.58", "LUBRIZOL116.58"
            tgt = "'" & wb.name & "'!KOV_Run_Lubrizol11658_Main"

        Case "INNOSPEC OLI 9000M", "OLI 9000M"
            tgt = "'" & wb.name & "'!KOV_Run_InnospecOLI9000M_Main"
        
        Case "INNOSPEC OLI 9200LN", "OLI 9200LN"
            tgt = "'" & wb.name & "'!KOV_Run_InnospecOLI9200LN_Main"
        
        Case Else
            tgt = "'" & wb.name & "'!KOV_Run_v2_Main"
        
    End Select

    On Error GoTo fallback
    Application.Run tgt
    Exit Sub
fallback:
    On Error GoTo 0
    MsgBox "Couldn't run: " & tgt & vbCrLf & "Falling back to default engine.", vbExclamation
    Application.Run "'" & wb.name & "'!KOV_Run_v2_Main"
End Sub

