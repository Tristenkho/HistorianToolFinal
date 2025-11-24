Attribute VB_Name = "KOV_Dispatch"
Option Explicit

' Map a product name to the right KOV runner.
' Works with many spellings (spaces/dots/hyphens ignored).
Public Function DispatchTargetFor(ByVal prod As String, Optional wb As Workbook) As String
    Dim key As String, macroName As String, prefix As String

    key = UCase$(Trim$(prod))
    key = Replace(key, " ", "")
    key = Replace(key, ".", "")
    key = Replace(key, "-", "")
    key = Replace(key, "/", "")

    Select Case key
        Case "INFINEUMC9242", "C9242"
            macroName = "KOV_Run_InfineumC9242_Main"

        Case "INFINEUMC9402", "C9402"
            macroName = "KOV_Run_v2_Main"

        Case "INFINEUMC9411", "C9411"
            macroName = "KOV_Run_v2_Main"

        Case "INFINEUMC9412", "C9412"
            macroName = "KOV_Run_InfineumC9412_Main"

        Case "INFINEUMC9283", "C9283"
            macroName = "KOV_Run_InfineumC9283_Main"

        Case "LUBRIZOL19858", "19858"
            macroName = "KOV_Run_Lubrizol19858_Main"

        ' 0276.6 normalizes to 02766 when dots are removed
        Case "LUBRIZOL02766", "02766", "0276.6"
            macroName = "KOV_Run_Lubrizol02766_Main"

        Case "LUBRIZOL11658", "11658"
            macroName = "KOV_Run_Lubrizol11658_Main"

        Case "INNOSPECASA", "ASA"
            macroName = "KOV_Run_InnospecASA_Main"

        Case "INNOSPECOLI9000M", "OLI9000M"
            macroName = "KOV_Run_InnospecOLI9000M_Main"

        Case "INNOSPECOLI9200LN", "OLI9200LN"
            macroName = "KOV_Run_InnospecOLI9200LN_Main"

        Case Else
            macroName = "KOV_Run_v2_Main"  ' sensible default
    End Select

    If wb Is Nothing Then
        prefix = ""
    Else
        prefix = "'" & wb.name & "'!"
    End If

    DispatchTargetFor = prefix & macroName
End Function

Public Sub KOV_Run_Dispatch()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ui As Worksheet: Set ui = wb.Worksheets("UI")

    Dim prod As String: prod = Trim$(CStr(ui.Range("B1").value))
    If Len(prod) = 0 Then
        MsgBox "Pick a Product in UI!B1.", vbExclamation
        Exit Sub
    End If

    Dim tgt As String: tgt = DispatchTargetFor(prod)

    On Error GoTo fallback
    Application.Run tgt
    Exit Sub

fallback:
    On Error GoTo 0
    MsgBox "Couldn't run: " & tgt & vbCrLf & "Falling back to default engine.", vbExclamation
    Application.Run DispatchTargetFor("C9402") ' v2 engine
End Sub

