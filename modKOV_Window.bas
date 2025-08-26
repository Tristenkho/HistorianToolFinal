Attribute VB_Name = "modKOV_Window"
Option Explicit

'------------- Global window flags (visible from all modules) -------------
Public G_SELECTED_PRODUCT  As String
Public G_KOV_UseWindow     As Boolean
Public G_KOV_WindowStart   As Double   ' Excel serial datetime
Public G_KOV_WindowEnd     As Double   ' Excel serial datetime

' Turn the window on in one call (optional helper)
Public Sub KOV_SetWindow(ByVal startTs As Double, ByVal endTs As Double)
    G_KOV_UseWindow = True
    G_KOV_WindowStart = startTs
    G_KOV_WindowEnd = endTs
End Sub

' Clear window after each run (WeekRunner calls this)
Public Sub KOV_ClearWindow()
    G_KOV_UseWindow = False
    G_KOV_WindowStart = 0#
    G_KOV_WindowEnd = 0#
End Sub

' (Optional) Handy for logging/debug
Public Function KOV_WindowText() As String
    If Not G_KOV_UseWindow Then
        KOV_WindowText = "Window: OFF"
    Else
        KOV_WindowText = "Window: " & Format(G_KOV_WindowStart, "m/d/yyyy hh:mm") & _
                         " – " & Format(G_KOV_WindowEnd, "m/d/yyyy hh:mm")
    End If
End Function


