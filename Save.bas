Attribute VB_Name = "Save"
Option Explicit

' === Set your archive folder (must already exist) ===
Private Const ARCHIVE_DIR As String = _
  "C:\Users\tristenk\OneDrive - South Coast Terminals\Desktop\process historian project\Versions\Archive"

' === Save an exact copy of this workbook (macros, charts, formulas) ===
Public Sub SaveRunSnapshot_Fixed()
    Dim filePath As String

    ' Require that this workbook has been saved at least once
    If Len(ThisWorkbook.path) = 0 Then
        MsgBox "Please save the workbook first (File > Save).", vbExclamation
        Exit Sub
    End If

    ' Optional sanity check: folder must exist
    If Dir(ARCHIVE_DIR, vbDirectory) = "" Then
        MsgBox "Archive folder not found:" & vbCrLf & ARCHIVE_DIR, vbCritical
        Exit Sub
    End If

    ' Timestamped filename (down to seconds for uniqueness)
    filePath = ARCHIVE_DIR & Application.PathSeparator & _
               "HistorianTool_" & Format(Now, "yyyy-mm-dd_HHMMSS") & ".xlsm"

    ' Exact snapshot (includes all sheets, charts, and VBA)
    ThisWorkbook.SaveCopyAs filePath

    MsgBox "Saved snapshot:" & vbCrLf & filePath, vbInformation
End Sub


