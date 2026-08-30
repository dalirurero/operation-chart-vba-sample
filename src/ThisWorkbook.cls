VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    ApplyStartupProtection Me
    Me.Worksheets(SH_MF).Activate

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "ブック起動時の保護設定でエラーが発生しました。" & vbCrLf & _
           "エラー " & Err.Number & "：" & Err.Description, _
           vbExclamation, "起動時処理"
    Resume CleanExit
End Sub
