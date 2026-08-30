Attribute VB_Name = "M07_WorkAssignmentPrint"
Option Explicit

Public Sub PrintWorkAssignment(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lastProductRow As Long
    Dim printLastRow As Long

    Set ws = wb.Worksheets(SH_WORK)
    lastProductRow = GetLastRow(ws, COL_REF, ROW_PRODUCT_TEMPLATE)
    If Len(ToText(ws.Cells(lastProductRow, COL_REF).Value)) = 0 Then
        lastProductRow = ROW_PRODUCT_TEMPLATE
    End If

    printLastRow = lastProductRow
    If printLastRow < ROW_WORK_PRINT_MIN Then printLastRow = ROW_WORK_PRINT_MIN

    ws.PageSetup.PrintTitleRows = PRINT_TITLE_WORK
    ws.PageSetup.PrintArea = "$A$1:$M$" & printLastRow
    ws.PrintOut Copies:=1
End Sub
