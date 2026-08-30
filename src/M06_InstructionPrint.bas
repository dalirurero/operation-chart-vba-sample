Attribute VB_Name = "M06_InstructionPrint"
Option Explicit

Public Function PrintInstructions(ByVal wb As Workbook, ByRef printedPeople As Long, _
                               ByRef printedPages As Long, ByRef message As String) As Boolean
    Dim wsWork As Worksheet
    Dim wsInstruction As Worksheet
    Dim workers As Object
    Dim duplicateName As String
    Dim invalidName As String
    Dim unassignedRows As String
    Dim answer As VbMsgBoxResult
    Dim r As Long
    Dim workerName As String
    Dim productCodes As Variant
    Dim productCount As Long
    Dim pages As Long

    Set wsWork = wb.Worksheets(SH_WORK)
    Set wsInstruction = wb.Worksheets(SH_INSTRUCTION)

    If Not ValidateWorkerList(wsWork, workers, duplicateName) Then
        message = "作業者名が重複しています。" & vbCrLf & "氏名：" & duplicateName
        Exit Function
    End If

    If Not ValidateAssignedWorkerNames(wsWork, workers, invalidName) Then
        message = "I列に作業者一覧へ存在しない氏名があります。" & vbCrLf & _
                  "氏名：" & invalidName
        Exit Function
    End If

    If HasUnassignedProducts(wsWork, unassignedRows) Then
        answer = MsgBox("担当作業者が未割当の商品があります。" & vbCrLf & _
                        "対象行：" & unassignedRows & vbCrLf & _
                        "このまま印刷を続行しますか？", _
                        vbExclamation + vbYesNo, "未割当商品の確認")
        If answer <> vbYes Then
            message = "利用者が印刷を中止しました。"
            Exit Function
        End If
    End If

    printedPeople = 0
    printedPages = 0

    For r = ROW_WORKER_START To ROW_WORKER_END
        workerName = NormalizeName(wsWork.Cells(r, COL_WORKER).Value)
        If Len(workerName) > 0 And IsPositiveValue(wsWork.Cells(r, COL_WORKER_QTY).Value) Then
            productCodes = GetWorkerProductCodes(wsWork, workerName, productCount)
            If productCount > 0 Then
                BuildInstructionSheet wsInstruction, wsWork, workerName, productCodes, productCount
                pages = GetPrintedPageCount(wsInstruction)
                wsInstruction.PrintOut Copies:=1
                printedPeople = printedPeople + 1
                printedPages = printedPages + pages
            End If
        End If
    Next r

    If printedPeople = 0 Then
        message = "印刷対象者はいません。"
        PrintInstructions = True
        Exit Function
    End If

    PrintInstructions = True
End Function

Private Function ValidateWorkerList(ByVal ws As Worksheet, ByRef workers As Object, _
                               ByRef duplicateName As String) As Boolean
    Dim r As Long
    Dim nameText As String
    Dim workerOnly As Object

    Set workers = CreateDictionary(True)
    Set workerOnly = CreateDictionary(True)

    nameText = NormalizeName(ws.Cells(5, COL_WORKER).Value)
    If Len(nameText) > 0 Then workers.Add nameText, True
    nameText = NormalizeName(ws.Cells(6, COL_WORKER).Value)
    If Len(nameText) > 0 And Not workers.Exists(nameText) Then workers.Add nameText, True

    For r = ROW_WORKER_START To ROW_WORKER_END
        nameText = NormalizeName(ws.Cells(r, COL_WORKER).Value)
        If Len(nameText) > 0 Then
            If workerOnly.Exists(nameText) Then
                duplicateName = nameText
                Exit Function
            End If
            workerOnly.Add nameText, True
            If Not workers.Exists(nameText) Then workers.Add nameText, True
        End If
    Next r

    ValidateWorkerList = True
End Function

Private Function ValidateAssignedWorkerNames(ByVal ws As Worksheet, ByVal workers As Object, _
                            ByRef invalidName As String) As Boolean
    Dim lastProductRow As Long
    Dim r As Long
    Dim nameText As String

    lastProductRow = GetLastRow(ws, COL_REF, ROW_PRODUCT_TEMPLATE)
    For r = ROW_PRODUCT_TEMPLATE To lastProductRow
        If Len(ToText(ws.Cells(r, COL_REF).Value)) > 0 Then
            nameText = NormalizeName(ws.Cells(r, COL_ASSIGNEE).Value)
            If Len(nameText) > 0 Then
                If Not workers.Exists(nameText) Then
                    invalidName = nameText & "（" & r & "行目）"
                    Exit Function
                End If
            End If
        End If
    Next r

    ValidateAssignedWorkerNames = True
End Function

Private Function HasUnassignedProducts(ByVal ws As Worksheet, ByRef rowList As String) As Boolean
    Dim lastProductRow As Long
    Dim r As Long
    Dim count As Long
    Dim fValue As String, gValue As String

    lastProductRow = GetLastRow(ws, COL_REF, ROW_PRODUCT_TEMPLATE)
    For r = ROW_PRODUCT_TEMPLATE To lastProductRow
        If Len(ToText(ws.Cells(r, COL_REF).Value)) > 0 Then
            fValue = ToText(ws.Cells(r, COL_REPLACE_TIME).Value)
            gValue = ToText(ws.Cells(r, COL_NO_REPLACE_TIME).Value)
            If StrComp(fValue, TEXT_SAME_DAY_REPLENISH, vbTextCompare) <> 0 And _
               StrComp(gValue, TEXT_SAME_DAY_REPLENISH, vbTextCompare) <> 0 And _
               Len(NormalizeName(ws.Cells(r, COL_ASSIGNEE).Value)) = 0 Then
                count = count + 1
                If count <= 15 Then
                    If Len(rowList) > 0 Then rowList = rowList & ", "
                    rowList = rowList & CStr(r)
                End If
            End If
        End If
    Next r

    If count > 15 Then rowList = rowList & " ほか" & (count - 15) & "件"
    HasUnassignedProducts = (count > 0)
End Function

Private Function GetWorkerProductCodes(ByVal ws As Worksheet, ByVal workerName As String, _
                                  ByRef productCount As Long) As Variant
    Dim values() As String
    Dim lastProductRow As Long
    Dim r As Long
    Dim assignedName As String

    productCount = 0
    lastProductRow = GetLastRow(ws, COL_REF, ROW_PRODUCT_TEMPLATE)

    For r = ROW_PRODUCT_TEMPLATE To lastProductRow
        assignedName = NormalizeName(ws.Cells(r, COL_ASSIGNEE).Value)
        If StrComp(assignedName, workerName, vbTextCompare) = 0 And _
           Len(ToText(ws.Cells(r, COL_REF).Value)) > 0 Then
            productCount = productCount + 1
            ReDim Preserve values(1 To productCount)
            values(productCount) = ToText(ws.Cells(r, COL_REF).Value)
        End If
    Next r

    If productCount > 0 Then GetWorkerProductCodes = values
End Function

Private Sub BuildInstructionSheet(ByVal ws As Worksheet, ByVal wsWork As Worksheet, _
                       ByVal workerName As String, ByVal productCodes As Variant, _
                       ByVal productCount As Long)
    Dim requiredLastRow As Long
    Dim i As Long
    Dim r As Long

    ResetInstructionVariableArea ws
    requiredLastRow = GetInstructionLastRow(productCount)
    If requiredLastRow > ROW_INSTRUCTION_FIRST_END Then
        CreateInstructionVariableArea ws, requiredLastRow
    End If

    ws.Range("B1").Value = wsWork.Range("E2").Value
    ws.Range("B1").NumberFormatLocal = DATE_FORMAT_DISPLAY
    ws.Range("C1").Value = workerName
    ws.Range("F1").Formula = "=SUM(D:D)"

    ws.Range("B3:B" & requiredLastRow).ClearContents
    ws.Range("E3:E" & requiredLastRow).ClearContents
    ws.Range("G3:G" & requiredLastRow).ClearContents

    SetInstructionFormulas ws, productCount

    For i = 1 To productCount
        r = ROW_INSTRUCTION_START + i - 1
        ws.Cells(r, 2).Value = CStr(productCodes(i))
    Next i

    ws.Calculate

    ws.PageSetup.PrintTitleRows = PRINT_TITLE_INSTRUCTION
    ws.PageSetup.PrintArea = "$B$1:$G$" & requiredLastRow
End Sub

Private Sub SetInstructionFormulas(ByVal ws As Worksheet, ByVal productCount As Long)
    Dim dataLastRow As Long

    If productCount <= 0 Then Exit Sub

    If Not ws.Range("C3").HasFormula Then
        Err.Raise vbObjectError + 601, "SetInstructionFormulas", _
                  "指示書のテンプレート数式がありません。対象セル：C3"
    End If
    If Not ws.Range("D3").HasFormula Then
        Err.Raise vbObjectError + 602, "SetInstructionFormulas", _
                  "指示書のテンプレート数式がありません。対象セル：D3"
    End If
    If Not ws.Range("F3").HasFormula Then
        Err.Raise vbObjectError + 603, "SetInstructionFormulas", _
                  "指示書のテンプレート数式がありません。対象セル：F3"
    End If

    dataLastRow = ROW_INSTRUCTION_START + productCount - 1
    ws.Range("C3:C" & dataLastRow).FormulaR1C1 = ws.Range("C3").FormulaR1C1
    ws.Range("D3:D" & dataLastRow).FormulaR1C1 = ws.Range("D3").FormulaR1C1
    ws.Range("F3:F" & dataLastRow).FormulaR1C1 = ws.Range("F3").FormulaR1C1
End Sub

Private Sub ResetInstructionVariableArea(ByVal ws As Worksheet)
    Dim lastUsedRow As Long

    lastUsedRow = GetLastUsedRowInColumnRange(ws, 2, 7, ROW_INSTRUCTION_EXTEND_START)
    If lastUsedRow >= ROW_INSTRUCTION_EXTEND_START Then
        ws.Range(ws.Cells(ROW_INSTRUCTION_EXTEND_START, 2), ws.Cells(lastUsedRow, 7)).Clear
    End If
    ws.PageSetup.PrintArea = "$B$1:$G$35"
End Sub

Private Sub CreateInstructionVariableArea(ByVal ws As Worksheet, ByVal requiredLastRow As Long)
    Dim blockStart As Long
    Dim sourceStart As Long
    Dim offset As Long

    For blockStart = ROW_INSTRUCTION_EXTEND_START To requiredLastRow Step 3
        sourceStart = blockStart - 3
        ws.Range(ws.Cells(sourceStart, 2), ws.Cells(sourceStart + 2, 7)).Copy _
            Destination:=ws.Range(ws.Cells(blockStart, 2), ws.Cells(blockStart + 2, 7))
        For offset = 0 To 2
            ws.Rows(blockStart + offset).RowHeight = ws.Rows(sourceStart + offset).RowHeight
        Next offset
    Next blockStart
    Application.CutCopyMode = False
End Sub

Private Function GetInstructionLastRow(ByVal productCount As Long) As Long
    Dim additionalProducts As Long
    Dim additionalBlocks As Long

    If productCount <= 33 Then
        GetInstructionLastRow = ROW_INSTRUCTION_FIRST_END
    Else
        additionalProducts = productCount - 33
        additionalBlocks = (additionalProducts + 2) \ 3
        GetInstructionLastRow = ROW_INSTRUCTION_FIRST_END + additionalBlocks * 3
    End If
End Function

Private Function GetPrintedPageCount(ByVal ws As Worksheet) As Long
    Dim previousSheet As Object
    Dim pageCount As Variant

    On Error GoTo Fallback
    Set previousSheet = ActiveSheet
    ws.Activate
    pageCount = ExecuteExcel4Macro("GET.DOCUMENT(50)")
    If Not previousSheet Is Nothing Then previousSheet.Activate

    If IsNumeric(pageCount) And CLng(pageCount) > 0 Then
        GetPrintedPageCount = CLng(pageCount)
    Else
        GoTo Fallback
    End If
    Exit Function

Fallback:
    On Error Resume Next
    If Not previousSheet Is Nothing Then previousSheet.Activate
    GetPrintedPageCount = 1
    On Error GoTo 0
End Function
