Attribute VB_Name = "M05_ReportUpdate"
Option Explicit

Public Sub UpdateTargetSide(ByVal wb As Workbook, ByVal isNext As Boolean, _
                      ByVal shouldImport As Boolean, ByVal dataValues As Variant, _
                      ByVal rowCount As Long, ByVal externalDate As Date, _
                      ByVal uniqueCodes As Variant)
    Dim wsImport As Worksheet
    Dim wsWork As Worksheet

    If isNext Then
        Set wsImport = wb.Worksheets(SH_IMPORT_NEXT)
        Set wsWork = wb.Worksheets(SH_WORK_NEXT)
    Else
        Set wsImport = wb.Worksheets(SH_IMPORT)
        Set wsWork = wb.Worksheets(SH_WORK)
    End If

    If shouldImport Then
        UpdateReceivingDetails wsImport, dataValues, rowCount, isNext
        UpdateWorkAssignmentSheet wsWork, uniqueCodes, externalDate, isNext
    Else
        ClearReceivingDetails wsImport, isNext
        ResetWorkAssignmentProductArea wsWork, isNext
    End If
End Sub

Private Sub UpdateReceivingDetails(ByVal ws As Worksheet, ByVal dataValues As Variant, _
                         ByVal rowCount As Long, ByVal isNext As Boolean)
    Dim sourceLastColumn As Long
    Dim destinationLastColumn As Long
    Dim outputValues() As Variant
    Dim r As Long, c As Long
    Dim code As String
    Dim parsedDate As Date

    sourceLastColumn = IIf(isNext, SRC_COL_SEND3_LAST, SRC_COL_SEND_LAST)
    destinationLastColumn = sourceLastColumn + 1

    ClearReceivingDetails ws, isNext
    ReDim outputValues(1 To rowCount, 1 To destinationLastColumn)

    For r = 1 To rowCount
        code = CreateReferenceCode(dataValues(r, SRC_COL_SHELF1), _
                         dataValues(r, SRC_COL_SHELF2), _
                         dataValues(r, SRC_COL_AREA))
        outputValues(r, 1) = code

        For c = 1 To sourceLastColumn
            If c = SRC_COL_DATE Then
                If TryConvertExternalDate(dataValues(r, c), parsedDate) Then
                    outputValues(r, c + 1) = parsedDate
                Else
                    outputValues(r, c + 1) = dataValues(r, c)
                End If
            ElseIf IsTextSourceColumn(c, isNext) Then
                outputValues(r, c + 1) = ToText(dataValues(r, c))
            Else
                outputValues(r, c + 1) = dataValues(r, c)
            End If
        Next c
    Next r

    With ws.Range(ws.Cells(ROW_IMPORT_DATA, 1), ws.Cells(ROW_IMPORT_DATA + rowCount - 1, destinationLastColumn))
        .value = outputValues
    End With

    ws.Range(ws.Cells(ROW_IMPORT_DATA, COL_IMPORT_DATE), _
             ws.Cells(ROW_IMPORT_DATA + rowCount - 1, COL_IMPORT_DATE)).NumberFormatLocal = DATE_FORMAT_DISPLAY

    '前ゼロ保持対象の転記先列を文字列形式にする。
    For c = 1 To sourceLastColumn
        If IsTextSourceColumn(c, isNext) Then
            ws.Range(ws.Cells(ROW_IMPORT_DATA, c + 1), _
                     ws.Cells(ROW_IMPORT_DATA + rowCount - 1, c + 1)).NumberFormat = "@"
        End If
    Next c
End Sub

Private Sub ClearReceivingDetails(ByVal ws As Worksheet, ByVal isNext As Boolean)
    Dim lastRow As Long
    Dim lastColumn As Long

    lastColumn = IIf(isNext, SRC_COL_SEND3_LAST + 1, SRC_COL_SEND_LAST + 1)
    lastRow = GetLastUsedRowInColumnRange(ws, 1, lastColumn, ROW_IMPORT_DATA)
    If lastRow >= ROW_IMPORT_DATA Then
        ws.Range(ws.Cells(ROW_IMPORT_DATA, 1), ws.Cells(lastRow, lastColumn)).ClearContents
    End If
End Sub

Private Sub UpdateWorkAssignmentSheet(ByVal ws As Worksheet, ByVal uniqueCodes As Variant, _
                            ByVal targetDate As Date, ByVal isNext As Boolean)
    Dim itemCount As Long
    Dim lastProductRow As Long
    Dim i As Long
    Dim deliveryCount As Long
    Dim coefficientColumn As Long

    ResetWorkAssignmentProductArea ws, isNext
    itemCount = GetArrayElementCount(uniqueCodes)

    If itemCount = 0 Then
        Exit Sub
    End If

    lastProductRow = ROW_PRODUCT_TEMPLATE + itemCount - 1
    ExpandProductTemplate ws, lastProductRow

    For i = 0 To itemCount - 1
        ws.Cells(ROW_PRODUCT_TEMPLATE + i, COL_REF).value = CStr(uniqueCodes(LBound(uniqueCodes) + i))
    Next i

    deliveryCount = GetDeliveryCountInHalfMonth(targetDate)
    coefficientColumn = GetCoefficientColumn(deliveryCount)
    SetProductFormulas ws, lastProductRow, isNext, deliveryCount, coefficientColumn
    SetWorkerValidation ws, lastProductRow
End Sub

Public Sub ResetWorkAssignmentProductArea(ByVal ws As Worksheet, ByVal isNext As Boolean)
    Dim lastUsedRow As Long
    Dim minimumEndRow As Long
    Dim target As Range
    Dim wasProtected As Boolean

    minimumEndRow = IIf(isNext, ROW_WORK_NEXT_DEFAULT_END, ROW_WORK_DEFAULT_END)
    lastUsedRow = GetLastUsedRowInColumnRange(ws, COL_REF, COL_ASSIGNEE, ROW_PRODUCT_TEMPLATE)
    If lastUsedRow < minimumEndRow Then lastUsedRow = minimumEndRow

    wasProtected = TemporarilyUnprotectWorksheet(ws)
    On Error GoTo ExitHandler

    '5行目はテンプレートとして残す。
    ws.Cells(ROW_PRODUCT_TEMPLATE, COL_REF).ClearContents
    ws.Cells(ROW_PRODUCT_TEMPLATE, COL_ASSIGNEE).ClearContents

    If lastUsedRow >= ROW_PRODUCT_TEMPLATE + 1 Then
        Set target = ws.Range(ws.Cells(ROW_PRODUCT_TEMPLATE + 1, COL_REF), _
                              ws.Cells(lastUsedRow, COL_ASSIGNEE))
        target.ClearContents
        DeleteValidationSafely ws.Range(ws.Cells(ROW_PRODUCT_TEMPLATE + 1, COL_ASSIGNEE), _
                                  ws.Cells(lastUsedRow, COL_ASSIGNEE))
        target.ClearFormats
    End If

ExitHandler:
    RestoreTemporaryWorksheetProtection ws, wasProtected
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description

    'テンプレート行の入力規則を要件の範囲へ統一する。
    SetWorkerValidation ws, ROW_PRODUCT_TEMPLATE
End Sub

Private Sub ExpandProductTemplate(ByVal ws As Worksheet, ByVal lastProductRow As Long)
    Dim r As Long

    If lastProductRow <= ROW_PRODUCT_TEMPLATE Then Exit Sub

    For r = ROW_PRODUCT_TEMPLATE + 1 To lastProductRow
        ws.Range(ws.Cells(ROW_PRODUCT_TEMPLATE, COL_REF), _
                 ws.Cells(ROW_PRODUCT_TEMPLATE, COL_ASSIGNEE)).Copy _
                 Destination:=ws.Range(ws.Cells(r, COL_REF), ws.Cells(r, COL_ASSIGNEE))
        ws.Rows(r).RowHeight = ws.Rows(ROW_PRODUCT_TEMPLATE).RowHeight
    Next r
    Application.CutCopyMode = False
End Sub

Private Sub SetProductFormulas(ByVal ws As Worksheet, ByVal lastProductRow As Long, _
                         ByVal isNext As Boolean, ByVal deliveryCount As Long, _
                         ByVal coefficientColumn As Long)
    Dim r As Long
    Dim importSheetName As String
    Dim coefficientIndex As Long

    importSheetName = IIf(isNext, SH_IMPORT_NEXT, SH_IMPORT)
    coefficientIndex = coefficientColumn

    For r = ROW_PRODUCT_TEMPLATE To lastProductRow
        ws.Cells(r, COL_AREA_NAME).Formula = "=VLOOKUP(A" & r & ",'" & SH_AREA & "'!$A:$G,7,FALSE)"
        ws.Cells(r, COL_TOTAL_QTY).Formula = "=SUMIF('" & importSheetName & "'!A:A,A" & r & ",'" & importSheetName & "'!L:L)"

        If isNext Then
            If deliveryCount <= 1 Then
                ws.Cells(r, COL_REPLACE_QTY).Formula = "=C" & r
                ws.Cells(r, COL_NO_REPLACE_QTY).Formula = "=0"
            Else
                ws.Cells(r, COL_REPLACE_QTY).Formula = "=0"
                ws.Cells(r, COL_NO_REPLACE_QTY).Formula = "=C" & r
            End If
        Else
            ws.Cells(r, COL_REPLACE_QTY).Formula = "=SUMIFS('" & importSheetName & "'!L:L,'" & _
                                                        importSheetName & "'!A:A,A" & r & ",'" & _
                                                        importSheetName & "'!O:O,"""")"
            ws.Cells(r, COL_NO_REPLACE_QTY).Formula = "=C" & r & "-D" & r
        End If

        ws.Cells(r, COL_REPLACE_TIME).Formula = "=VLOOKUP(A" & r & ",'" & SH_AREA & "'!A:L,8,FALSE)"
        ws.Cells(r, COL_NO_REPLACE_TIME).Formula = "=VLOOKUP(A" & r & ",'" & SH_AREA & "'!A:L," & coefficientIndex & ",FALSE)"
        ws.Cells(r, COL_PLAN_TIME).Formula = "=IFERROR(ROUNDUP(D" & r & "*F" & r & "+E" & r & "*G" & r & ",0),0)"
    Next r
End Sub

Private Sub SetWorkerValidation(ByVal ws As Worksheet, ByVal lastProductRow As Long)
    Dim templateCell As Range
    Dim copyTarget As Range
    Dim wasProtected As Boolean

    Set templateCell = ws.Cells(ROW_PRODUCT_TEMPLATE, COL_ASSIGNEE)

    wasProtected = TemporarilyUnprotectWorksheet(ws)
    On Error GoTo ExitHandler

    On Error Resume Next
    templateCell.Validation.Delete
    On Error GoTo ExitHandler

    templateCell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                                Operator:=xlBetween, Formula1:=VALIDATION_WORKERS
    templateCell.Validation.IgnoreBlank = True
    templateCell.Validation.InCellDropdown = True
    templateCell.Validation.ShowError = True

    If lastProductRow > ROW_PRODUCT_TEMPLATE Then
        Set copyTarget = ws.Range(ws.Cells(ROW_PRODUCT_TEMPLATE + 1, COL_ASSIGNEE), _
                                  ws.Cells(lastProductRow, COL_ASSIGNEE))
        templateCell.Copy
        copyTarget.PasteSpecial Paste:=xlPasteValidation
        Application.CutCopyMode = False
    End If

ExitHandler:
    Application.CutCopyMode = False
    RestoreTemporaryWorksheetProtection ws, wasProtected
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub DeleteValidationSafely(ByVal target As Range)
    Dim cell As Range

    For Each cell In target.Cells
        On Error Resume Next
        cell.Validation.Delete
        On Error GoTo 0
    Next cell
End Sub
