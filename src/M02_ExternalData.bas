Attribute VB_Name = "M02_ExternalData"
Option Explicit

Public Function LoadExternalData(ByVal filePath As String, ByVal isNext As Boolean, _
                                    ByRef dataValues As Variant, ByRef rowCount As Long, _
                                    ByRef externalDate As Date, ByRef storeCode As String, _
                                    ByRef status As String, ByRef message As String) As Boolean
    Dim extWb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim dateDic As Object
    Dim r As Long
    Dim rawDate As Variant
    Dim parsedDate As Date
    Dim dateKeys As Variant

    status = vbNullString
    rowCount = 0

    If Len(Dir$(filePath)) = 0 Then
        status = STATUS_NO_FILE
        LoadExternalData = True
        Exit Function
    End If

    On Error GoTo ErrorHandler
    Set extWb = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)

    If Not WorksheetExists(extWb, EXTERNAL_SHEET) Then
        status = STATUS_SIDE_ERROR
        message = filePath & " にシート「" & EXTERNAL_SHEET & "」がありません。対象側は更新しません。"
        LoadExternalData = True
        GoTo CleanExit
    End If

    Set ws = extWb.Worksheets(EXTERNAL_SHEET)
    lastColumn = IIf(isNext, SRC_COL_SEND3_LAST, SRC_COL_SEND_LAST)
    lastRow = ws.Cells(ws.Rows.Count, SRC_COL_STORE).End(xlUp).Row

    If lastRow < ROW_EXTERNAL_DATA Then
        status = STATUS_NO_DATA
        LoadExternalData = True
        GoTo CleanExit
    End If

    dataValues = ws.Range(ws.Cells(ROW_EXTERNAL_DATA, 1), ws.Cells(lastRow, lastColumn)).Value2
    rowCount = UBound(dataValues, 1)
    Set dateDic = CreateDictionary(False)

    For r = 1 To rowCount
        If Len(ToText(dataValues(r, SRC_COL_STORE))) = 0 Then
            message = filePath & " の店舗番号が空白です。外部データ行：" & (r + 1)
            status = STATUS_SIDE_ERROR
            LoadExternalData = True
            GoTo CleanExit
        End If

        rawDate = dataValues(r, SRC_COL_DATE)
        If Not TryConvertExternalDate(rawDate, parsedDate) Then
            message = filePath & " の納品日が不正です。外部データ行：" & (r + 1)
            status = STATUS_SIDE_ERROR
            LoadExternalData = True
            GoTo CleanExit
        End If
        If Not dateDic.Exists(CLng(parsedDate)) Then dateDic.Add CLng(parsedDate), True
    Next r

    If dateDic.Count <> 1 Then
        message = filePath & " に複数の納品日が含まれています。処理を中止します。"
        status = STATUS_FATAL
        LoadExternalData = True
        GoTo CleanExit
    End If

    dateKeys = dateDic.Keys
    externalDate = CDate(dateKeys(0))
    storeCode = ToText(dataValues(1, SRC_COL_STORE))
    status = STATUS_VALID
    LoadExternalData = True

CleanExit:
    On Error Resume Next
    If Not extWb Is Nothing Then extWb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

ErrorHandler:
    message = filePath & " の読込み中にエラーが発生しました。" & vbCrLf & _
              "エラー " & Err.Number & "：" & Err.Description
    status = STATUS_ERROR
    Resume CleanExit
End Function

Public Function TryConvertExternalDate(ByVal rawValue As Variant, ByRef resultDate As Date) As Boolean
    Dim s As String
    Dim y As Long, m As Long, d As Long

    On Error GoTo InvalidDate

    If IsDate(rawValue) Then
        resultDate = DateValue(CDate(rawValue))
        TryConvertExternalDate = True
        Exit Function
    End If

    s = Replace$(Replace$(Trim$(CStr(rawValue)), "/", vbNullString), "-", vbNullString)
    If Len(s) <> 8 Or Not IsNumeric(s) Then Exit Function

    y = CLng(Left$(s, 4))
    m = CLng(Mid$(s, 5, 2))
    d = CLng(Right$(s, 2))
    resultDate = DateSerial(y, m, d)

    If Year(resultDate) <> y Or Month(resultDate) <> m Or Day(resultDate) <> d Then Exit Function
    TryConvertExternalDate = True
    Exit Function

InvalidDate:
    TryConvertExternalDate = False
End Function

Public Function IsTextSourceColumn(ByVal sourceColumn As Long, ByVal isNext As Boolean) As Boolean
    If isNext Then
        Select Case sourceColumn
            Case 5, 6, 7, 8, 12, 13
                IsTextSourceColumn = True
        End Select
    Else
        Select Case sourceColumn
            Case 5, 6, 7, 8, 12, 13, 14
                IsTextSourceColumn = True
        End Select
    End If
End Function
