Attribute VB_Name = "M04_ReferenceCode"
Option Explicit

Public Function BuildReferenceCodeList(ByVal dataValues As Variant, ByVal rowCount As Long, _
                               ByRef uniqueCodes As Variant, ByRef message As String) As Boolean
    Dim dic As Object
    Dim r As Long
    Dim code As String
    Dim invalidRows As String
    Dim invalidCount As Long
    Dim keys As Variant

    Set dic = CreateDictionary(False)

    For r = 1 To rowCount
        code = CreateReferenceCode(dataValues(r, SRC_COL_SHELF1), _
                         dataValues(r, SRC_COL_SHELF2), _
                         dataValues(r, SRC_COL_AREA))
        If Len(code) = 0 Then
            invalidCount = invalidCount + 1
            If invalidCount <= 10 Then
                If Len(invalidRows) > 0 Then invalidRows = invalidRows & ", "
                invalidRows = invalidRows & CStr(r + 1)
            End If
        ElseIf Not dic.Exists(code) Then
            dic.Add code, True
        End If
    Next r

    If invalidCount > 0 Then
        message = "CreateReferenceCodeできない明細があります。" & vbCrLf & _
                  "外部データ行：" & invalidRows
        If invalidCount > 10 Then message = message & " ほか" & (invalidCount - 10) & "件"
        message = message & vbCrLf & "棚1・棚2・エリアを確認してください。"
        Exit Function
    End If

    keys = dic.Keys
    If dic.Count > 1 Then QuickSortStringsAscending keys, LBound(keys), UBound(keys)
    uniqueCodes = keys
    BuildReferenceCodeList = True
End Function

Public Function CreateReferenceCode(ByVal shelf1 As Variant, ByVal shelf2 As Variant, _
                           ByVal areaCode As Variant) As String
    Dim s1 As String, s2 As String, area As String
    Dim numericPart As Long

    s1 = ToText(shelf1)
    s2 = ToText(shelf2)
    area = ToText(areaCode)

    If Len(s1) = 0 Or Len(s2) = 0 Or Len(area) = 0 Then Exit Function
    If Not IsNumeric(s1) Or Not IsNumeric(s2) Then Exit Function

    numericPart = CLng(s1) * 100 + CLng(s2)
    CreateReferenceCode = CStr(numericPart) & area
End Function

Public Function ValidateAreaData(ByVal wb As Workbook, ByVal codes As Variant, _
                                  ByRef message As String) As Boolean
    Dim ws As Worksheet
    Dim dic As Object
    Dim lastRow As Long
    Dim r As Long
    Dim i As Long
    Dim code As String
    Dim missing As String
    Dim missingCount As Long

    Set ws = wb.Worksheets(SH_AREA)
    Set dic = CreateDictionary(False)
    lastRow = GetLastRow(ws, COL_AREA_REF, 2)

    For r = 2 To lastRow
        code = ToText(ws.Cells(r, COL_AREA_REF).Value)
        If Len(code) > 0 Then
            If Not dic.Exists(code) Then dic.Add code, True
        End If
    Next r

    If GetArrayElementCount(codes) = 0 Then
        ValidateAreaData = True
        Exit Function
    End If

    For i = LBound(codes) To UBound(codes)
        code = CStr(codes(i))
        If Not dic.Exists(code) Then
            missingCount = missingCount + 1
            If missingCount <= 10 Then
                If Len(missing) > 0 Then missing = missing & vbCrLf
                missing = missing & code
            End If
        End If
    Next i

    If missingCount > 0 Then
        message = "エリアデータに未登録の参照CDが" & missingCount & "件あります。" & vbCrLf & _
                  missing
        If missingCount > 10 Then message = message & vbCrLf & "ほか" & (missingCount - 10) & "件"
        message = message & vbCrLf & "エリアデータを更新してから再実行してください。"
        Exit Function
    End If

    ValidateAreaData = True
End Function

Private Sub QuickSortStringsAscending(ByRef values As Variant, ByVal first As Long, ByVal last As Long)
    Dim low As Long, high As Long
    Dim pivot As String
    Dim temp As Variant

    low = first
    high = last
    pivot = CStr(values((first + last) \ 2))

    Do While low <= high
        Do While StrComp(CStr(values(low)), pivot, vbTextCompare) < 0
            low = low + 1
        Loop
        Do While StrComp(CStr(values(high)), pivot, vbTextCompare) > 0
            high = high - 1
        Loop
        If low <= high Then
            temp = values(low)
            values(low) = values(high)
            values(high) = temp
            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSortStringsAscending values, first, high
    If low < last Then QuickSortStringsAscending values, low, last
End Sub
