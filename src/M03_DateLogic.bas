Attribute VB_Name = "M03_DateLogic"
Option Explicit

Public Function GetProcessingBaseDate(ByVal wb As Workbook, ByRef baseDate As Date, _
                                ByRef message As String) As Boolean
    Dim ws As Worksheet
    Dim offsetValue As Variant
    Dim answer As VbMsgBoxResult

    Set ws = wb.Worksheets(SH_MF)
    offsetValue = ws.Range(CELL_DATE_OFFSET).Value

    If Len(Trim$(CStr(offsetValue))) > 0 Then
        If Not IsIntegerValue(offsetValue) Then
            message = "日付調整値には整数を入力してください。" & vbCrLf & _
                      "空白または0は当日、-1は前日、1は翌日です。"
            Exit Function
        End If
    End If

    If Not IsDate(ws.Range(CELL_BASE_DATE).Value) Then
        message = "処理基準日が日付として認識できません。" & vbCrLf & _
                  "MFシートの" & CELL_BASE_DATE & "を確認してください。"
        Exit Function
    End If

    baseDate = DateValue(ws.Range(CELL_BASE_DATE).Value)

    If Len(Trim$(CStr(offsetValue))) > 0 And CDbl(offsetValue) <> 0 Then
        answer = MsgBox("処理基準日は " & Format$(baseDate, "yyyy/m/d") & " です。" & vbCrLf & _
                        "本日（" & Format$(Date, "yyyy/m/d") & "）とは異なります。" & vbCrLf & _
                        "この日付を基準に処理を続行しますか？", _
                        vbQuestion + vbYesNo, "処理基準日の確認")
        If answer <> vbYes Then
            message = "利用者が処理を中止しました。"
            Exit Function
        End If
    End If

    GetProcessingBaseDate = True
End Function

Public Function GetNextDeliveryDate(ByVal baseDate As Date) As Date
    Select Case Weekday(baseDate, vbMonday)
        Case 1: GetNextDeliveryDate = DateAdd("d", 1, baseDate) '月→火
        Case 2: GetNextDeliveryDate = DateAdd("d", 2, baseDate) '火→木
        Case 3: GetNextDeliveryDate = DateAdd("d", 1, baseDate) '水→木
        Case 4, 5, 6: GetNextDeliveryDate = DateAdd("d", 1, baseDate)
        Case 7: GetNextDeliveryDate = DateAdd("d", 1, baseDate) '日→月
    End Select
End Function

Public Function EvaluateExternalDate(ByVal externalDate As Date, ByVal baseDate As Date, _
                              ByVal isNext As Boolean, ByRef status As String, _
                              ByRef warningContinued As Boolean) As Boolean
    Dim expectedDate As Date
    Dim answer As VbMsgBoxResult
    Dim titleText As String
    Dim message As String

    warningContinued = False

    If isNext Then
        expectedDate = GetNextDeliveryDate(baseDate)
        titleText = FILE_SEND3 & "の日付確認"

        If externalDate = expectedDate Then
            status = STATUS_VALID
            EvaluateExternalDate = True
        ElseIf externalDate <= baseDate Then
            status = STATUS_OLD
            EvaluateExternalDate = False
        Else
            message = FILE_SEND3 & "の納品予定日が想定日と異なります。" & vbCrLf & _
                      "外部ファイルの日付：" & Format$(externalDate, "yyyy/m/d") & vbCrLf & _
                      "想定する次回納品日：" & Format$(expectedDate, "yyyy/m/d") & vbCrLf & _
                      "外部ファイルの日付を使用して続行しますか？"
            answer = MsgBox(message, vbExclamation + vbYesNo, titleText)
            If answer = vbYes Then
                status = STATUS_WARNING
                warningContinued = True
                EvaluateExternalDate = True
            Else
                status = STATUS_ERROR
                EvaluateExternalDate = False
            End If
        End If
    Else
        titleText = FILE_SEND & "の日付確認"
        If externalDate = baseDate Then
            status = STATUS_VALID
            EvaluateExternalDate = True
        ElseIf externalDate < baseDate Then
            status = STATUS_OLD
            EvaluateExternalDate = False
        Else
            message = FILE_SEND & "の納品日が処理基準日より後です。" & vbCrLf & _
                      "外部ファイルの日付：" & Format$(externalDate, "yyyy/m/d") & vbCrLf & _
                      "処理基準日：" & Format$(baseDate, "yyyy/m/d") & vbCrLf & _
                      "外部ファイルの日付を使用して続行しますか？"
            answer = MsgBox(message, vbExclamation + vbYesNo, titleText)
            If answer = vbYes Then
                status = STATUS_WARNING
                warningContinued = True
                EvaluateExternalDate = True
            Else
                status = STATUS_ERROR
                EvaluateExternalDate = False
            End If
        End If
    End If
End Function

Public Function GetDeliveryCountInHalfMonth(ByVal targetDate As Date) As Long
    Dim startDate As Date
    Dim d As Date
    Dim targetGroup As Long
    Dim weekdayNo As Long
    Dim count As Long

    weekdayNo = Weekday(targetDate, vbMonday)
    If weekdayNo = 3 Then
        GetDeliveryCountInHalfMonth = 0
        Exit Function
    End If

    If Day(targetDate) <= 15 Then
        startDate = DateSerial(Year(targetDate), Month(targetDate), 1)
    Else
        startDate = DateSerial(Year(targetDate), Month(targetDate), 16)
    End If

    '1=菓子（月・木・土）、2=食品（火・金・日）
    If weekdayNo = 1 Or weekdayNo = 4 Or weekdayNo = 6 Then
        targetGroup = 1
    Else
        targetGroup = 2
    End If

    For d = startDate To targetDate
        weekdayNo = Weekday(d, vbMonday)
        If targetGroup = 1 Then
            If weekdayNo = 1 Or weekdayNo = 4 Or weekdayNo = 6 Then count = count + 1
        Else
            If weekdayNo = 2 Or weekdayNo = 5 Or weekdayNo = 7 Then count = count + 1
        End If
    Next d

    GetDeliveryCountInHalfMonth = count
End Function

Public Function GetCoefficientColumn(ByVal deliveryCount As Long) As Long
    Select Case deliveryCount
        Case Is <= 1: GetCoefficientColumn = COL_AREA_COEFF_1
        Case 2: GetCoefficientColumn = COL_AREA_COEFF_2
        Case 3: GetCoefficientColumn = COL_AREA_COEFF_3
        Case 4: GetCoefficientColumn = COL_AREA_COEFF_4
        Case Else: GetCoefficientColumn = COL_AREA_COEFF_5
    End Select
End Function
