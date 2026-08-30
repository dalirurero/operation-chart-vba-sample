Attribute VB_Name = "M01_Main"
Option Explicit

Public Sub ImportReceivingData()
    Dim wb As Workbook
    Dim appState As TApplicationState
    Dim protectionState As TProtectionState
    Dim baseDate As Date
    Dim sendData As Variant, send3Data As Variant
    Dim sendRows As Long, send3Rows As Long
    Dim sendDate As Date, send3Date As Date
    Dim sendStore As String, send3Store As String
    Dim sendCodes As Variant, send3Codes As Variant
    Dim sendStatus As String, send3Status As String
    Dim sendValid As Boolean, send3Valid As Boolean
    Dim sendWarning As Boolean, send3Warning As Boolean
    Dim message As String
    Dim sendNote As String, send3Note As String
    Dim fatalError As Boolean
    Dim updateSend As Boolean, updateSend3 As Boolean
    Dim processStage As String

    On Error GoTo ErrorHandler
    Set wb = ThisWorkbook
    CaptureApplicationState appState

    processStage = "必須シート確認"
    If Not ValidateRequiredWorksheets(wb, message) Then GoTo UserStop
    CaptureProtectionState wb, protectionState
    PrepareProtectionForVbaUpdate wb

    If Not GetProcessingBaseDate(wb, baseDate, message) Then GoTo UserStop

    If Not LoadExternalData(wb.Path & Application.PathSeparator & FILE_SEND, False, _
                               sendData, sendRows, sendDate, sendStore, sendStatus, sendNote) Then
        fatalError = True
        message = sendNote
        GoTo UserStop
    End If
    If sendStatus = STATUS_FATAL Or sendStatus = STATUS_ERROR Then
        fatalError = True
        message = sendNote
        GoTo UserStop
    End If

    If Not LoadExternalData(wb.Path & Application.PathSeparator & FILE_SEND3, True, _
                               send3Data, send3Rows, send3Date, send3Store, send3Status, send3Note) Then
        fatalError = True
        message = send3Note
        GoTo UserStop
    End If
    If send3Status = STATUS_FATAL Or send3Status = STATUS_ERROR Then
        fatalError = True
        message = send3Note
        GoTo UserStop
    End If

    '参照CDとエリアデータは、既存帳票を消す前に検証する。
    If sendRows > 0 And sendStatus <> STATUS_SIDE_ERROR Then
        If Not BuildReferenceCodeList(sendData, sendRows, sendCodes, message) Then fatalError = True: GoTo UserStop
        If Not ValidateAreaData(wb, sendCodes, message) Then fatalError = True: GoTo UserStop
    End If
    If send3Rows > 0 And send3Status <> STATUS_SIDE_ERROR Then
        If Not BuildReferenceCodeList(send3Data, send3Rows, send3Codes, message) Then fatalError = True: GoTo UserStop
        If Not ValidateAreaData(wb, send3Codes, message) Then fatalError = True: GoTo UserStop
    End If

    If sendRows > 0 And sendStatus <> STATUS_SIDE_ERROR Then
        sendValid = EvaluateExternalDate(sendDate, baseDate, False, sendStatus, sendWarning)
        If sendStatus = STATUS_ERROR Then message = "当日データの処理を中止しました。": GoTo UserStop
    End If

    If send3Rows > 0 And send3Status <> STATUS_SIDE_ERROR Then
        send3Valid = EvaluateExternalDate(send3Date, baseDate, True, send3Status, send3Warning)
        If send3Status = STATUS_ERROR Then message = "次回予定データの処理を中止しました。": GoTo UserStop
    End If

    updateSend = (sendStatus <> STATUS_SIDE_ERROR)
    updateSend3 = (send3Status <> STATUS_SIDE_ERROR)

    If updateSend Then
        processStage = "作業編成画面の更新"
        UpdateTargetSide wb, False, sendValid, sendData, sendRows, sendDate, sendCodes
    End If
    If updateSend3 Then
        processStage = "作業編成画面Aの更新"
        UpdateTargetSide wb, True, send3Valid, send3Data, send3Rows, send3Date, send3Codes
    End If

    processStage = "作業編成画面の再計算"
    wb.Worksheets(SH_WORK).Calculate
    wb.Worksheets(SH_WORK_NEXT).Calculate
    message = BuildCompletionMessage(sendStatus, sendRows, sendCodes, sendWarning, sendNote, _
                         send3Status, send3Rows, send3Codes, send3Warning, send3Note)

CleanExit:
    On Error Resume Next
    RestoreProtectionState wb, protectionState
    RestoreApplicationState appState
    wb.Worksheets(SH_WORK).Activate
    On Error GoTo 0

    If Len(message) > 0 Then
        If fatalError Then
            MsgBox message, vbCritical, "ImportReceivingData"
        Else
            MsgBox message, vbInformation, "ImportReceivingData"
        End If
    End If
    Exit Sub

UserStop:
    GoTo CleanExit

ErrorHandler:
    fatalError = True
    message = "ImportReceivingData中にエラーが発生しました。" & vbCrLf & _
              "処理段階：" & processStage & vbCrLf & _
              "エラー " & Err.Number & "：" & Err.Description
    Resume CleanExit
End Sub

Public Sub PrintIndividualInstructions()
    Dim wb As Workbook
    Dim appState As TApplicationState
    Dim protectionState As TProtectionState
    Dim printedPeople As Long
    Dim printedPages As Long
    Dim message As String
    Dim succeeded As Boolean

    On Error GoTo ErrorHandler
    Set wb = ThisWorkbook
    CaptureApplicationState appState
    If Not ValidateRequiredWorksheets(wb, message) Then GoTo CleanExit
    CaptureProtectionState wb, protectionState

    wb.Worksheets(SH_WORK).Calculate
    succeeded = PrintInstructions(wb, printedPeople, printedPages, message)
    If succeeded And printedPeople > 0 Then
        message = "指示書の印刷が完了しました。" & vbCrLf & _
                  "印刷人数：" & printedPeople & "人" & vbCrLf & _
                  "印刷枚数：" & printedPages & "枚"
    End If

CleanExit:
    On Error Resume Next
    RestoreProtectionState wb, protectionState
    RestoreApplicationState appState
    wb.Worksheets(SH_WORK).Activate
    On Error GoTo 0
    If Len(message) > 0 Then MsgBox message, IIf(succeeded, vbInformation, vbExclamation), "指示書印刷"
    Exit Sub

ErrorHandler:
    message = "指示書印刷中にエラーが発生しました。" & vbCrLf & _
              "エラー " & Err.Number & "：" & Err.Description
    succeeded = False
    Resume CleanExit
End Sub

Public Sub PrintWorkAssignmentSheet()
    Dim wb As Workbook
    Dim appState As TApplicationState
    Dim protectionState As TProtectionState
    Dim message As String
    Dim succeeded As Boolean

    On Error GoTo ErrorHandler
    Set wb = ThisWorkbook
    CaptureApplicationState appState
    If Not ValidateRequiredWorksheets(wb, message) Then GoTo CleanExit
    CaptureProtectionState wb, protectionState

    PrintWorkAssignment wb
    succeeded = True

CleanExit:
    On Error Resume Next
    RestoreProtectionState wb, protectionState
    RestoreApplicationState appState
    wb.Worksheets(SH_WORK).Activate
    On Error GoTo 0
    If Len(message) > 0 Then MsgBox message, vbExclamation, "帳票印刷"
    Exit Sub

ErrorHandler:
    message = "作業編成帳票の印刷中にエラーが発生しました。" & vbCrLf & _
              "エラー " & Err.Number & "：" & Err.Description
    Resume CleanExit
End Sub

Private Function BuildCompletionMessage(ByVal sendStatus As String, ByVal sendRows As Long, _
                            ByVal sendCodes As Variant, ByVal sendWarning As Boolean, _
                            ByVal sendNote As String, ByVal send3Status As String, _
                            ByVal send3Rows As Long, ByVal send3Codes As Variant, _
                            ByVal send3Warning As Boolean, ByVal send3Note As String) As String
    Dim text As String

    text = "ImportReceivingDataが完了しました。" & vbCrLf & vbCrLf
    text = text & "【当日分】" & vbCrLf & FormatStatusMessage(sendStatus, sendRows, sendCodes, sendWarning, sendNote) & vbCrLf & vbCrLf
    text = text & "【次回予定分】" & vbCrLf & FormatStatusMessage(send3Status, send3Rows, send3Codes, send3Warning, send3Note)
    BuildCompletionMessage = text
End Function

Private Function FormatStatusMessage(ByVal status As String, ByVal rowCount As Long, _
                       ByVal codes As Variant, ByVal warningContinued As Boolean, _
                       ByVal noteText As String) As String
    Select Case status
        Case STATUS_VALID, STATUS_WARNING
            FormatStatusMessage = "参照CD件数：" & GetArrayElementCount(codes) & "件"
            If warningContinued Then FormatStatusMessage = FormatStatusMessage & vbCrLf & "日付警告後に続行しました。"
        Case STATUS_NO_FILE
            FormatStatusMessage = "外部ファイルがないため、データなしとして処理しました。"
        Case STATUS_NO_DATA
            FormatStatusMessage = "外部ファイルに明細がないため、データなしとして処理しました。"
        Case STATUS_OLD
            FormatStatusMessage = "古いデータのため取り込みませんでした。"
        Case STATUS_SIDE_ERROR
            FormatStatusMessage = noteText
            If Len(FormatStatusMessage) = 0 Then FormatStatusMessage = "対象側でエラーが発生したため、既存データを維持しました。"
        Case Else
            FormatStatusMessage = "データなし"
    End Select
End Function
