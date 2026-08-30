Attribute VB_Name = "M90_Common"
Option Explicit

Public Type TApplicationState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As XlCalculation
    StatusBar As Variant
End Type

Public Sub CaptureApplicationState(ByRef state As TApplicationState)
    With Application
        state.ScreenUpdating = .ScreenUpdating
        state.EnableEvents = .EnableEvents
        state.DisplayAlerts = .DisplayAlerts
        state.Calculation = .Calculation
        state.StatusBar = .StatusBar

        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .StatusBar = False
    End With
End Sub

Public Sub RestoreApplicationState(ByRef state As TApplicationState)
    On Error Resume Next
    With Application
        .ScreenUpdating = state.ScreenUpdating
        .EnableEvents = state.EnableEvents
        .DisplayAlerts = state.DisplayAlerts
        .Calculation = state.Calculation
        .StatusBar = state.StatusBar
    End With
    On Error GoTo 0
End Sub

Public Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Public Function ValidateRequiredWorksheets(ByVal wb As Workbook, ByRef message As String) As Boolean
    Dim names As Variant
    Dim i As Long

    names = Array(SH_MF, SH_WORK, SH_WORK_NEXT, SH_INSTRUCTION, SH_AREA, SH_IMPORT, SH_IMPORT_NEXT)
    For i = LBound(names) To UBound(names)
        If Not WorksheetExists(wb, CStr(names(i))) Then
            message = "必須シートが見つかりません。" & vbCrLf & "シート名：" & CStr(names(i))
            Exit Function
        End If
    Next i
    ValidateRequiredWorksheets = True
End Function

Public Function GetLastRow(ByVal ws As Worksheet, ByVal columnNumber As Long, Optional ByVal minimumRow As Long = 1) As Long
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row
    If r < minimumRow Then r = minimumRow
    GetLastRow = r
End Function

Public Function GetLastUsedRowInColumnRange(ByVal ws As Worksheet, ByVal firstColumn As Long, _
                                      ByVal lastColumn As Long, ByVal minimumRow As Long) As Long
    Dim found As Range
    Dim target As Range

    Set target = ws.Range(ws.Cells(minimumRow, firstColumn), ws.Cells(ws.Rows.Count, lastColumn))
    On Error Resume Next
    Set found = target.Find(What:="*", After:=target.Cells(1, 1), LookIn:=xlFormulas, _
                            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                            MatchCase:=False)
    On Error GoTo 0

    If found Is Nothing Then
        GetLastUsedRowInColumnRange = minimumRow
    Else
        GetLastUsedRowInColumnRange = found.Row
    End If
End Function

Public Function ToText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        ToText = vbNullString
    Else
        ToText = Trim$(CStr(value))
    End If
End Function

Public Function IsIntegerValue(ByVal value As Variant) As Boolean
    If IsError(value) Or IsDate(value) Or Not IsNumeric(value) Then Exit Function
    IsIntegerValue = (CDbl(value) = Fix(CDbl(value)))
End Function

Public Function GetArrayElementCount(ByVal values As Variant) As Long
    On Error GoTo EmptyArray
    GetArrayElementCount = UBound(values) - LBound(values) + 1
    Exit Function
EmptyArray:
    GetArrayElementCount = 0
End Function

Public Function CreateDictionary(Optional ByVal textCompare As Boolean = True) As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    If textCompare Then dic.CompareMode = vbTextCompare
    Set CreateDictionary = dic
End Function

Public Function GetColumnLetter(ByVal columnNumber As Long) As String
    GetColumnLetter = Split(Cells(1, columnNumber).Address(False, False), "1")(0)
End Function

Public Function IsPositiveValue(ByVal value As Variant) As Boolean
    If IsError(value) Or Not IsNumeric(value) Then Exit Function
    IsPositiveValue = (CDbl(value) > 0)
End Function

Public Function NormalizeName(ByVal value As Variant) As String
    NormalizeName = Trim$(ToText(value))
End Function
