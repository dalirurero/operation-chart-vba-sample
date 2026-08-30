Attribute VB_Name = "M08_Protection"
Option Explicit

Public Type TProtectionState
    SheetNames As Variant
    SheetProtected As Variant
    SheetEnableSelection As Variant
    WorkbookStructureProtected As Boolean
End Type

Public Sub CaptureProtectionState(ByVal wb As Workbook, ByRef state As TProtectionState)
    Dim names As Variant
    Dim protectedValues() As Boolean
    Dim selectionValues() As Long
    Dim i As Long
    Dim ws As Worksheet

    names = Array(SH_MF, SH_WORK, SH_WORK_NEXT, SH_INSTRUCTION, SH_AREA)
    ReDim protectedValues(LBound(names) To UBound(names))
    ReDim selectionValues(LBound(names) To UBound(names))

    For i = LBound(names) To UBound(names)
        Set ws = wb.Worksheets(CStr(names(i)))
        protectedValues(i) = ws.ProtectContents
        selectionValues(i) = ws.EnableSelection
    Next i

    state.SheetNames = names
    state.SheetProtected = protectedValues
    state.SheetEnableSelection = selectionValues
    state.WorkbookStructureProtected = wb.ProtectStructure
End Sub

Public Sub ApplyStartupProtection(ByVal wb As Workbook)
    Dim names As Variant
    Dim i As Long
    Dim ws As Worksheet

    names = Array(SH_MF, SH_WORK, SH_WORK_NEXT, SH_INSTRUCTION, SH_AREA)
    For i = LBound(names) To UBound(names)
        Set ws = wb.Worksheets(CStr(names(i)))
        ProtectWorksheetForUI ws
    Next i

    If Not wb.ProtectStructure Then
        wb.Protect Structure:=True, Windows:=False
    End If
End Sub

Public Sub ProtectWorksheetForUI(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, _
               AllowFormattingCells:=False, AllowFormattingColumns:=False, _
               AllowFormattingRows:=False, AllowInsertingColumns:=False, _
               AllowInsertingRows:=False, AllowInsertingHyperlinks:=False, _
               AllowDeletingColumns:=False, AllowDeletingRows:=False, _
               AllowSorting:=False, AllowFiltering:=False, _
               AllowUsingPivotTables:=False
    ws.EnableSelection = xlUnlockedCells
End Sub


Public Sub PrepareProtectionForVbaUpdate(ByVal wb As Workbook)
    Dim names As Variant
    Dim i As Long
    Dim ws As Worksheet

    names = Array(SH_MF, SH_WORK, SH_WORK_NEXT, SH_INSTRUCTION, SH_AREA)
    For i = LBound(names) To UBound(names)
        Set ws = wb.Worksheets(CStr(names(i)))
        If ws.ProtectContents Then ProtectWorksheetForUI ws
    Next i
End Sub

Public Function TemporarilyUnprotectWorksheet(ByVal ws As Worksheet) As Boolean
    TemporarilyUnprotectWorksheet = ws.ProtectContents
    If TemporarilyUnprotectWorksheet Then ws.Unprotect
End Function

Public Sub RestoreTemporaryWorksheetProtection(ByVal ws As Worksheet, ByVal wasProtected As Boolean)
    If wasProtected Then ProtectWorksheetForUI ws
End Sub

Public Sub RestoreProtectionState(ByVal wb As Workbook, ByRef state As TProtectionState)
    Dim i As Long
    Dim ws As Worksheet
    Dim names As Variant
    Dim protectedValues As Variant
    Dim selectionValues As Variant

    On Error Resume Next
    names = state.SheetNames
    protectedValues = state.SheetProtected
    selectionValues = state.SheetEnableSelection

    For i = LBound(names) To UBound(names)
        Set ws = wb.Worksheets(CStr(names(i)))
        ws.Unprotect
        If CBool(protectedValues(i)) Then
            ProtectWorksheetForUI ws
        Else
            ws.EnableSelection = CLng(selectionValues(i))
        End If
    Next i

    If state.WorkbookStructureProtected Then
        If Not wb.ProtectStructure Then wb.Protect Structure:=True, Windows:=False
    Else
        If wb.ProtectStructure Then wb.Unprotect
    End If
    On Error GoTo 0
End Sub
