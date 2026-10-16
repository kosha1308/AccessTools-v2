VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATExcelFieldReader"
Option Explicit

Public Function Read(ByVal ws As Worksheet, ByVal spec As ATExcelFieldSpec) As ATExcelFieldValue
    Dim data As Variant, r As Long, c As Long, foundRow As Long, foundCol As Long
    Dim key As String, area As Range, target As Range, targetRow As Long, targetCol As Long
    Dim id As String, sheetName As String, anchorRef As String, cellRef As String
    Dim found As Boolean, raw As Variant, result As ATEvalResult, value As New ATExcelFieldValue, code As String
    On Error GoTo Failed
    If spec Is Nothing Then Err.Raise vbObjectError + 2810, , "EXCEL_FIELD_SPEC_REQUIRED"
    id = spec.FieldId: spec.Validate
    If ws Is Nothing Then Err.Raise vbObjectError + 2810, , "EXCEL_FIELD_SHEET_REQUIRED"
    sheetName = ws.Name: key = AT_ExcelHeader(spec.AnchorText)
    data = AT_ExcelBlock(ws.Range(ws.Cells(spec.FirstRow, spec.FirstColumn), ws.Cells(spec.LastRow, spec.LastColumn)))
    For r = 1 To UBound(data, 1)
        For c = 1 To UBound(data, 2)
            If Not IsError(data(r, c)) Then
                If Not AT_ExcelBlank(data(r, c)) Then
                    If AT_ExcelHeader(AT_ExcelText(data(r, c))) = key Then
                        If foundRow > 0 Then Err.Raise vbObjectError + 2810, , "EXCEL_ANCHOR_AMBIGUOUS"
                        foundRow = r + spec.FirstRow - 1: foundCol = c + spec.FirstColumn - 1
                    End If
                End If
            End If
        Next c
    Next r
    If foundRow = 0 Then
        Set result = AT_RuntimeOK(AT_RuntimeState("TEXT", atValueMissing))
        GoTo BuildValue
    End If
    Set area = ws.Cells(foundRow, foundCol).MergeArea
    If area.Row < spec.FirstRow Or area.Row + area.Rows.Count - 1 > spec.LastRow Or area.Column < spec.FirstColumn Or area.Column + area.Columns.Count - 1 > spec.LastColumn Then Err.Raise vbObjectError + 2810, , "EXCEL_MERGE_BOUNDS"
    anchorRef = AT_ExcelCellReference(sheetName, area.Row, area.Column)
    targetRow = area.Row: targetCol = area.Column
    If spec.RowOffset > 0 Then targetRow = targetRow + area.Rows.Count - 1
    If spec.ColumnOffset > 0 Then targetCol = targetCol + area.Columns.Count - 1
    targetRow = targetRow + spec.RowOffset: targetCol = targetCol + spec.ColumnOffset
    If targetRow < 1 Or targetRow > ws.Rows.Count Or targetCol < 1 Or targetCol > ws.Columns.Count Then Err.Raise vbObjectError + 2810, , "EXCEL_FIELD_TARGET_BOUNDS"
    Set target = ws.Cells(targetRow, targetCol).MergeArea.Cells(1, 1)
    cellRef = AT_ExcelCellReference(sheetName, target.Row, target.Column): found = True
    raw = target.Value2
    If IsError(raw) Then
        Set result = AT_RuntimeFailure(atEvalError, "EXCEL_CELL_ERROR", Nothing)
    ElseIf AT_ExcelBlank(raw) Then
        Set result = AT_RuntimeOK(AT_RuntimeState("TEXT", atValueEmpty))
    Else
        Set result = AT_RuntimeOK(AT_RuntimeScalar("TEXT", AT_ExcelText(raw)))
    End If
BuildValue:
    On Error GoTo 0
    value.Initialize id, found, sheetName, anchorRef, cellRef, result
    Set Read = value
    Exit Function
Failed:
    code = Err.Description
    If Left$(code, 6) <> "EXCEL_" Then code = "EXCEL_FIELD_READ_FAILED"
    Set result = AT_RuntimeFailure(atEvalError, code, Nothing)
    Resume BuildValue
End Function
