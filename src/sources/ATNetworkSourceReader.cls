VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkSourceReader"
Option Explicit

' Raw source snapshots include hidden rows. Filtering belongs to the result view.
Public Function Read(ByVal ws As Worksheet, ByVal sourceId As String) As ATEvalResult
    Dim spec As ATExcelTableSpec, layout As ATExcelTableLayout
    Dim lastRow As Long, firstRow As Long, data As Variant, block As Range, mergeState As Variant
    Dim rows As New Collection, row As ATRuntimeRow, result As New ATRuntimeValue
    Dim r As Long, i As Long, col As Long, used As Boolean, raw As Variant, field As ATEvalResult
    Dim code As String
    On Error GoTo Failed
    Set spec = AT_NetworkSourceSpec(sourceId)
    If ws Is Nothing Then Err.Raise vbObjectError + 2860, , "SOURCE_READ_FAILED"
    Set layout = spec.Locate(ws)
    If layout.HeaderFirstRow <> layout.HeaderLastRow Then Err.Raise vbObjectError + 2860, , "SOURCE_MERGED_TABLE"
    Set block = ws.Range(ws.Cells(layout.HeaderFirstRow, layout.FirstColumn), ws.Cells(layout.HeaderLastRow, layout.LastColumn))
    RejectMerges block
    lastRow = LastSourceRow(ws, layout): firstRow = layout.HeaderLastRow + 1
    If lastRow - firstRow + 1 > 100000 Then Err.Raise vbObjectError + 2860, , "SOURCE_ROW_LIMIT"
    If CDbl(lastRow - firstRow + 1) * (layout.LastColumn - layout.FirstColumn + 1) > 1000000# Then Err.Raise vbObjectError + 2860, , "SOURCE_CELL_LIMIT"
    If lastRow >= firstRow Then
        Set block = ws.Range(ws.Cells(firstRow, layout.FirstColumn), ws.Cells(lastRow, layout.LastColumn))
        RejectMerges block
        data = AT_ExcelBlock(block)
        For r = 1 To UBound(data, 1)
            used = False
            For i = 1 To layout.Count
                col = layout.ColumnNumber(i) - layout.FirstColumn + 1
                If Not AT_ExcelBlank(data(r, col)) Then used = True: Exit For
            Next i
            If used Then
                Set row = New ATRuntimeRow
                row.Initialize layout.TableId, CStr(firstRow + r - 1)
                For i = 1 To layout.Count
                    col = layout.ColumnNumber(i): raw = data(r, col - layout.FirstColumn + 1)
                    If IsError(raw) Then
                        Set field = AT_RuntimeFailure(atEvalError, "EXCEL_CELL_ERROR", Nothing)
                    ElseIf AT_ExcelBlank(raw) Then
                        Set field = AT_RuntimeOK(AT_RuntimeState("TEXT", atValueEmpty))
                    Else
                        Set field = AT_RuntimeOK(AT_RuntimeScalar("TEXT", AT_ExcelText(raw)))
                    End If
                    row.AddField layout.FieldId(i), field, AT_ExcelCellReference(ws.Name, firstRow + r - 1, col)
                Next i
                rows.Add row
            End If
        Next r
    End If
    result.InitializeRows layout.TableId, rows
    Set Read = AT_RuntimeOK(result)
    Exit Function
Failed:
    code = Err.Description
    If Left$(code, 7) <> "SOURCE_" And Left$(code, 6) <> "EXCEL_" Then code = "SOURCE_READ_FAILED"
    Set Read = AT_RuntimeFailure(atEvalError, code, Nothing)
End Function

Private Sub RejectMerges(ByVal block As Range)
    Dim merged As Variant
    merged = block.MergeCells
    If IsNull(merged) Then Err.Raise vbObjectError + 2860, , "SOURCE_MERGED_TABLE"
    If CBool(merged) Then Err.Raise vbObjectError + 2860, , "SOURCE_MERGED_TABLE"
End Sub

Private Function LastSourceRow(ByVal ws As Worksheet, ByVal layout As ATExcelTableLayout) As Long
    Dim scan As Range, populated As Range, area As Range, kind As Variant
    Dim i As Long, r As Long, work As Long, data As Variant, number As Long
    LastSourceRow = layout.HeaderLastRow
    For i = 1 To layout.Count
        Set scan = ws.Range(ws.Cells(layout.HeaderLastRow + 1, layout.ColumnNumber(i)), ws.Cells(ws.Rows.Count, layout.ColumnNumber(i)))
        ' Find/End skip filtered rows even with xlFormulas. Constants/formulas do not.
        For Each kind In Array(xlCellTypeConstants, xlCellTypeFormulas)
            Set populated = Nothing
            On Error Resume Next
            Set populated = scan.SpecialCells(CLng(kind)): number = Err.Number
            Err.Clear
            On Error GoTo 0
            If number <> 0 And number <> 1004 Then Err.Raise vbObjectError + 2860, , "SOURCE_READ_FAILED"
            If Not populated Is Nothing Then
                For Each area In populated.Areas
                    work = work + 1
                    If work > 100000 Then Err.Raise vbObjectError + 2860, , "SOURCE_WORK_LIMIT"
                    If area.Row + area.Rows.Count - 1 > LastSourceRow Then
                        If area.Rows.Count > 100000 Then Err.Raise vbObjectError + 2860, , "SOURCE_ROW_LIMIT"
                        data = AT_ExcelBlock(area)
                        For r = UBound(data, 1) To 1 Step -1
                            work = work + 1
                            If work > 100000 Then Err.Raise vbObjectError + 2860, , "SOURCE_WORK_LIMIT"
                            If Not AT_ExcelBlank(data(r, 1)) Then
                                If area.Row + r - 1 > LastSourceRow Then LastSourceRow = area.Row + r - 1
                                Exit For
                            End If
                        Next r
                    End If
                Next area
            End If
        Next kind
    Next i
End Function
