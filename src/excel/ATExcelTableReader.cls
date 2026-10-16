VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATExcelTableReader"
Option Explicit

Public Function Read(ByVal ws As Worksheet, ByVal layout As ATExcelTableLayout, ByVal lastDataRow As Long) As ATEvalResult
    Dim firstRow As Long, firstCol As Long, lastCol As Long, block As Range
    Dim data As Variant, mergeState As Variant, hasMerges As Boolean
    Dim r As Long, i As Long, col As Long, area As Range, raw As Variant
    Dim ends() As Long, anchorsR() As Long, anchorsC() As Long
    Dim values() As Variant, refs() As String, used As Boolean
    Dim rows As New Collection, row As ATRuntimeRow, value As New ATRuntimeValue, field As ATEvalResult
    Dim code As String
    On Error GoTo Failed
    layout.ValidateSheet ws
    firstRow = layout.HeaderLastRow + 1: firstCol = layout.FirstColumn: lastCol = layout.LastColumn
    If lastDataRow < firstRow - 1 Or lastDataRow > ws.Rows.Count Then Err.Raise vbObjectError + 2800, , "EXCEL_DATA_BOUNDS"
    If lastDataRow - firstRow + 1 > 10000 Then Err.Raise vbObjectError + 2800, , "EXCEL_ROW_LIMIT"
    If CDbl(lastDataRow - firstRow + 1) * (lastCol - firstCol + 1) > 1000000# Then Err.Raise vbObjectError + 2800, , "EXCEL_CELL_LIMIT"
    If lastDataRow >= firstRow Then
        Set block = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastDataRow, lastCol))
        data = AT_ExcelBlock(block)
        mergeState = block.MergeCells
        If IsNull(mergeState) Then hasMerges = True Else hasMerges = CBool(mergeState)
        ReDim ends(1 To layout.Count): ReDim anchorsR(1 To layout.Count): ReDim anchorsC(1 To layout.Count)
        ReDim values(1 To layout.Count): ReDim refs(1 To layout.Count)
        For r = firstRow To lastDataRow
            used = False
            For i = 1 To layout.Count
                col = layout.ColumnNumber(i)
                If hasMerges Then
                    If r > ends(i) Then
                        Set area = ws.Cells(r, col).MergeArea
                        ValidateMerge area, layout, firstRow, lastDataRow
                        ends(i) = area.Row + area.Rows.Count - 1
                        anchorsR(i) = area.Row: anchorsC(i) = area.Column
                    End If
                Else
                    anchorsR(i) = r: anchorsC(i) = col
                End If
                raw = data(anchorsR(i) - firstRow + 1, anchorsC(i) - firstCol + 1)
                values(i) = raw
                refs(i) = CellReference(ws.Name, anchorsR(i), anchorsC(i))
                If Not AT_ExcelBlank(raw) Then
                    If layout.Usage(i) = "INHERITED" Or (layout.Usage(i) = "DIRECT" And anchorsR(i) = r) Then used = True
                End If
            Next i
            If used Then
                Set row = New ATRuntimeRow
                row.Initialize layout.TableId, CStr(r)
                For i = 1 To layout.Count
                    If IsError(values(i)) Then
                        Set field = AT_RuntimeFailure(atEvalError, "EXCEL_CELL_ERROR", Nothing)
                    ElseIf AT_ExcelBlank(values(i)) Then
                        Set field = AT_RuntimeOK(AT_RuntimeState("TEXT", atValueEmpty))
                    Else
                        Set field = AT_RuntimeOK(AT_RuntimeScalar("TEXT", AT_ExcelText(values(i))))
                    End If
                    row.AddField layout.FieldId(i), field, refs(i)
                Next i
                rows.Add row
            End If
        Next r
    End If
    value.InitializeRows layout.TableId, rows
    Set Read = AT_RuntimeOK(value)
    Exit Function
Failed:
    code = Err.Description
    If Left$(code, 6) <> "EXCEL_" Then code = "EXCEL_READ_FAILED"
    Set Read = AT_RuntimeFailure(atEvalError, code, Nothing)
End Function

Public Function LastValueRow(ByVal ws As Worksheet, ByVal layout As ATExcelTableLayout) As Long
    Dim i As Long, scan As Range, hit As Range, firstAddress As String, candidates As Long
    Dim area As Range, bottom As Long
    layout.ValidateSheet ws
    LastValueRow = layout.HeaderLastRow
    If layout.HeaderLastRow = ws.Rows.Count Then Exit Function
    For i = 1 To layout.Count
        Set scan = ws.Range(ws.Cells(layout.HeaderLastRow + 1, layout.ColumnNumber(i)), ws.Cells(ws.Rows.Count, layout.ColumnNumber(i)))
        Set hit = FindPreviousValue(scan, scan.Cells(1, 1))
        If Not hit Is Nothing Then
            firstAddress = hit.Address: candidates = 0
            Do
                candidates = candidates + 1
                If candidates > 10000 Then Err.Raise vbObjectError + 2800, , "EXCEL_FIND_LIMIT"
                If Not AT_ExcelBlank(hit.Value2) Then
                    Set area = hit.MergeArea
                    ValidateMerge area, layout, layout.HeaderLastRow + 1, ws.Rows.Count
                    bottom = area.Row + area.Rows.Count - 1
                    If bottom > LastValueRow Then LastValueRow = bottom
                    Exit Do
                End If
                Set hit = FindPreviousValue(scan, hit)
                If hit Is Nothing Then Exit Do
                If hit.Address = firstAddress Then Exit Do
            Loop
        End If
    Next i
End Function

Private Sub ValidateMerge(ByVal area As Range, ByVal layout As ATExcelTableLayout, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim i As Long, mapped As Long, lastCol As Long, col As Long
    lastCol = area.Column + area.Columns.Count - 1
    If area.Row < firstRow Or area.Row + area.Rows.Count - 1 > lastRow Or area.Column < layout.FirstColumn Or lastCol > layout.LastColumn Then Err.Raise vbObjectError + 2800, , "EXCEL_MERGE_BOUNDS"
    For i = 1 To layout.Count
        col = layout.ColumnNumber(i)
        If col >= area.Column And col <= lastCol Then mapped = mapped + 1
    Next i
    If mapped > 1 Then Err.Raise vbObjectError + 2800, , "EXCEL_MERGE_COLUMNS"
End Sub

Private Function FindPreviousValue(ByVal scan As Range, ByVal after As Range) As Range
    Set FindPreviousValue = scan.Find(What:="*", After:=after, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
End Function

Private Function CellReference(ByVal sheetName As String, ByVal rowNumber As Long, ByVal columnNumber As Long) As String
    Dim letters As String, remainder As Long
    Do While columnNumber > 0
        remainder = (columnNumber - 1) Mod 26
        letters = Chr$(65 + remainder) & letters
        columnNumber = (columnNumber - 1) \ 26
    Loop
    CellReference = "'" & Replace(sheetName, "'", "''") & "'!" & letters & CStr(rowNumber)
End Function
