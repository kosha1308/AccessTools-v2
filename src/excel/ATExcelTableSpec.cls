VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATExcelTableSpec"
Option Explicit
Private mTable As String, mTop As Long, mBottom As Long, mLeft As Long, mRight As Long
Private mColumns As Collection, mInitialized As Boolean

Public Sub Initialize(ByVal tableId As String, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstColumn As Long, ByVal lastColumn As Long)
    If mInitialized Then Err.Raise vbObjectError + 2800, , "EXCEL_SPEC_INITIALIZED"
    If firstRow < 1 Or lastRow > 1048576 Or lastRow < firstRow Then GoTo Invalid
    If firstColumn < 1 Or lastColumn > 16384 Or lastColumn < firstColumn Then GoTo Invalid
    If CDbl(lastRow - firstRow + 1) * (lastColumn - firstColumn + 1) > 1000000# Then GoTo Invalid
    mTable = AT_CanonicalId(tableId): mTop = firstRow: mBottom = lastRow
    mLeft = firstColumn: mRight = lastColumn: Set mColumns = New Collection
    mInitialized = True
    Exit Sub
Invalid:
    Err.Raise vbObjectError + 2800, , "EXCEL_SPEC_BOUNDS"
End Sub

Public Sub AddColumn(ByVal fieldId As String, ByVal headerText As String, ByVal usage As String)
    Dim item As Variant, key As String
    If Not mInitialized Then Err.Raise vbObjectError + 2800, , "EXCEL_SPEC_REQUIRED"
    fieldId = AT_CanonicalId(fieldId): key = AT_ExcelHeader(headerText)
    If Len(key) = 0 Then Err.Raise vbObjectError + 2800, , "EXCEL_HEADER_REQUIRED"
    If usage <> "INHERITED" And usage <> "DIRECT" And usage <> "NONE" Then Err.Raise vbObjectError + 2800, , "EXCEL_USAGE_INVALID"
    If mColumns.Count >= 256 Then Err.Raise vbObjectError + 2800, , "EXCEL_COLUMN_LIMIT"
    For Each item In mColumns
        If item(0) = fieldId Or item(1) = key Then Err.Raise vbObjectError + 2800, , "EXCEL_COLUMN_DUPLICATE"
    Next item
    mColumns.Add Array(fieldId, key, usage)
End Sub

Public Function Locate(ByVal ws As Worksheet) As ATExcelTableLayout
    Dim data As Variant, r As Long, c As Long, i As Long, key As String
    Dim hits() As Long, positions() As Long, item As Variant, complete As Boolean
    Dim chosen As Long, found As Collection, area As Range
    Dim bottom As Long, leftCol As Long, rightCol As Long, result As New ATExcelTableLayout
    If Not mInitialized Then Err.Raise vbObjectError + 2800, , "EXCEL_SPEC_REQUIRED"
    If mColumns.Count = 0 Then Err.Raise vbObjectError + 2800, , "EXCEL_COLUMN_REQUIRED"
    data = AT_ExcelBlock(ws.Range(ws.Cells(mTop, mLeft), ws.Cells(mBottom, mRight)))
    For r = 1 To UBound(data, 1)
        ReDim hits(1 To mColumns.Count): ReDim positions(1 To mColumns.Count)
        For c = 1 To UBound(data, 2)
            If Not IsError(data(r, c)) Then
                If Not AT_ExcelBlank(data(r, c)) Then
                    key = AT_ExcelHeader(AT_ExcelText(data(r, c)))
                    For i = 1 To mColumns.Count
                        item = mColumns(i)
                        If key = item(1) Then hits(i) = hits(i) + 1: positions(i) = c + mLeft - 1
                    Next i
                End If
            End If
        Next c
        complete = True
        For i = 1 To mColumns.Count
            If hits(i) = 0 Then complete = False
        Next i
        If complete Then
            For i = 1 To mColumns.Count
                If hits(i) > 1 Then Err.Raise vbObjectError + 2800, , "EXCEL_HEADER_DUPLICATE"
            Next i
            If chosen > 0 Then Err.Raise vbObjectError + 2800, , "EXCEL_HEADER_AMBIGUOUS"
            chosen = r + mTop - 1: bottom = chosen: leftCol = mRight: rightCol = mLeft
            Set found = New Collection
            For i = 1 To mColumns.Count
                item = mColumns(i)
                Set area = ws.Cells(chosen, positions(i)).MergeArea
                If area.Row < mTop Or area.Row + area.Rows.Count - 1 > mBottom Or area.Column < mLeft Or area.Column + area.Columns.Count - 1 > mRight Then Err.Raise vbObjectError + 2800, , "EXCEL_MERGE_BOUNDS"
                If area.Row + area.Rows.Count - 1 > bottom Then bottom = area.Row + area.Rows.Count - 1
                If area.Column < leftCol Then leftCol = area.Column
                If area.Column + area.Columns.Count - 1 > rightCol Then rightCol = area.Column + area.Columns.Count - 1
                found.Add Array(item(0), area.Column, item(2))
            Next i
        End If
    Next r
    If chosen = 0 Then Err.Raise vbObjectError + 2800, , "EXCEL_HEADER_MISSING"
    result.Initialize ws, mTable, chosen, bottom, leftCol, rightCol, found
    Set Locate = result
End Function
