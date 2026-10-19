VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtTableSnapshot"
Option Explicit

Private mId As String, mStartRow As Long, mRows As Collection, mColumns As Collection
Private mStatus As CAtEvalResult

Friend Sub Initialize(ByVal id As String, ByVal startRow As Long, ByVal columns As Collection, ByVal status As CAtEvalResult)
    mId = id: mStartRow = startRow: Set mColumns = New Collection: Set mRows = New Collection
    Dim item As Variant
    For Each item In columns: mColumns.Add CStr(item): Next item
    Set mStatus = status.Snapshot()
End Sub

Friend Sub AppendRow(ByVal cells As Object)
    Dim row As Object, id As Variant, r As CAtEvalResult
    Set row = AtDictionary()
    For Each id In mColumns
        Set r = cells(id): row.Add CStr(id), r.Snapshot()
        If r.State <> "OK" And r.State <> "EMPTY" Then
            If mStatus.State = "OK" Then mStatus.SetData "Text", r.State, "", mId
            mStatus.MergeEffects r: mStatus.BoundDiagnostics 128, mId, ""
        End If
    Next id
    mRows.Add row
End Sub

Public Property Get TableId() As String
    TableId = mId
End Property
Public Property Get StartRow() As Long
    StartRow = mStartRow
End Property
Public Property Get RowCount() As Long
    RowCount = mRows.Count
End Property
Public Property Get ColumnIds() As Collection
    Dim copy As New Collection, item As Variant
    For Each item In mColumns: copy.Add CStr(item): Next item
    Set ColumnIds = copy
End Property
Public Property Get Status() As CAtEvalResult
    Set Status = mStatus.Snapshot()
End Property
Public Function ReadCell(ByVal rowIndex As Long, ByVal columnId As String) As CAtEvalResult
    If rowIndex < 1 Or rowIndex > mRows.Count Then Err.Raise 5, , "Row index outside snapshot"
    If Not mRows(rowIndex).Exists(columnId) Then Err.Raise 5, , "Column outside snapshot"
    Set ReadCell = mRows(rowIndex)(columnId).Snapshot()
End Function
Public Function IsEmptyRow(ByVal rowIndex As Long) As Boolean
    Dim id As Variant
    If rowIndex < 1 Or rowIndex > mRows.Count Then Err.Raise 5, , "Row index outside snapshot"
    IsEmptyRow = True
    For Each id In mColumns
        If ReadCell(rowIndex, CStr(id)).State <> "EMPTY" Then IsEmptyRow = False: Exit Function
    Next id
End Function
