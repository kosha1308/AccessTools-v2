VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATExcelTableLayout"
Option Explicit
Private mSheet As Worksheet, mTable As String, mTop As Long, mBottom As Long
Private mFirst As Long, mLast As Long, mColumns As Collection

Friend Sub Initialize(ByVal ws As Worksheet, ByVal tableId As String, ByVal topRow As Long, ByVal bottomRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal columns As Collection)
    Dim item As Variant
    If Not mSheet Is Nothing Then Err.Raise vbObjectError + 2800, , "EXCEL_LAYOUT_FROZEN"
    Set mSheet = ws: mTable = tableId: mTop = topRow: mBottom = bottomRow
    mFirst = firstCol: mLast = lastCol: Set mColumns = New Collection
    For Each item In columns
        mColumns.Add Array(CStr(item(0)), CLng(item(1)), CStr(item(2)))
    Next item
End Sub
Public Sub ValidateSheet(ByVal ws As Worksheet)
    If Not mSheet Is Nothing Then
        If ws Is mSheet Then Exit Sub
    End If
    Err.Raise vbObjectError + 2800, , "EXCEL_LAYOUT_SHEET_MISMATCH"
End Sub
Public Property Get TableId() As String
    TableId = mTable
End Property
Public Property Get HeaderFirstRow() As Long
    HeaderFirstRow = mTop
End Property
Public Property Get HeaderLastRow() As Long
    HeaderLastRow = mBottom
End Property
Public Property Get FirstColumn() As Long
    FirstColumn = mFirst
End Property
Public Property Get LastColumn() As Long
    LastColumn = mLast
End Property
Public Property Get Count() As Long
    Count = mColumns.Count
End Property
Public Function FieldId(ByVal index As Long) As String
    FieldId = mColumns(index)(0)
End Function
Public Function ColumnNumber(ByVal index As Long) As Long
    ColumnNumber = mColumns(index)(1)
End Function
Public Function Usage(ByVal index As Long) As String
    Usage = mColumns(index)(2)
End Function
