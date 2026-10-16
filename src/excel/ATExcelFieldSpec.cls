VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATExcelFieldSpec"
Option Explicit
Private mId As String, mAnchor As String, mTop As Long, mBottom As Long
Private mLeft As Long, mRight As Long, mRowOffset As Long, mColOffset As Long
Private mInitialized As Boolean

Public Sub Initialize(ByVal fieldId As String, ByVal anchorText As String, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstColumn As Long, ByVal lastColumn As Long, ByVal rowOffset As Long, ByVal columnOffset As Long)
    If mInitialized Then Err.Raise vbObjectError + 2810, , "EXCEL_FIELD_SPEC_FROZEN"
    If firstRow < 1 Or lastRow > 1048576 Or lastRow < firstRow Then GoTo Invalid
    If firstColumn < 1 Or lastColumn > 16384 Or lastColumn < firstColumn Then GoTo Invalid
    If CDbl(lastRow - firstRow + 1) * (lastColumn - firstColumn + 1) > 1000000# Then GoTo Invalid
    If rowOffset < -1048575 Or rowOffset > 1048575 Or columnOffset < -16383 Or columnOffset > 16383 Then GoTo Invalid
    If rowOffset = 0 And columnOffset = 0 Then GoTo Invalid
    If Len(AT_ExcelHeader(anchorText)) = 0 Then Err.Raise vbObjectError + 2810, , "EXCEL_ANCHOR_REQUIRED"
    mId = AT_CanonicalId(fieldId): mAnchor = anchorText
    mTop = firstRow: mBottom = lastRow: mLeft = firstColumn: mRight = lastColumn
    mRowOffset = rowOffset: mColOffset = columnOffset: mInitialized = True
    Exit Sub
Invalid:
    Err.Raise vbObjectError + 2810, , "EXCEL_FIELD_SPEC_BOUNDS"
End Sub
Public Sub Validate()
    If Not mInitialized Then Err.Raise vbObjectError + 2810, , "EXCEL_FIELD_SPEC_REQUIRED"
End Sub
Public Property Get FieldId() As String
    FieldId = mId
End Property
Public Property Get AnchorText() As String
    AnchorText = mAnchor
End Property
Public Property Get FirstRow() As Long
    FirstRow = mTop
End Property
Public Property Get LastRow() As Long
    LastRow = mBottom
End Property
Public Property Get FirstColumn() As Long
    FirstColumn = mLeft
End Property
Public Property Get LastColumn() As Long
    LastColumn = mRight
End Property
Public Property Get RowOffset() As Long
    RowOffset = mRowOffset
End Property
Public Property Get ColumnOffset() As Long
    ColumnOffset = mColOffset
End Property
