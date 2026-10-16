VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATExcelFieldValue"
Option Explicit
Private mId As String, mFound As Boolean, mSheet As String, mAnchor As String, mCell As String
Private mResult As ATEvalResult, mFrozen As Boolean

Friend Sub Initialize(ByVal fieldId As String, ByVal found As Boolean, ByVal sheetName As String, ByVal anchorRef As String, ByVal cellRef As String, ByVal result As ATEvalResult)
    If mFrozen Then Err.Raise vbObjectError + 2810, , "EXCEL_FIELD_VALUE_FROZEN"
    mId = fieldId: mFound = found: mSheet = sheetName: mAnchor = anchorRef: mCell = cellRef
    Set mResult = result.Clone: mFrozen = True
End Sub
Public Property Get FieldId() As String
    FieldId = mId
End Property
Public Property Get Found() As Boolean
    Found = mFound
End Property
Public Property Get SheetName() As String
    SheetName = mSheet
End Property
Public Property Get AnchorRef() As String
    AnchorRef = mAnchor
End Property
Public Property Get CellRef() As String
    CellRef = mCell
End Property
Public Property Get Result() As ATEvalResult
    If Not mFrozen Then Err.Raise vbObjectError + 2810, , "EXCEL_FIELD_VALUE_REQUIRED"
    Set Result = mResult.Clone
End Property
