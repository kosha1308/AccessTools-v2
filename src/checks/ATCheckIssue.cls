VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATCheckIssue"
Option Explicit
Private mCheck As String, mCode As String, mSeverity As String, mField As String, mCell As String, mMessage As String, mFrozen As Boolean
Friend Sub Initialize(ByVal checkId As String, ByVal code As String, ByVal severity As String, ByVal fieldId As String, ByVal cellRef As String, ByVal message As String)
    If mFrozen Then Err.Raise vbObjectError + 2940, , "CHECK_VALUE_FROZEN"
    If Len(checkId) = 0 Or Len(checkId) > 128 Or Len(message) = 0 Or Len(message) > 32767 Then Err.Raise vbObjectError + 2940, , "CHECK_ISSUE_INVALID"
    If severity <> "ERROR" And severity <> "WARNING" And severity <> "INFO" Then Err.Raise vbObjectError + 2940, , "CHECK_ISSUE_INVALID"
    mCode = AT_CanonicalId(code): mCheck = checkId: mSeverity = severity
    mField = fieldId: mCell = cellRef: mMessage = message: mFrozen = True
End Sub
Public Property Get CheckId() As String
    CheckId = mCheck
End Property
Public Property Get Code() As String
    Code = mCode
End Property
Public Property Get Severity() As String
    Severity = mSeverity
End Property
Public Property Get FieldId() As String
    FieldId = mField
End Property
Public Property Get CellRef() As String
    CellRef = mCell
End Property
Public Property Get Message() As String
    Message = mMessage
End Property
