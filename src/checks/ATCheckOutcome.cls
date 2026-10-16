VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATCheckOutcome"
Option Explicit
Private mCheck As String, mStatus As String, mReason As String, mFrozen As Boolean, mIncomplete As Boolean
Friend Sub Initialize(ByVal checkId As String, ByVal status As String, ByVal reason As String, Optional ByVal incomplete As Boolean = False)
    If mFrozen Then Err.Raise vbObjectError + 2940, , "CHECK_VALUE_FROZEN"
    If Len(checkId) = 0 Or Len(checkId) > 128 Then Err.Raise vbObjectError + 2940, , "CHECK_OUTCOME_INVALID"
    If status <> "PASS" And status <> "FAIL" And status <> "SKIPPED" And status <> "BLOCKED" Then Err.Raise vbObjectError + 2940, , "CHECK_OUTCOME_INVALID"
    mCheck = checkId: mStatus = status: mReason = reason: mIncomplete = incomplete: mFrozen = True
End Sub
Public Property Get CheckId() As String
    CheckId = mCheck
End Property
Public Property Get Status() As String
    Status = mStatus
End Property
Public Property Get Reason() As String
    Reason = mReason
End Property
Public Property Get Incomplete() As Boolean
    Incomplete = mIncomplete
End Property
