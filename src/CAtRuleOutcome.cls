VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtRuleOutcome"
Option Explicit

Private mRule As String, mStatus As String, mResult As CAtEvalResult
Public Property Get RuleId() As String
    RuleId = mRule
End Property
Public Property Get Status() As String
    Status = mStatus
End Property
Public Property Get Result() As CAtEvalResult
    Set Result = mResult
End Property
Friend Sub Initialize(ByVal ruleId As String, ByVal result As CAtEvalResult)
    mRule = ruleId: Set mResult = result: mStatus = "CANNOT_EVALUATE"
    If result Is Nothing Then Exit Sub
    If Not result.IsComplete Or result.ResultType <> "Bool" Then Exit Sub
    If result.Value = "true" Then mStatus = "PASS"
    If result.Value = "false" Then mStatus = "FAIL"
End Sub
