VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATInitiatorPolicyCompiler"
Option Explicit
Public Function Compile(ByVal initiators As ATEvalResult, ByVal ipRules As ATEvalResult, ByVal noPolicyAction As String) As ATInitiatorPolicyResult
    Dim policy As New ATInitiatorPolicy, result As New ATInitiatorPolicyResult, failure As ATEvalResult, code As String
    On Error GoTo Failed
    policy.Initialize initiators, ipRules, noPolicyAction
    result.Initialize atEvalOK, policy, Nothing: Set Compile = result
    Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "INITIATOR_INPUT_INVALID")
    If Left$(code, 10) <> "INITIATOR_" And Left$(code, 14) <> "RUNTIME_LIMIT_" Then code = "INITIATOR_INPUT_INVALID"
    Set failure = AT_RuntimeFailure(atEvalError, code, Nothing)
    result.Initialize atEvalError, Nothing, failure.Diagnostics: Set Compile = result
End Function
