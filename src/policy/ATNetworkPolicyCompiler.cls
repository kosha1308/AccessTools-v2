VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkPolicyCompiler"
Option Explicit
Public Function Compile(ByVal protocols As ATEvalResult, ByVal problems As ATEvalResult, ByVal routing As ATEvalResult, ByVal commonText As ATEvalResult) As ATNetworkPolicyResult
    Dim policy As New ATNetworkPolicy, result As New ATNetworkPolicyResult, failure As ATEvalResult, code As String
    On Error GoTo Failed
    policy.Initialize protocols, problems, routing, commonText
    result.Initialize atEvalOK, policy, Nothing: Set Compile = result
    Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "POLICY_INPUT_INVALID")
    If Left$(code, 7) <> "POLICY_" And Left$(code, 14) <> "RUNTIME_LIMIT_" Then code = "POLICY_INPUT_INVALID"
    Set failure = AT_RuntimeFailure(atEvalError, code, Nothing)
    result.Initialize atEvalError, Nothing, failure.Diagnostics: Set Compile = result
End Function
