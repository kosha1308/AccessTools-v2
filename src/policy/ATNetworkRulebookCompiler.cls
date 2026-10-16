VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkRulebookCompiler"
Option Explicit
Public Function Compile(ByVal raw As ATEvalResult, ByVal protocols As Collection) As ATNetworkRulebookResult
    Dim r As New ATNetworkRulebookResult, f As ATEvalResult, policy As New ATNetworkRulebook
    On Error GoTo Failed
    policy.Initialize raw, protocols
    r.Initialize atEvalOK, policy, Nothing: Set Compile = r
    Exit Function
Failed:
    Set f = AT_RuleFailure(Err.Description)
    r.Initialize atEvalError, Nothing, f.Diagnostics: Set Compile = r
End Function
