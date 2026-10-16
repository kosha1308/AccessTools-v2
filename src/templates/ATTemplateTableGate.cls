VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATTemplateTableGate"
Option Explicit
Public Function Read(ByVal profile As ATTemplateProfile, ByVal ws As Worksheet, ByVal layout As ATExcelTableLayout, ByVal lastDataRow As Long) As ATEvalResult
    Dim decision As ATTemplateGateResult, reader As New ATExcelTableReader
    On Error GoTo Failed
    If profile Is Nothing Then
        Set decision = New ATTemplateGateResult
        decision.Initialize False, "TEMPLATE_PROFILE_INVALID", "PROFILE_REQUIRED"
    Else
        Set decision = profile.Verify(ws)
    End If
    If decision.Allowed Then
        Set Read = reader.Read(ws, layout, lastDataRow)
    Else
        Set Read = Rejected(decision)
    End If
    Exit Function
Failed:
    Set decision = New ATTemplateGateResult
    decision.Initialize False, "TEMPLATE_CHECK_ERROR", "READ_FAILED"
    Set Read = Rejected(decision)
End Function
Private Function Rejected(ByVal decision As ATTemplateGateResult) As ATEvalResult
    Dim diagnostics As New Collection, diagnostic As New ATDiagnostic, result As New ATEvalResult
    diagnostic.Initialize decision.Code, "TEMPLATE", "ERROR", 1, 0, decision.Message
    diagnostics.Add diagnostic
    result.Initialize atEvalError, Nothing, diagnostics, Nothing
    Set Rejected = result
End Function
