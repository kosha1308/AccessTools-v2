VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATTemplateTableGate"
Option Explicit
Public Function Read(ByVal profile As ATTemplateProfile, ByVal ws As Worksheet, ByVal spec As ATExcelTableSpec, ByVal lastDataRow As Long) As ATEvalResult
    Dim decision As ATTemplateGateResult, reader As New ATExcelTableReader
    Dim layout As ATExcelTableLayout
    On Error GoTo Failed
    If profile Is Nothing Then
        Set decision = New ATTemplateGateResult
        decision.Initialize False, "TEMPLATE_PROFILE_INVALID", "PROFILE_REQUIRED"
    Else
        Set decision = profile.Verify(ws)
    End If
    If decision.Allowed Then
        If spec Is Nothing Then
            Set Read = AT_RuntimeFailure(atEvalError, "EXCEL_SPEC_REQUIRED", Nothing)
            Exit Function
        End If
        Set layout = spec.Locate(ws)
        Set Read = reader.Read(ws, layout, lastDataRow)
    Else
        Set Read = Rejected(decision)
    End If
    Exit Function
Failed:
    If Left$(Err.Description, 6) = "EXCEL_" Then
        Set Read = AT_RuntimeFailure(atEvalError, Err.Description, Nothing)
        Exit Function
    End If
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
