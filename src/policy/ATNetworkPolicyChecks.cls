VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkPolicyChecks"
Option Explicit
Public Function Run(ByVal source As ATNetworkReadResult, ByVal loaded As ATNetworkPolicyResult) As ATNetworkPolicyReport
    Dim localRunner As New ATNetworkLocalChecks, policy As ATNetworkPolicy, localReport As ATCheckReport
    Dim result As New ATNetworkPolicyReport, report As New ATCheckReport
    Dim decisions As New Collection, issues As New Collection, checks As New Collection
    Dim issue As ATCheckIssue, outcome As ATCheckOutcome, updated As ATCheckOutcome, decision As ATNetworkPolicyDecision
    Dim status As String, code As String, errors As Object
    On Error GoTo Failed
    If loaded Is Nothing Then Err.Raise vbObjectError + 2960, , "POLICY_INPUT_INVALID"
    If loaded.State <> atEvalOK Then Err.Raise vbObjectError + 2960, , loaded.Diagnostics(1).Code
    Set policy = loaded.Policy
    Set localReport = localRunner.Run(source, policy.Protocols)
    If localReport.State <> atEvalOK Then
        result.Initialize localReport, decisions: Set Run = result: Exit Function
    End If
    Set errors = CreateObject("Scripting.Dictionary")
    For Each issue In localReport.Issues
        Set decision = policy.Decide(issue, AT_NetworkIssuePolicyKey(issue))
        decisions.Add decision: issues.Add decision.Issue
        If decision.Issue.Severity = "ERROR" Then errors(decision.Issue.CheckId) = True
    Next issue
    For Each outcome In localReport.Checks
        status = outcome.Status
        If status = "PASS" Or status = "FAIL" Then
            status = "PASS"
            If errors.Exists(outcome.CheckId) Then status = "FAIL"
        End If
        Set updated = New ATCheckOutcome
        updated.Initialize outcome.CheckId, status, outcome.Reason: checks.Add updated
    Next outcome
    report.Initialize atEvalOK, "NETWORK_LOCAL_POLICY", issues, checks
    result.Initialize report, decisions: Set Run = result
    Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "POLICY_INPUT_INVALID")
    If Left$(code, 7) <> "POLICY_" And Left$(code, 14) <> "RUNTIME_LIMIT_" And code <> "CHECK_REPORT_LIMIT" Then code = "POLICY_INPUT_INVALID"
    Set decisions = New Collection: Set issues = New Collection: Set checks = New Collection
    Set issue = New ATCheckIssue: Set outcome = New ATCheckOutcome: Set report = New ATCheckReport: Set result = New ATNetworkPolicyReport
    issue.Initialize "NET.POLICY", code, "ERROR", "", "", "Не удалось применить книгу правил. Исправьте настройки Rulebook и повторите проверку.": issues.Add issue
    outcome.Initialize "NET.POLICY", "BLOCKED", "Применение правил не завершено.": checks.Add outcome
    report.Initialize atEvalError, "NETWORK_LOCAL_POLICY", issues, checks
    result.Initialize report, decisions: Set Run = result
End Function
