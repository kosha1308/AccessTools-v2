VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkAccessChecks"
Option Explicit
Public Function RunWithRules(ByVal source As ATNetworkReadResult, ByVal catalog As ATNetworkPolicyResult, ByVal actor As ATInitiatorPolicyResult, ByVal rules As ATNetworkRulebookResult) As ATNetworkPolicyReport
    Dim localRunner As New ATNetworkLocalChecks, actorRunner As New ATNetworkInitiatorChecks, ruleRunner As New ATNetworkRuleChecks
    Dim presenter As New ATNetworkPolicyChecks, localReport As ATCheckReport, identity As ATCheckReport, ruleReport As ATCheckReport
    Dim combined As New ATCheckReport, issues As New Collection, checks As New Collection, issue As ATCheckIssue, outcome As ATCheckOutcome
    On Error GoTo FailedRules
    If catalog Is Nothing Then Set RunWithRules = presenter.Run(source, catalog): Exit Function
    If catalog.State <> atEvalOK Then Set RunWithRules = presenter.Run(source, catalog): Exit Function
    Set localReport = localRunner.Run(source, catalog.Policy.Protocols)
    If localReport.State <> atEvalOK Then Set RunWithRules = presenter.Apply(localReport, catalog): Exit Function
    Set identity = actorRunner.Run(source, actor)
    If identity.State <> atEvalOK Then Set RunWithRules = presenter.Apply(identity, catalog): Exit Function
    For Each issue In localReport.Issues: issues.Add issue: Next issue
    For Each issue In identity.Issues: issues.Add issue: Next issue
    For Each outcome In localReport.Checks: checks.Add outcome: Next outcome
    For Each outcome In identity.Checks: checks.Add outcome: Next outcome
    Set ruleReport = ruleRunner.Run(source, rules)
    If issues.Count + ruleReport.Issues.Count > 1000 Then GoTo RuleLimit
    For Each issue In ruleReport.Issues: issues.Add issue: Next issue
    For Each outcome In ruleReport.Checks: checks.Add outcome: Next outcome
    GoTo Present
RuleLimit:
    Set outcome = New ATCheckOutcome
    outcome.Initialize "NET.IP_EXCEPTIONS", "BLOCKED", "Превышен объём сообщений общих правил. Предыдущие проверки сохранены.": checks.Add outcome
    If issues.Count < 1000 Then
        Set issue = New ATCheckIssue
        issue.Initialize "NET.IP_EXCEPTIONS", "CHECK_REPORT_LIMIT", "ERROR", "", "", "Общие правила не включены: превышен объём отчёта.": issues.Add issue
    End If
Present:
    combined.Initialize atEvalOK, "NETWORK_ACCESS_PARTIAL", issues, checks
    Set RunWithRules = presenter.Apply(combined, catalog): Exit Function
FailedRules:
    Set RunWithRules = Run(source, catalog, actor)
End Function
Public Function Run(ByVal source As ATNetworkReadResult, ByVal catalog As ATNetworkPolicyResult, ByVal actor As ATInitiatorPolicyResult) As ATNetworkPolicyReport
    Dim localRunner As New ATNetworkLocalChecks, actorRunner As New ATNetworkInitiatorChecks, presenter As New ATNetworkPolicyChecks
    Dim localReport As ATCheckReport, identity As ATCheckReport, combined As New ATCheckReport
    Dim issues As New Collection, checks As New Collection, issue As ATCheckIssue, outcome As ATCheckOutcome, code As String
    On Error GoTo Failed
    If catalog Is Nothing Then Set Run = presenter.Run(source, catalog): Exit Function
    If catalog.State <> atEvalOK Then Set Run = presenter.Run(source, catalog): Exit Function
    Set localReport = localRunner.Run(source, catalog.Policy.Protocols)
    If localReport.State <> atEvalOK Then Set Run = presenter.Apply(localReport, catalog): Exit Function
    Set identity = actorRunner.Run(source, actor)
    If identity.State <> atEvalOK Then Set Run = presenter.Apply(identity, catalog): Exit Function
    For Each issue In localReport.Issues: issues.Add issue: Next issue
    For Each issue In identity.Issues: issues.Add issue: Next issue
    For Each outcome In localReport.Checks: checks.Add outcome: Next outcome
    For Each outcome In identity.Checks: checks.Add outcome: Next outcome
    combined.Initialize atEvalOK, "NETWORK_ACCESS_PARTIAL", issues, checks
    Set Run = presenter.Apply(combined, catalog)
    Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "INITIATOR_INPUT_INVALID")
    Set issues = New Collection: Set checks = New Collection: Set issue = New ATCheckIssue: Set outcome = New ATCheckOutcome: Set combined = New ATCheckReport
    issue.Initialize "NET.INITIATOR_POLICY", code, "ERROR", "", "", "Не удалось собрать полный результат проверок заявки.": issues.Add issue
    outcome.Initialize "NET.INITIATOR_POLICY", "BLOCKED", "Общий отчёт не завершён.": checks.Add outcome
    combined.Initialize atEvalError, "NETWORK_ACCESS_PARTIAL", issues, checks
    Set Run = presenter.Apply(combined, catalog)
End Function
