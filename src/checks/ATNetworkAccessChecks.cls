VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkAccessChecks"
Option Explicit
Public Function ApplyRegistry(ByVal previous As ATNetworkPolicyReport, ByVal registry As ATNetworkRegistryResult, ByVal catalog As ATNetworkPolicyResult) As ATNetworkPolicyReport
    Dim presenter As New ATNetworkPolicyChecks, added As ATNetworkPolicyReport, ready As New ATCheckReport
    Dim raw As ATCheckReport, capacity As Long, code As String
    On Error GoTo FailedRegistry
    If previous Is Nothing Then Set ApplyRegistry = presenter.Apply(Nothing, catalog): Exit Function
    If previous.AllReport.State <> atEvalOK Then Set ApplyRegistry = previous: Exit Function
    capacity = 1000 - previous.AllReport.Issues.Count
    If registry Is Nothing Then
        Set added = BlockedRegistry(catalog, "REGISTRY_INPUT_INVALID", capacity)
    Else
        Set raw = registry.Report
        If raw.Issues.Count > capacity Then
            Set added = BlockedRegistry(catalog, "CHECK_REPORT_LIMIT", capacity)
        Else
            ready.Initialize atEvalOK, "NETWORK_REGISTRY", raw.Issues, raw.Checks
            Set added = presenter.Apply(ready, catalog)
            If added.AllReport.State <> atEvalOK Then Set added = BlockedRegistry(catalog, added.AllReport.Issues(1).Code, capacity)
        End If
    End If
    Set ApplyRegistry = Combine(previous, added, "NETWORK_ACCESS_REGISTRY_POLICY"): Exit Function
FailedRegistry:
    code = AT_StableErrorCode(Err.Description, "REGISTRY_INPUT_INVALID")
    If previous Is Nothing Then Set ApplyRegistry = presenter.Apply(Nothing, catalog): Exit Function
    If previous.AllReport.State <> atEvalOK Then Set ApplyRegistry = previous: Exit Function
    Set ApplyRegistry = Combine(previous, BlockedRegistry(catalog, code, 1000 - previous.AllReport.Issues.Count), "NETWORK_ACCESS_REGISTRY_POLICY")
End Function
Private Function BlockedRegistry(ByVal catalog As ATNetworkPolicyResult, ByVal code As String, ByVal capacity As Long) As ATNetworkPolicyReport
    Dim raw As New ATCheckReport, issues As New Collection, checks As New Collection, outcome As New ATCheckOutcome
    Dim presenter As New ATNetworkPolicyChecks
    If capacity > 0 Then AT_RegistryIssue issues, code, "ERROR", "REGISTRY", "", "Не удалось подготовить результат проверки Superset. Предыдущие проверки сохранены."
    outcome.Initialize "NET.IP_REGISTRY", "BLOCKED", "Проверка Superset не завершена.": checks.Add outcome
    raw.Initialize atEvalOK, "NETWORK_REGISTRY", issues, checks: Set BlockedRegistry = presenter.Apply(raw, catalog)
End Function
Public Function RunWithRules(ByVal source As ATNetworkReadResult, ByVal catalog As ATNetworkPolicyResult, ByVal actor As ATInitiatorPolicyResult, ByVal rules As ATNetworkRulebookResult) As ATNetworkPolicyReport
    Dim ruleRunner As New ATNetworkRuleChecks, presenter As New ATNetworkPolicyChecks
    Dim previous As ATNetworkPolicyReport, added As ATNetworkPolicyReport, raw As ATCheckReport, ready As New ATCheckReport
    Dim capacity As Long, code As String
    On Error GoTo FailedRules
    Set previous = Run(source, catalog, actor)
    If previous.AllReport.State <> atEvalOK Then Set RunWithRules = previous: Exit Function
    capacity = 1000 - previous.AllReport.Issues.Count
    Set raw = ruleRunner.Run(source, rules)
    If raw.Issues.Count > capacity Then
        Set added = BlockedRules(catalog, "CHECK_REPORT_LIMIT", capacity)
    Else
        ' The new component can be BLOCKED while the completed report is assembled.
        ready.Initialize atEvalOK, "NETWORK_RULES", raw.Issues, raw.Checks
        Set added = presenter.Apply(ready, catalog)
        If added.AllReport.State <> atEvalOK Then Set added = BlockedRules(catalog, added.AllReport.Issues(1).Code, capacity)
    End If
    Set RunWithRules = Combine(previous, added): Exit Function
FailedRules:
    code = AT_StableErrorCode(Err.Description, "RULEBOOK_INPUT_INVALID")
    If previous Is Nothing Then Set RunWithRules = Run(source, catalog, actor): Exit Function
    If previous.AllReport.State <> atEvalOK Then Set RunWithRules = previous: Exit Function
    Set RunWithRules = Combine(previous, BlockedRules(catalog, code, 1000 - previous.AllReport.Issues.Count))
End Function
Private Function BlockedRules(ByVal catalog As ATNetworkPolicyResult, ByVal code As String, ByVal capacity As Long) As ATNetworkPolicyReport
    Dim raw As New ATCheckReport, issues As New Collection, checks As New Collection, issue As New ATCheckIssue, outcome As New ATCheckOutcome
    Dim presenter As New ATNetworkPolicyChecks
    If capacity > 0 Then
        issue.Initialize "NET.IP_EXCEPTIONS", code, "ERROR", "TECHNICAL_ROWS", "", "Общие правила не включены: не удалось подготовить итоговые сообщения. Проверьте настройки текста и объём отчёта.": issues.Add issue
    End If
    outcome.Initialize "NET.IP_EXCEPTIONS", "BLOCKED", "Общие правила не завершены. Предыдущие проверки сохранены.": checks.Add outcome
    raw.Initialize atEvalOK, "NETWORK_RULES", issues, checks: Set BlockedRules = presenter.Apply(raw, catalog)
End Function
Private Function Combine(ByVal previous As ATNetworkPolicyReport, ByVal added As ATNetworkPolicyReport, Optional ByVal scope As String = "NETWORK_ACCESS_PARTIAL_POLICY") As ATNetworkPolicyReport
    Dim report As New ATCheckReport, result As New ATNetworkPolicyReport, issues As New Collection, checks As New Collection
    Dim decisions As New Collection, issue As ATCheckIssue, outcome As ATCheckOutcome, decision As ATNetworkPolicyDecision
    For Each issue In previous.AllReport.Issues: issues.Add issue: Next issue
    For Each issue In added.AllReport.Issues: issues.Add issue: Next issue
    For Each outcome In previous.AllReport.Checks: checks.Add outcome: Next outcome
    For Each outcome In added.AllReport.Checks: checks.Add outcome: Next outcome
    For Each decision In previous.Decisions: decisions.Add decision: Next decision
    For Each decision In added.Decisions: decisions.Add decision: Next decision
    report.Initialize atEvalOK, scope, issues, checks
    result.Initialize report, decisions: Set Combine = result
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
