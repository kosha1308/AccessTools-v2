VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkSupersetRules"
Option Explicit
Public Function Run(ByVal registry As ATNetworkRegistryResult, ByVal rules As ATNetworkRulebookResult, Optional ByVal externalBudget As ATEvalBudget = Nothing) As ATNetworkSupersetResult
    Dim result As New ATNetworkSupersetResult, annotations As New Collection, issues As New Collection, checks As New Collection
    Dim report As New ATCheckReport, outcome As New ATCheckOutcome, budget As New ATEvalBudget, limits As New ATEvalLimits
    Dim entry As ATNetworkRegistryEntry, row As ATRuntimeRow, manual As Boolean, status As String, code As String
    Dim issue As ATCheckIssue
    On Error GoTo Failed
    If registry Is Nothing Then Err.Raise vbObjectError + 3140, , "SUPERSET_RULES_INPUT_INVALID"
    If registry.State <> atEvalOK Then Err.Raise vbObjectError + 3140, , "SUPERSET_RULES_REGISTRY_BLOCKED"
    If Not rules Is Nothing Then
        If rules.State <> atEvalOK Then Err.Raise vbObjectError + 3140, , rules.Diagnostics(1).Code
    End If
    If externalBudget Is Nothing Then budget.Initialize limits Else Set budget = externalBudget
    If registry.Entries.Count = 0 Then manual = True
    For Each entry In registry.Entries
        budget.CheckList annotations.Count + 1: budget.Step
        Set row = Annotate(entry, rules, issues, manual, budget): annotations.Add row
    Next entry
    status = "PASS"
    If manual Then status = "SKIPPED"
    For Each issue In issues
        If issue.Severity = "ERROR" Then status = "FAIL"
    Next issue
    outcome.Initialize "NET.HOSTNAME_RULES", status, "Проверены только контексты Superset. Технические порты/протоколы оцениваются отдельно.", manual: checks.Add outcome
    report.Initialize atEvalOK, "SUPERSET_RULES", issues, checks
    result.Initialize atEvalOK, registry, annotations, report: Set Run = result: Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "SUPERSET_RULES_INPUT_INVALID")
    Set annotations = New Collection: Set issues = New Collection: Set checks = New Collection
    Set report = New ATCheckReport: Set outcome = New ATCheckOutcome: Set result = New ATNetworkSupersetResult
    AT_SupersetRuleIssue issues, code, "ERROR", "", "Правила Superset не проверены. Частичные аннотации не выданы; исходные результаты поиска сохранены."
    outcome.Initialize "NET.HOSTNAME_RULES", "BLOCKED", "Правила Superset не завершены.": checks.Add outcome
    report.Initialize atEvalError, "SUPERSET_RULES", issues, checks
    result.Initialize atEvalError, registry, annotations, report: Set Run = result
End Function
Private Function Annotate(ByVal entry As ATNetworkRegistryEntry, ByVal rules As ATNetworkRulebookResult, ByVal issues As Collection, ByRef incomplete As Boolean, ByVal budget As ATEvalBudget) As ATRuntimeRow
    Dim row As New ATRuntimeRow, matches As ATRuntimeValue, source As ATRuntimeRow, answer As ATEvalResult, hit As ATRuntimeRow
    Dim services As New ATSupersetTextSet, seen As Object, instructions As String, manual As Boolean, failed As Boolean
    Dim i As Long, j As Long, ip As String, host As String, service As String, valid As Boolean, block As String, key As String
    Dim ref As String, prefix As String, status As String, severity As String, conflict As Boolean
    Set seen = CreateObject("Scripting.Dictionary"): Set matches = entry.Matches
    ref = entry.References.Row(1).CellRef("CELL_REF")
    If rules Is Nothing Then
        manual = True: instructions = "Нет книги правил: сервисы и инструкции не вычислены."
    ElseIf matches.Count = 0 Then
        manual = True: instructions = "Нет строк Superset: сервисы и инструкции не вычислены."
    Else
        For i = 1 To matches.Count
            Set source = matches.Row(i): budget.Step
            ip = AT_SupersetIntersection(entry, source, budget, manual)
            host = AT_RegistryText(source.ResolveField("HOSTNAME", budget), valid, budget)
            If Not valid Then
                manual = True
            Else
                service = AT_RegistryText(source.ResolveField("SERVICE", budget), valid, budget)
                If Not valid Then
                    manual = True
                Else
                    Set answer = rules.Policy.Superset(ip, host, service, budget)
                    Select Case answer.State
                        Case atEvalOK
                            conflict = False
                            prefix = "№ " & entry.Id & ", IP/сеть " & entry.Address & ", строка Superset " & source.RowId & ": " & source.CellRef("HOSTNAME")
                            For j = 1 To answer.Value.Count
                                Set hit = answer.Value.Row(j): block = AT_SupersetRuleBlock(hit, budget)
                                key = hit.RowId
                                If Not seen.Exists(key) Then
                                    budget.CheckList seen.Count + 1: seen.Add key, True
                                    services.AddLines AT_PolicyCell(hit, "SERVICE_RESULT", budget)
                                    AT_SupersetAppend instructions, block
                                End If
                                severity = AT_PolicyCell(hit, "SEVERITY", budget)
                                If severity = "ERROR" Then failed = True
                                AT_SupersetRuleIssue issues, "RULEBOOK_MATCH", severity, ref, prefix & vbLf & block
                                If hit.ResolveField("CONFLICT", budget).Value.Scalar Then
                                    failed = True: conflict = True
                                End If
                            Next j
                            If conflict Then AT_SupersetRuleIssue issues, "RULEBOOK_CONFLICT", "ERROR", ref, prefix & vbLf & "Конфликт сервисов или групп Rulebook. Все результаты сохранены; требуется ручная проверка."
                        Case atEvalManualReview, atEvalSkipped: manual = True
                        Case Else: Err.Raise vbObjectError + 3140, , answer.Diagnostics(1).Code
                    End Select
                End If
            End If
        Next i
    End If
    If Not entry.SourceComplete Then manual = True
    If manual Then
        incomplete = True
        AT_SupersetAppend instructions, "Правила Superset проверены не полностью. Требуется источник или ручная проверка."
        AT_SupersetRuleIssue issues, "RULEBOOK_CONTEXT_UNCHECKED", "WARNING", ref, "№ " & entry.Id & ", " & entry.Address & ": правила Superset проверены не полностью; подтверждённые результаты других строк сохранены."
    End If
    status = "PASS": If manual Then status = "SKIPPED"
    If failed Then status = "FAIL"
    row.Initialize "SUPERSET_ANNOTATION", entry.Id
    row.AddField "SERVICE_RESULT", AT_RuntimeOK(AT_RuntimeScalar("TEXT", services.Text))
    row.AddField "INSTRUCTION", AT_RuntimeOK(AT_RuntimeScalar("TEXT", instructions))
    row.AddField "STATUS", AT_RuntimeOK(AT_RuntimeScalar("TEXT", status))
    Set Annotate = row.Snapshot
End Function
