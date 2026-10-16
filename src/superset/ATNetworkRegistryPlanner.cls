VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkRegistryPlanner"
Option Explicit
Public Function Run(ByVal source As ATNetworkReadResult, ByVal loaded As ATIPSourceResult, Optional ByVal externalBudget As ATEvalBudget = Nothing) As ATNetworkRegistryResult
    Dim result As New ATNetworkRegistryResult, entries As New Collection, issues As New Collection, checks As New Collection
    Dim report As New ATCheckReport, outcome As New ATCheckOutcome, builders As Object, ordered As New Collection
    Dim budget As New ATEvalBudget, limits As New ATEvalLimits, context As ATEvaluationContext
    Dim item As ATNetworkRegistryEntry, query As ATIPMatchResult, raw As ATRuntimeValue, emptyRows As New Collection
    Dim sourceIssues As ATEvalResult, available As Boolean, complete As Boolean, manual As Boolean, hostIncomplete As Boolean
    Dim issue As ATCheckIssue, i As Long, links As Long, status As String, reason As String, code As String
    On Error GoTo Failed
    If source Is Nothing Then Err.Raise vbObjectError + 3010, , "REGISTRY_INPUT_INVALID"
    If source.State <> atEvalOK Then Err.Raise vbObjectError + 3010, , source.Diagnostics(1).Code
    If externalBudget Is Nothing Then
        budget.Initialize limits
    Else
        Set budget = externalBudget
    End If
    Set context = source.Context: context.Validate: Set builders = CreateObject("Scripting.Dictionary")
    Collect context, False, builders, ordered, issues, manual, budget
    Collect context, True, builders, ordered, issues, manual, budget
    If ordered.Count > 0 Then
        If Not loaded Is Nothing Then available = loaded.State = atEvalOK
        If available Then
            If loaded.Index.SourceId <> "SUPERSET" Then Err.Raise vbObjectError + 3010, , "REGISTRY_SOURCE_INVALID"
            budget.CheckList loaded.Index.IssueCount: budget.Step loaded.Index.IssueCount * 4
            Set sourceIssues = loaded.Index.DataIssues
            complete = loaded.Index.IssueCount = 0
            If Not complete Then
                manual = True
                AT_RegistryIssue issues, "REGISTRY_SOURCE_INCOMPLETE", "WARNING", "REGISTRY", "", "Некоторые IP-ячейки Superset не разобраны. Совпадения сохранены, отсутствие адресов не подтверждается. Проверьте проблемы источника."
            End If
        Else
            manual = True
            AT_RegistryIssue issues, "REGISTRY_SOURCE_UNAVAILABLE", "WARNING", "REGISTRY", "", "Нет доступной проверенной выгрузки Superset. Адреса заявки сохранены для повторной проверки."
        End If
    End If
    For i = 1 To ordered.Count
        Set item = ordered(i): budget.Step
        If available Then
            Set query = loaded.Index.Query(item.QueryText, budget)
            If query.State <> atEvalOK Then Err.Raise vbObjectError + 3010, , query.Diagnostics(1).Code
            Set raw = query.Matches
        Else
            Set raw = New ATRuntimeValue: raw.InitializeRows "SUPERSET_RAW", emptyRows
        End If
        links = links + raw.Count
        If links > 100000 Then Err.Raise vbObjectError + 3010, , "REGISTRY_LINK_LIMIT"
        item.Finish CStr(i), raw, complete, hostIncomplete, budget
        entries.Add item
        If hostIncomplete Then
            manual = True
            AT_RegistryIssue issues, "REGISTRY_HOSTNAME_UNCHECKED", "WARNING", "HOSTNAME", item.References.Row(1).CellRef("CELL_REF"), "Hostname в Superset не удалось проверить. Все IP-совпадения для " & item.Address & " сохранены."
        End If
        If item.Status = "NOT_FOUND" Then
            manual = True
            AT_RegistryIssue issues, "REGISTRY_NOT_FOUND", "WARNING", "REGISTRY", item.References.Row(1).CellRef("CELL_REF"), "IP/сеть " & item.Address & " не найдена в текущей выгрузке Superset. Требуется ручная проверка. Все ссылки на вхождения адреса сохранены в данных сводки."
        End If
    Next i
    status = "PASS": reason = ""
    If ordered.Count = 0 Or manual Then status = "SKIPPED": reason = "IP-поиск выполнен не полностью или нет числовых адресов для проверки. Требуется источник или ручная проверка."
    For Each issue In issues
        If issue.Severity = "ERROR" Then status = "FAIL"
    Next issue
    outcome.Initialize "NET.IP_REGISTRY", status, reason, manual: checks.Add outcome
    report.Initialize atEvalOK, "NETWORK_REGISTRY", issues, checks
    result.Initialize atEvalOK, entries, report, sourceIssues: Set Run = result: Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "REGISTRY_INPUT_INVALID")
    Set entries = New Collection: Set issues = New Collection: Set checks = New Collection
    Set report = New ATCheckReport: Set outcome = New ATCheckOutcome: Set result = New ATNetworkRegistryResult
    AT_RegistryIssue issues, code, "ERROR", "REGISTRY", "", "Проверка Superset остановлена. Проверьте структуру заявки, источник и объём данных; частичные результаты не выданы."
    outcome.Initialize "NET.IP_REGISTRY", "BLOCKED", "IP-поиск не завершён.": checks.Add outcome
    report.Initialize atEvalError, "NETWORK_REGISTRY", issues, checks
    result.Initialize atEvalError, entries, report: Set Run = result
End Function
Private Sub Collect(ByVal context As ATEvaluationContext, ByVal lower As Boolean, ByVal builders As Object, ByVal ordered As Collection, ByVal issues As Collection, ByRef manual As Boolean, ByVal budget As ATEvalBudget)
    Dim rows As ATRuntimeValue, row As ATRuntimeRow, fields As Variant, field As Variant, text As String, host As String
    Dim i As Long, valid As Boolean, hostValid As Boolean, rowManual As Boolean, token As Variant, prefix As String, key As String
    Dim parsed As ATRuntimeValue, item As ATNetworkRegistryEntry, groupSeen As Boolean, candidate As Boolean
    If lower Then
        Set rows = AT_InitiatorRows(context.ResolveField("ADDRESS_ROWS_RAW"), "ADDRESS_TABLE", budget)
        fields = Array(Array("ADDRESS_RAW", "ADDRESS_IP"))
    Else
        Set rows = AT_InitiatorRows(context.ResolveField("TECHNICAL_ROWS_RAW"), "TECHNICAL_TABLE", budget)
        fields = Array(Array("SOURCE_RAW", "TECH_SOURCE_IP"), Array("DESTINATION_RAW", "TECH_DEST_IP"))
    End If
    For i = 1 To rows.Count
        Set row = rows.Row(i): host = ""
        If lower Then
            host = AT_RegistryText(row.ResolveField("HOSTNAME_RAW", budget), hostValid, budget)
            If Not hostValid Then
                manual = True
                AT_RegistryIssue issues, "REGISTRY_REQUEST_UNCHECKED", "WARNING", "HOSTNAME", row.CellRef("HOSTNAME_RAW"), "В ячейке hostname ошибка Excel. IP-совпадения будут сохранены без уточнения имени."
            End If
        End If
        For Each field In fields
            text = AT_RegistryText(row.ResolveField(CStr(field(0)), budget), valid, budget)
            rowManual = Not valid: prefix = "": groupSeen = False
            If valid Then
                text = Replace(Replace(text, "х", "x"), "Х", "x")
                For Each token In AT_DomainTokens(text, False, budget)
                    budget.Step
                    candidate = AT_RegistryCandidate(CStr(token))
                    If groupSeen And Len(prefix) = 0 Then
                        If CStr(token) Like String$(Len(CStr(token)), "#") Then candidate = False
                    End If
                    If candidate Then
                        Set parsed = AT_RegistryToken(CStr(token), prefix, rowManual, budget)
                        If Not parsed Is Nothing Then
                            key = AT_RegistryKey(parsed.Domain, budget)
                            If Not builders.Exists(key) Then
                                budget.CheckList ordered.Count + 1
                                Set item = New ATNetworkRegistryEntry: item.Initialize AT_RegistryDisplay(CStr(token), parsed), key, parsed.Domain
                                builders.Add key, item: ordered.Add item
                            Else
                                Set item = builders(key)
                            End If
                            item.AddReference CStr(field(1)), row.CellRef(CStr(field(0))), budget
                            If lower Then item.AddHostname host, budget, hostValid
                        End If
                    Else
                        prefix = "": groupSeen = True
                    End If
                Next token
            End If
            If rowManual Then
                manual = True
                AT_RegistryIssue issues, "REGISTRY_REQUEST_UNCHECKED", "WARNING", CStr(field(1)), row.CellRef(CStr(field(0))), "Некоторые адреса в ячейке не разобраны или содержат ошибку Excel. Корректные отдельные объявления сохранены; требуется ручная проверка."
            End If
        Next field
    Next i
End Sub
