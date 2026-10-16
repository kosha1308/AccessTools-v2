Attribute VB_Name = "modATRegistryPlanTests"
Option Explicit
Public Sub AT_TestRegistryPlan()
    Dim book As Workbook, ws As Worksheet, reader As New ATNetworkRequestReader, source As ATNetworkReadResult
    Dim planner As New ATNetworkRegistryPlanner, result As ATNetworkRegistryResult, loaded As ATIPSourceResult
    Dim item As ATNetworkRegistryEntry, records As Variant, indexCompiler As New ATIPSourceCompiler, fields As Collection
    Dim issue As ATCheckIssue, found As Boolean, n As Long, description As String, caught As Boolean
    Dim budget As ATEvalBudget, view As Collection, groupLabel As Variant
    On Error GoTo Failed
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    ws.Range("A15:D15").Value2 = Array("Group Alpha", "192.0.2.1", "TCP", "443"): ws.Range("E15").Value2 = "Причина"
    ws.Range("A17").Value2 = "ИС": ws.Range("B17").Value2 = "192.0.2.2": ws.Range("C17").Value2 = "lower.test": ws.Range("E17").Value2 = "Источник"
    ws.Range("A18").Value2 = "ИС": ws.Range("B18").Value2 = "192.0.2.3": ws.Range("C18").Value2 = "missing.test": ws.Range("E18").Value2 = "Назначение"
    records = Array(Array("upper.test", "192.0.2.1", "", "", "IS1", "Service1", "SKUF1", "Team1"), Array("LOWER.TEST.", "192.0.2.2", "", "", "IS2", "Service2", "SKUF2", "Team2"), Array("other.test", "192.0.2.2", "", "", "Other IS", "Other Service", "", ""))
    Set fields = AT_TestRegistryFields()
    Set loaded = indexCompiler.Compile(AT_TestRegistryRaw(records), "SUPERSET", fields, "registry fixture")
    Set source = reader.Read(ws, "REGISTRY", "2026-10-16", "UTC")
    Set result = planner.Run(source, loaded)
    AT_Assert result.State = atEvalOK And result.Entries.Count = 3, "group plus lower actual IPs yields complete effective address set"
    If result.State <> atEvalOK Or result.Entries.Count <> 3 Then GoTo Cleanup
    Set item = result.Entries(1)
    AT_Assert item.Id = "1" And item.Address = "192.0.2.1" And item.Status = "FOUND", "stable IDs follow first upper then lower occurrence"
    AT_Assert item.References.Row(1).CellRef("CELL_REF") = "'Лист1'!B15", "upper request reference preserved"
    Set item = result.Entries(2)
    AT_Assert item.Address = "192.0.2.2" And item.Matches.Count = 1 And item.HostnameRefined, "lower IP participates and exact hostname narrows IP matches"
    AT_Assert AT_PolicyCell(item.Matches.Row(1), "SERVICE", AT_TestInitiatorBudget()) = "Service2", "actual refined row payload retained"
    AT_Assert result.Entries(3).Status = "NOT_FOUND" And result.Entries(3).Matches.Count = 0, "unmatched lower IP retained as summary entry"
    AT_Assert result.Report.Checks(1).Status = "SKIPPED" And result.Report.WarningCount = 1, "published not-found warning requires manual check without group-only noise"
    AT_Assert result.Report.Issues(1).CellRef = "'Лист1'!B18" And result.Report.Issues(1).Code = "REGISTRY_NOT_FOUND", "unique not-found issue links actual lower cell"
    For Each groupLabel In Array("Group 42", "Группа x")
        ws.Range("A15").Value2 = groupLabel
        Set result = planner.Run(reader.Read(ws, "GROUP_LABEL", "2026-10-16", "UTC"), loaded)
        AT_Assert result.Report.Issues.Count = 1 And result.Entries.Count = 3, "text group labels do not create spurious numeric fragment warnings"
    Next groupLabel
    ws.Range("A15").NumberFormat = "@": ws.Range("A15").Value2 = "192.0.2.1,2"
    ws.Range("B17").Value2 = "192.0.2.2/32"
    Set result = planner.Run(reader.Read(ws, "DEDUP", "2026-10-16", "UTC"), loaded)
    AT_Assert result.Entries.Count = 3 And result.Entries(1).References.Count = 2 And result.Entries(2).References.Count = 2, "compact values, upper/lower and /32 deduplicated without losing refs"
    AT_Assert result.Report.Issues.Count = 1, "repeated address produces one registry lookup/issue"
    ws.Range("B18").Value2 = "192.0.2.2": ws.Range("C18").Value2 = "lower.test"
    ws.Range("C17").Value2 = CVErr(xlErrValue)
    Set result = planner.Run(reader.Read(ws, "REQUEST_HOST_ERROR", "2026-10-16", "UTC"), loaded)
    AT_Assert result.Entries(2).Matches.Count = 2 And Not result.Entries(2).HostnameRefined, "duplicate request hostname error prevents narrowing by another row"
    ws.Range("B18").Value2 = "192.0.2.3": ws.Range("C18").Value2 = "missing.test"
    ws.Range("C17").Value2 = "not-found-host.test"
    Set result = planner.Run(reader.Read(ws, "HOST_FALLBACK", "2026-10-16", "UTC"), loaded)
    AT_Assert result.Entries(2).Matches.Count = 2 And Not result.Entries(2).HostnameRefined, "absent exact hostname preserves IP matches and services"
    ws.Range("C17").Value2 = "lower.test"
    ws.Range("A15").Value2 = "192.0.2.x": ws.Range("B17").Value2 = "192.0.2.0/24"
    Set result = planner.Run(reader.Read(ws, "NETWORK", "2026-10-16", "UTC"), loaded)
    AT_Assert result.Entries.Count = 3 And result.Entries(1).References.Count = 2 And result.Entries(1).Domain.Count = 1, "mathematically equal mask/network dedup, no IP expansion"
    AT_Assert result.Entries(1).Matches.Count = 3 And Not result.Entries(1).HostnameRefined, "one hostname on a network cannot hide matches for other IPs"
    ws.Range("A15").Value2 = "192.0.2.1,999.1.1.1,2": ws.Range("B17").ClearContents
    Set result = planner.Run(reader.Read(ws, "BAD_PREFIX", "2026-10-16", "UTC"), loaded)
    AT_Assert result.Entries.Count = 2 And result.Entries(2).Address = "192.0.2.3", "bad numeric fragment resets compact prefix while preserving independent valid IPs"
    found = False
    For Each issue In result.Report.Issues
        If issue.Code = "REGISTRY_REQUEST_UNCHECKED" And issue.CellRef = "'Лист1'!A15" Then found = True
    Next issue
    AT_Assert found And result.Report.Checks(1).Incomplete, "invalid fragment remains explicit incomplete with exact source ref"
    ws.Range("A15").Value2 = "Group Alpha": ws.Range("B17").Value2 = "192.0.2.2"
    Set loaded = indexCompiler.Compile(AT_TestRegistryRaw(Array(records(0), records(1), Array("bad", "192.0.2.3;bad", "", "", "", "", "", ""))), "SUPERSET", fields, "incomplete source")
    Set result = planner.Run(reader.Read(ws, "INCOMPLETE_SOURCE", "2026-10-16", "UTC"), loaded)
    AT_Assert result.Entries(1).Status = "FOUND" And result.Entries(3).Status = "UNCONFIRMED" And Not result.Entries(3).SourceComplete, "incomplete source cannot claim definitive missing IP"
    AT_Assert result.SourceIssues.Value.Count = 1 And result.Report.Issues.Count = 1, "source defects preserved with one aggregate warning"
    Set result = planner.Run(source, Nothing)
    AT_Assert result.State = atEvalOK And result.Entries.Count = 3 And result.Entries(1).Status = "UNCONFIRMED" And result.Report.Checks(1).Status = "SKIPPED", "unavailable source retains all request entries without false absence"
    Set loaded = indexCompiler.Compile(AT_TestRegistryRaw(Array(records(0), records(1), Array(CVErr(xlErrValue), "192.0.2.2", "", "", "", "", "", ""))), "SUPERSET", fields, "bad host")
    Set result = planner.Run(source, loaded)
    AT_Assert result.Entries(2).Matches.Count = 2 And Not result.Entries(2).HostnameRefined And Not result.Entries(2).SourceComplete, "hostname Excel error does not hide IP matches"
    ws.Range("C17").Value2 = "СЕРВЕР.TEST."
    Set loaded = indexCompiler.Compile(AT_TestRegistryRaw(Array(records(0), Array("сервер.test", "192.0.2.2", "", "", "", "", "", ""), records(2))), "SUPERSET", fields, "unicode host")
    Set result = planner.Run(reader.Read(ws, "HOST_CASE", "2026-10-16", "UTC"), loaded)
    AT_Assert result.Entries(2).Matches.Count = 1 And result.Entries(2).HostnameRefined, "hostname case-insensitive normalization consistent for Cyrillic labels"
    Set view = result.Entries: view.Remove 1
    AT_Assert result.Entries.Count = 3, "registry entry collection defensive snapshot"
    Set item = result.Entries(1)
    On Error Resume Next
    item.AddReference "ADDRESS_IP", "'Лист1'!Z9", AT_TestInitiatorBudget()
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    AT_Assert caught, "published registry entry frozen"
    AT_TestRegistryComposition source, result, ws
    Set budget = AT_TestInitiatorBudget(): budget.Step 200000
    Set result = planner.Run(source, loaded, budget)
    AT_Assert result.State = atEvalError And result.Entries.Count = 0 And result.Report.Checks(1).Status = "BLOCKED", "exhausted caller budget cannot leak partial registry entries"
    Set result = planner.Run(Nothing, loaded)
    AT_Assert result.State = atEvalError And result.Entries.Count = 0, "bad request snapshot gives controlled blocked result"
    ws.Range("A15:D15").ClearContents: ws.Range("E15").MergeArea.ClearContents
    ws.Range("B17").ClearContents: ws.Range("B18").ClearContents
    ws.Range("C17").MergeArea.ClearContents: ws.Range("C18").MergeArea.ClearContents
    Set result = planner.Run(reader.Read(ws, "NO_ADDRESSES", "2026-10-16", "UTC"), Nothing)
    AT_Assert result.State = atEvalOK And result.Entries.Count = 0 And result.Report.Issues.Count = 0 And result.Report.Checks(1).Status = "SKIPPED", "no source dependency or pointless warning when there are no IPs"
Cleanup:
    book.Close False: Set book = Nothing: Exit Sub
Failed:
    n = Err.Number: description = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise n, "AT_TestRegistryPlan", description
End Sub
Private Sub AT_TestRegistryComposition(ByVal source As ATNetworkReadResult, ByVal registry As ATNetworkRegistryResult, ByVal ws As Worksheet)
    Dim runner As New ATNetworkAccessChecks, catalogCompiler As New ATNetworkPolicyCompiler, actorCompiler As New ATInitiatorPolicyCompiler
    Dim ruleCompiler As New ATNetworkRulebookCompiler, catalog As ATNetworkPolicyResult, actor As ATInitiatorPolicyResult
    Dim rules As ATNetworkRulebookResult, old As ATNetworkPolicyReport, added As ATNetworkPolicyReport
    Dim writer As New ATExcelReportWriter, output As Workbook, i As Long, n As Long, description As String
    On Error GoTo Failed
    Set catalog = catalogCompiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP"))), AT_TestPolicyRaw("PROBLEMS", Array()), AT_TestPolicyRaw("ROUTING", Array()), AT_TestPolicyRaw("COMMON_TEXT", Array()))
    Set actor = actorCompiler.Compile(AT_TestInitiatorRaw("INITIATORS", Array()), AT_TestInitiatorRaw("INITIATOR_IP", Array()), "ALLOW")
    Set rules = ruleCompiler.Compile(AT_TestRulesRaw(Array()), catalog.Policy.Protocols)
    Set old = runner.RunWithRules(source, catalog, actor, rules)
    Set added = runner.ApplyRegistry(old, registry, catalog)
    AT_Assert added.AllReport.Checks.Count = 15 And old.AllReport.Checks.Count = 14, "registry composes fifteenth check without mutating finished fourteen"
    For i = 1 To 14
        AT_Assert old.AllReport.Checks(i).CheckId = added.AllReport.Checks(i).CheckId And old.AllReport.Checks(i).Status = added.AllReport.Checks(i).Status, "previous outcome preserved when registry appended"
    Next i
    AT_Assert added.UserErrorCount = old.UserErrorCount, "registry manual/source problems not user-input errors"
    Set output = writer.Create(added.AllReport, ws)
    AT_Assert InStr(output.Worksheets("Отчёт").Range("A3").Value2, "ещё не") > 0 And InStr(output.Worksheets("Отчёт").Range("A3").Value2, "не выполнялись") = 0, "registry report banner distinguishes IP search from pending final view"
    AT_Assert output.Worksheets("Проверки").Range("A19").Value2 = "IP в Superset", "registry outcome translated in Excel"
    output.Close False: Set output = Nothing
    Set catalog = catalogCompiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP"))), AT_TestPolicyRaw("PROBLEMS", Array(Array("Да", "NET.IP_REGISTRY|WARNING|Проверка IP", "Ошибка", "Один раз: {message}", "", "Нет", "Нет"))), AT_TestPolicyRaw("ROUTING", Array()), AT_TestPolicyRaw("COMMON_TEXT", Array()))
    Set added = runner.ApplyRegistry(old, registry, catalog)
    AT_Assert added.AllReport.Checks(15).Status = "FAIL" And added.UserErrorCount = old.UserErrorCount, "configured registry error wins over incomplete without blaming request input"
    AT_Assert InStr(added.SupersetIssues(added.SupersetIssues.Count).Message, "Один раз: ") = 1 And InStr(2, added.SupersetIssues(added.SupersetIssues.Count).Message, "Один раз: ") = 0, "registry message template applied once"
    Set catalog = catalogCompiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP"))), AT_TestPolicyRaw("PROBLEMS", Array(Array("Да", "NET.IP_REGISTRY|WARNING|Проверка IP", "Предупреждение", Replace(String$(300, "x"), "x", "{message}"), "", "Нет", "Нет"))), AT_TestPolicyRaw("ROUTING", Array()), AT_TestPolicyRaw("COMMON_TEXT", Array()))
    Set added = runner.ApplyRegistry(old, registry, catalog)
    AT_Assert added.AllReport.Checks.Count = 15 And added.AllReport.Checks(15).Status = "BLOCKED", "registry template overflow cannot erase completed fourteen outcomes"
    Set output = writer.Create(registry.Report, ws)
    AT_Assert InStr(output.Worksheets("Отчёт").Range("A3").Value2, "IP-поиска") > 0, "standalone registry report banner also reflects performed lookup"
    output.Close False: Set output = Nothing: Exit Sub
Failed:
    n = Err.Number: description = Err.Description
    On Error Resume Next
    If Not output Is Nothing Then output.Close False
    On Error GoTo 0
    Err.Raise n, "AT_TestRegistryComposition", description
End Sub
Public Function AT_TestRegistryFields() As Collection
    Dim r As New Collection
    r.Add "SOURCE_TELEGRAPH": r.Add "SOURCE_CXD": r.Add "SOURCE_AWX": Set AT_TestRegistryFields = r
End Function
Public Function AT_TestRegistryRaw(ByVal records As Variant) As ATEvalResult
    Dim row As ATRuntimeRow, rows As New Collection, fields As Variant, record As Variant, i As Long, number As Long
    Dim value As New ATRuntimeValue, field As ATEvalResult
    fields = Array("HOSTNAME", "SOURCE_TELEGRAPH", "SOURCE_CXD", "SOURCE_AWX", "IS_NAME", "SERVICE", "SKUF_GROUP", "SUPPORT_TEAM")
    For Each record In records
        number = number + 1: Set row = New ATRuntimeRow: row.Initialize "SUPERSET_RAW", CStr(number + 1)
        For i = 0 To 7
            If IsError(record(i)) Then
                Set field = AT_RuntimeFailure(atEvalError, "EXCEL_CELL_ERROR", Nothing)
            ElseIf Len(CStr(record(i))) = 0 Then
                Set field = AT_RuntimeOK(AT_RuntimeState("TEXT", atValueEmpty))
            Else
                Set field = AT_RuntimeOK(AT_RuntimeScalar("TEXT", CStr(record(i))))
            End If
            row.AddField CStr(fields(i)), field, "'Registry source'!" & Chr$(66 + i) & CStr(number + 1)
        Next i
        rows.Add row
    Next record
    value.InitializeRows "SUPERSET_RAW", rows: Set AT_TestRegistryRaw = AT_RuntimeOK(value)
End Function
