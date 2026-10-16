Attribute VB_Name = "modATSupersetRuleTests"
Option Explicit
Public Sub AT_TestSupersetRules()
    Dim book As Workbook, ws As Worksheet, reader As New ATNetworkRequestReader, planner As New ATNetworkRegistryPlanner
    Dim compiler As New ATIPSourceCompiler, loaded As ATIPSourceResult, registry As ATNetworkRegistryResult
    Dim ruleCompiler As New ATNetworkRulebookCompiler, rules As ATNetworkRulebookResult, runner As New ATNetworkSupersetRules
    Dim result As ATNetworkSupersetResult, rows As ATRuntimeValue, records As Variant, definitions As Variant
    Dim number As Long, description As String, writer As New ATCoordinatorWorkbookWriter, output As Workbook
    Dim budget As ATEvalBudget, other As ATNetworkRegistryResult, caught As Boolean, code As String, books As Long
    On Error GoTo Failed
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    ws.Range("A15:D15").Value2 = Array("192.0.2.0/24", "203.0.113.1", "TCP", "443")
    ws.Range("B17").Value2 = "192.0.2.1": ws.Range("C17").Value2 = "web.test"
    ws.Range("B18").Value2 = "203.0.113.1"
    records = Array(Array("web.test", "192.0.2.1", "", "", "IS 1", "Original service", "G1", "T1"), Array("db.test", "192.0.2.2", "", "", "IS 2", "Other service", "G2", "T2"))
    Set loaded = compiler.Compile(AT_TestRegistryRaw(records), "SUPERSET", AT_TestRegistryFields(), "rules fixture")
    Set registry = planner.Run(reader.Read(ws, "RULES", "2026-10-16", "UTC"), loaded)
    definitions = Array(AT_TestSupersetRule("web exact", "192.0.2.1", "web.test", "Original service", "Web output", "Use web"), AT_TestSupersetRule("db", "192.0.2.2", "", "Other service", "DB output", "Use DB"), AT_TestSupersetRule("false cross row", "192.0.2.1", "db.test", "Other service", "WRONG", "WRONG"), AT_TestSupersetRule("no chain", "", "", "Web output", "CHAINED", "WRONG"))
    Set rules = ruleCompiler.Compile(AT_TestRulesRaw(definitions), AT_TestRuleProtocols())
    Set result = runner.Run(registry, rules)
    AT_Assert Not result Is Nothing, "Superset rule runner returns bound annotations"
    If result Is Nothing Then GoTo Cleanup
    Set rows = result.Annotations
    AT_Assert result.State = atEvalOK And rows.Count = 3, "one annotation per unique registry entry including no hit"
    AT_Assert AT_ViewText(rows.Row(1).ResolveField("SERVICE_RESULT")) = "Web output" & vbLf & "DB output", "different actual services retained without false conflict"
    AT_Assert InStr(AT_ViewText(rows.Row(1).ResolveField("SERVICE_RESULT")), "WRONG") = 0, "network context uses intersection with each source row not entire query"
    AT_Assert InStr(AT_ViewText(rows.Row(1).ResolveField("SERVICE_RESULT")), "CHAINED") = 0, "rule output never becomes input service of another rule"
    AT_Assert AT_ViewText(rows.Row(3).ResolveField("SERVICE_RESULT")) = "Web output", "singleton lower IP receives exact hostname/service rule"
    AT_Assert Len(AT_ViewText(rows.Row(2).ResolveField("SERVICE_RESULT"))) = 0 And InStr(AT_ViewText(rows.Row(2).ResolveField("INSTRUCTION")), "не вычислены") > 0, "no source rows never invent service context"
    AT_Assert result.Report.Checks(1).CheckId = "NET.HOSTNAME_RULES" And result.Report.Checks(1).Incomplete, "new independent outcome records no-hit incompleteness"
    AT_Assert result.Report.ErrorCount = 0, "different source row services are not a rule conflict"
    AT_Assert InStr(AT_ViewText(rows.Row(1).ResolveField("INSTRUCTION")), "Use web") > 0 And InStr(AT_ViewText(rows.Row(1).ResolveField("INSTRUCTION")), "'Правила и исключения'!B2") > 0, "instructions retain rule source reference"
    Set output = writer.Create(registry, result)
    AT_Assert output.Worksheets(1).Cells(6, 5).Value2 = "Web output" & vbLf & "DB output" And InStr(output.Worksheets(1).Cells(6, 6).Value2, "Use DB") > 0, "writer fills coordinator service and instructions from bound result"
    AT_Assert output.Worksheets(2).Cells(6, 7).Value2 = "Original service" And output.Worksheets(1).Cells(6, 8).Value2 = "Original service" & vbLf & "Other service", "annotations never overwrite original Superset services or enrichment"
    AT_Assert InStr(output.Worksheets(1).Cells(3, 1).Value2, "ещё не применены") = 0, "writer banner reflects supplied Superset annotations"
    output.Close False: Set output = Nothing
    AT_TestSupersetComposition reader.Read(ws, "COMPOSE", "2026-10-16", "UTC"), registry, rules, result
    Set other = planner.Run(reader.Read(ws, "OTHER", "2026-10-16", "UTC"), loaded): books = Application.Workbooks.Count
    On Error Resume Next
    Set output = writer.Create(other, result): caught = Err.Number <> 0: code = Err.Description: Err.Clear
    On Error GoTo Failed
    AT_Assert caught And code = "SUPERSET_RULES_BINDING" And Application.Workbooks.Count = books, "annotations cannot bind to another registry even with identical addresses"
    If Not output Is Nothing Then output.Close False: Set output = Nothing
    Set budget = AT_TestInitiatorBudget(): budget.Step 200000
    Set result = runner.Run(registry, rules, budget)
    AT_Assert result.State = atEvalError And result.Annotations.Count = 0 And result.Report.Checks(1).Status = "BLOCKED", "budget failure discards all new partial annotations"
    Set output = writer.Create(registry, result)
    AT_Assert output.Worksheets(1).Cells(6, 2).Value2 = "192.0.2.0/24" And InStr(output.Worksheets(1).Cells(6, 6).Value2, "не завершены") > 0, "blocked annotations retain base workbook with explicit not-checked instructions"
    output.Close False: Set output = Nothing
    definitions = Array(AT_TestSupersetRule("tie A", "192.0.2.1", "web.test", "Original service", "A", "Instruction A"), AT_TestSupersetRule("tie B", "192.0.2.1", "WEB.TEST.", "Original service", "B", "Instruction B"), AT_TestSupersetRule("db", "192.0.2.2", "", "Other service", "DB", "Use DB"))
    Set rules = ruleCompiler.Compile(AT_TestRulesRaw(definitions), AT_TestRuleProtocols())
    Set result = runner.Run(registry, rules)
    AT_Assert result.Report.Checks(1).Status = "FAIL" And InStr(AT_ViewText(result.Annotations.Row(1).ResolveField("SERVICE_RESULT")), "B") > 0, "tied conflicting outputs kept with failure outcome"
    AT_Assert result.Report.ErrorCount = 2, "one conflict issue per actual entry-source context not per tied rule"
    records(0) = Array("", "192.0.2.1", "", "", "IS 1", "Original service", "", "")
    Set loaded = compiler.Compile(AT_TestRegistryRaw(records), "SUPERSET", AT_TestRegistryFields(), "missing host")
    Set other = planner.Run(reader.Read(ws, "MISSING_HOST", "2026-10-16", "UTC"), loaded)
    Set result = runner.Run(other, rules)
    AT_Assert result.State = atEvalOK And AT_ViewText(result.Annotations.Row(1).ResolveField("SERVICE_RESULT")) = "DB" And result.Report.Checks(1).Status = "SKIPPED", "missing hostname does not hide confirmed rule results of independent source rows"
    records(0) = Array("web.test", "192.0.2.1,bad", "192.0.2.2", "", "IS 1", "Original service", "", "")
    Set loaded = compiler.Compile(AT_TestRegistryRaw(records), "SUPERSET", AT_TestRegistryFields(), "strict source IP")
    Set other = planner.Run(reader.Read(ws, "STRICT_IP", "2026-10-16", "UTC"), loaded)
    Set result = runner.Run(other, rules)
    AT_Assert result.State = atEvalOK And AT_ViewText(result.Annotations.Row(1).ResolveField("SERVICE_RESULT")) = "DB" And result.Report.ErrorCount = 0, "invalid whole source IP cell cannot leak salvaged address into rule context"
    Set result = runner.Run(registry, Nothing)
    AT_Assert result.State = atEvalOK And result.Report.Checks(1).Status = "SKIPPED" And result.Annotations.Count = 3, "missing rulebook retains all entries with explicit skipped annotations"
Cleanup:
    If Not output Is Nothing Then output.Close False
    book.Close False: Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not output Is Nothing Then output.Close False
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise number, "AT_TestSupersetRules", description
End Sub
Private Sub AT_TestSupersetComposition(ByVal source As ATNetworkReadResult, ByVal registry As ATNetworkRegistryResult, ByVal rules As ATNetworkRulebookResult, ByVal annotations As ATNetworkSupersetResult)
    Dim runner As New ATNetworkAccessChecks, catalogCompiler As New ATNetworkPolicyCompiler, actorCompiler As New ATInitiatorPolicyCompiler
    Dim catalog As ATNetworkPolicyResult, actor As ATInitiatorPolicyResult, previous As ATNetworkPolicyReport, result As ATNetworkPolicyReport, i As Long
    Set catalog = catalogCompiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP"))), AT_TestPolicyRaw("PROBLEMS", Array()), AT_TestPolicyRaw("ROUTING", Array()), AT_TestPolicyRaw("COMMON_TEXT", Array()))
    Set actor = actorCompiler.Compile(AT_TestInitiatorRaw("INITIATORS", Array()), AT_TestInitiatorRaw("INITIATOR_IP", Array()), "ALLOW")
    Set previous = runner.ApplyRegistry(runner.RunWithRules(source, catalog, actor, rules), registry, catalog)
    Set result = runner.ApplySupersetRules(previous, annotations, catalog)
    AT_Assert previous.AllReport.Checks.Count = 15 And result.AllReport.Checks.Count = 16, "Superset rules independently append sixteenth outcome"
    For i = 1 To 15
        AT_Assert result.AllReport.Checks(i).Status = previous.AllReport.Checks(i).Status And result.AllReport.Checks(i).CheckId = previous.AllReport.Checks(i).CheckId, "finished fifteen outcomes preserved"
    Next i
    AT_Assert result.UserErrorCount = previous.UserErrorCount, "Superset rule outcomes are not user-input errors"
    Set catalog = catalogCompiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP"))), AT_TestPolicyRaw("PROBLEMS", Array(Array("Да", "NET.HOSTNAME_RULES|INFO|Имя хоста", "Ошибка", "Один раз: {message}", "", "Нет", "Нет"))), AT_TestPolicyRaw("ROUTING", Array()), AT_TestPolicyRaw("COMMON_TEXT", Array()))
    Set result = runner.ApplySupersetRules(previous, annotations, catalog)
    If result.AllReport.Checks.Count = 16 Then
        AT_Assert result.AllReport.Checks(16).Status = "FAIL" And result.UserErrorCount = previous.UserErrorCount, "configured Superset error wins over manual without blaming input"
        AT_Assert InStr(result.SupersetIssues(result.SupersetIssues.Count).Message, "Один раз:") = 1 And InStr(2, result.SupersetIssues(result.SupersetIssues.Count).Message, "Один раз:") = 0, "published hostname key applies message template exactly once"
    End If
    Set catalog = catalogCompiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP"))), AT_TestPolicyRaw("PROBLEMS", Array(Array("Да", "NET.HOSTNAME_RULES|INFO|Имя хоста", "Информация", Replace(String$(300, "x"), "x", "{message}"), "", "Нет", "Нет"))), AT_TestPolicyRaw("ROUTING", Array()), AT_TestPolicyRaw("COMMON_TEXT", Array()))
    Set result = runner.ApplySupersetRules(previous, annotations, catalog)
    AT_Assert result.AllReport.Checks.Count = 16, "Superset template overflow preserves all fifteen outcomes"
    If result.AllReport.Checks.Count = 16 Then AT_Assert result.AllReport.Checks(16).Status = "BLOCKED", "only Superset rule presentation blocked by oversized template"
End Sub
Public Function AT_TestSupersetRule(ByVal name As String, ByVal ip As String, ByVal host As String, ByVal condition As String, ByVal service As String, ByVal instruction As String) As Variant
    AT_TestSupersetRule = Array("Да", name, ip, "", host, "", "", "", condition, service, "", "", "", "Информация", instruction)
End Function
