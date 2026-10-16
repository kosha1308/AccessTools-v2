Attribute VB_Name = "modATRuleCheckTests"
Option Explicit
Public Sub AT_TestRuleChecks()
    Dim book As Workbook, ws As Worksheet, output As Workbook, reader As New ATNetworkRequestReader, source As ATNetworkReadResult
    Dim compiler As New ATNetworkRulebookCompiler, loaded As ATNetworkRulebookResult, checker As New ATNetworkRuleChecks
    Dim catalogCompiler As New ATNetworkPolicyCompiler, catalog As ATNetworkPolicyResult, actorCompiler As New ATInitiatorPolicyCompiler
    Dim actor As ATInitiatorPolicyResult, runner As New ATNetworkAccessChecks, result As ATNetworkPolicyReport, report As ATCheckReport, overflowCatalog As ATNetworkPolicyResult
    Dim writer As New ATExcelReportWriter, raw As Variant, issue As ATCheckIssue, matched As Boolean, original As String
    Dim n As Long, description As String, value As Variant, bigRule As Variant, warningRule As Variant, upgradeCatalog As ATNetworkPolicyResult
    On Error GoTo Failed
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    ws.Range("A5").Value2 = "Инициатор": ws.Range("D5").Value2 = "Отдел": ws.Range("E5").Value2 = "Должность": ws.Range("F5").Value2 = "user@example.test"
    ws.Range("C7").Value2 = "Предоставить доступ": ws.Range("C8").Value2 = "Да": ws.Range("C11").Value2 = "Задача": ws.Range("C12").Value2 = "Бессрочно"
    ws.Range("A15:D15").Value2 = Array("192.0.2.1", "192.0.2.2", "TCP", "443"): ws.Range("E15").Value2 = "Причина"
    ws.Range("A17").Value2 = "ИС": ws.Range("B17").Value2 = "192.0.2.1": ws.Range("C17").Value2 = "host1.test": ws.Range("E17").Value2 = "Источник"
    ws.Range("A18").Value2 = "ИС": ws.Range("B18").Value2 = "192.0.2.2": ws.Range("C18").Value2 = "host2.test": ws.Range("E18").Value2 = "Назначение"
    raw = Array("Да", "Web", "192.0.2.2", "Назначение", "", "", "443", "TCP", "", "Service A", "Team A", "Владелец", "Комментарий", "Ошибка", "Сначала согласовать")
    Set loaded = compiler.Compile(AT_TestRulesRaw(Array(raw)), AT_TestRuleProtocols())
    Set source = reader.Read(ws, "RULES", "2026-10-16", "UTC")
    Set report = checker.Run(source, loaded)
    AT_Assert report.State = atEvalOK And report.Checks.Count = 1 And report.Checks(1).CheckId = "NET.IP_EXCEPTIONS", "one independent technical rules check"
    AT_Assert report.Checks(1).Status = "FAIL" And report.Issues.Count = 1, "error severity rule hit produces business FAIL"
    Set issue = report.Issues(1)
    AT_Assert issue.CellRef = "'Лист1'!A15" And issue.Code = "RULEBOOK_MATCH", "rule result points to actual request connection"
    AT_Assert InStr(issue.Message, "443") > 0 And InStr(issue.Message, "192.0.2.2") > 0 And InStr(issue.Message, "TCP") > 0, "rule message retains technical context"
    AT_Assert InStr(issue.Message, "Сначала согласовать") > 0 And InStr(issue.Message, "Комментарий") > 0 And InStr(issue.Message, "Team A") > 0 And InStr(issue.Message, "Владелец") > 0 And InStr(issue.Message, "'Правила и исключения'!B2") > 0, "all outputs and provenance included"
    Set actor = actorCompiler.Compile(AT_TestInitiatorRaw("INITIATORS", Array(Array("Да", "example.test", "", "", "", ""))), AT_TestInitiatorRaw("INITIATOR_IP", Array()), "ALLOW")
    Set catalog = catalogCompiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP"), Array("Да", "UDP"))), AT_TestPolicyRaw("PROBLEMS", Array()), AT_TestPolicyRaw("ROUTING", Array()), AT_TestPolicyRaw("COMMON_TEXT", Array()))
    Set result = runner.RunWithRules(source, catalog, actor, loaded)
    AT_Assert result.AllReport.State = atEvalOK And result.AllReport.Checks.Count = 14, "new runner adds rules while preserving thirteen outcomes"
    AT_Assert result.UserErrorCount = 0 And result.UserIssues.Count = 0 And result.SupersetIssues.Count = 1, "coordinator rule hit is not user-input error by fallback"
    Set result = runner.Run(source, catalog, actor)
    AT_Assert result.AllReport.Checks.Count = 13 And result.AllReport.Issues.Count = 0, "original thirteen-check API preserved"
    Set catalog = catalogCompiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP"), Array("Да", "UDP"))), AT_TestPolicyRaw("PROBLEMS", Array(Array("Да", "NET.IP_EXCEPTIONS|ERROR|Правило / исключение", "Предупреждение", "Один раз: {message}", "", "Нет", "Нет"))), AT_TestPolicyRaw("ROUTING", Array()), AT_TestPolicyRaw("COMMON_TEXT", Array()))
    Set result = runner.RunWithRules(source, catalog, actor, loaded): Set issue = result.AllReport.Issues(1)
    AT_Assert Left$(issue.Message, 10) = "Один раз: " And InStr(2, issue.Message, "Один раз: ") = 0, "merged presentation exactly once"
    AT_Assert issue.Severity = "WARNING" And result.AllReport.Checks(14).Status = "PASS", "configured severity recalculates new check"
    Set output = writer.Create(result.AllReport, ws)
    AT_Assert output.Worksheets("Отчёт").Range("C10").Value2 = "Правила и исключения" And output.Worksheets("Отчёт").Range("D10").Value2 = "Правило / исключение", "new rule check and field translated in Excel"
    AT_Assert InStr(output.Worksheets("Отчёт").Range("E10").Value2, "Сначала согласовать") > 0 And output.Worksheets("Отчёт").Hyperlinks.Count = 1, "rule instruction rendered and source link exists"
    AT_Assert ws.Range("A15").Value2 = "192.0.2.1" And book.Worksheets.Count = 1, "new rule writer path leaves request unchanged"
    output.Close False: Set output = Nothing
    ws.Range("D15").NumberFormat = "@": ws.Range("D15").Value2 = "443,9000"
    Set report = checker.Run(reader.Read(ws, "PARTIAL", "2026-10-16", "UTC"), loaded)
    AT_Assert report.Checks(1).Status = "PASS" And report.Issues.Count = 0, "partial port cannot trigger rule: " & report.Checks(1).Status & "/" & report.Issues.Count & "/" & ws.Range("D15").Value2
    ws.Range("D15").Value2 = "443"
    For Each value In Array("TCP;UDP", "tcp TCP", "TCP|UDP")
        ws.Range("C15").Value2 = value
        Set report = checker.Run(reader.Read(ws, "PROTOCOLS", "2026-10-16", "UTC"), loaded)
        AT_Assert report.Issues.Count = 1 And report.Checks(1).Status = "FAIL", "all literal request protocols evaluated, duplicate tokens not duplicated"
    Next value
    ws.Range("C15").Value2 = "TCP BOGUS"
    Set report = checker.Run(reader.Read(ws, "UNKNOWN", "2026-10-16", "UTC"), loaded)
    AT_Assert report.Checks(1).Status = "FAIL" And report.WarningCount = 1 And report.ErrorCount = 1, "unknown protocol does not hide confirmed rule hit or incomplete warning"
    Set result = runner.RunWithRules(reader.Read(ws, "DOWNGRADE_INCOMPLETE", "2026-10-16", "UTC"), catalog, actor, loaded)
    AT_Assert result.AllReport.Checks(14).Status = "SKIPPED", "downgraded error cannot turn incomplete rule evaluation into PASS"
    warningRule = raw: warningRule(13) = "Предупреждение"
    Set loaded = compiler.Compile(AT_TestRulesRaw(Array(warningRule)), AT_TestRuleProtocols())
    Set upgradeCatalog = catalogCompiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP"), Array("Да", "UDP"))), AT_TestPolicyRaw("PROBLEMS", Array(Array("Да", "NET.IP_EXCEPTIONS|WARNING|Правило / исключение", "Ошибка", "", "", "Нет", "Нет"))), AT_TestPolicyRaw("ROUTING", Array()), AT_TestPolicyRaw("COMMON_TEXT", Array()))
    Set result = runner.RunWithRules(reader.Read(ws, "UPGRADE_INCOMPLETE", "2026-10-16", "UTC"), upgradeCatalog, actor, loaded)
    AT_Assert result.AllReport.Checks(14).Status = "FAIL", "upgraded error takes priority over incomplete evaluation"
    Set loaded = compiler.Compile(AT_TestRulesRaw(Array(raw)), AT_TestRuleProtocols())
    ws.Range("C15").Value2 = "TCP"
    For Each value In Array("group A", CVErr(xlErrValue))
        ws.Range("A15").Value2 = value
        Set report = checker.Run(reader.Read(ws, "UNCHECKED", "2026-10-16", "UTC"), loaded)
        AT_Assert report.Checks(1).Status = "SKIPPED" And report.WarningCount = 1, "unresolved or Excel-error context explicitly incomplete"
    Next value
    ws.Range("A15").Value2 = "192.0.2.1"
    Set loaded = compiler.Compile(AT_TestRulesRaw(Array(raw, AT_TestHostRule("conflict", "", "", "Service B"))), AT_TestRuleProtocols())
    Set report = checker.Run(source, loaded): matched = False
    For Each issue In report.Issues
        If issue.Code = "RULEBOOK_CONFLICT" Then matched = True
    Next issue
    AT_Assert matched And report.Checks(1).Status = "FAIL", "conflict retained as explicit separate error"
    Set loaded = compiler.Compile(Nothing, AT_TestRuleProtocols())
    Set result = runner.RunWithRules(source, catalog, actor, loaded)
    AT_Assert result.AllReport.Checks.Count = 14 And result.AllReport.Checks(14).Status = "BLOCKED" And result.UserErrorCount = 0, "bad new rules do not erase completed thirteen or blame input"
    Set result = runner.RunWithRules(source, catalog, actor, Nothing)
    AT_Assert result.AllReport.Checks.Count = 14 And result.AllReport.Checks(14).Status = "SKIPPED", "missing new rulebook explicit skipped, thirteen preserved"
    bigRule = raw: bigRule(12) = String$(17000, "c")
    Set loaded = compiler.Compile(AT_TestRulesRaw(Array(bigRule)), AT_TestRuleProtocols())
    Set overflowCatalog = catalogCompiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP"), Array("Да", "UDP"))), AT_TestPolicyRaw("PROBLEMS", Array(Array("Да", "NET.IP_EXCEPTIONS|ERROR|Правило / исключение", "Ошибка", "{message}{message}", "", "Нет", "Нет"))), AT_TestPolicyRaw("ROUTING", Array()), AT_TestPolicyRaw("COMMON_TEXT", Array()))
    Set result = runner.RunWithRules(source, overflowCatalog, actor, loaded)
    AT_Assert result.AllReport.Checks.Count = 14 And result.AllReport.State = atEvalOK, "rule presentation overflow preserves thirteen completed outcomes"
    If result.AllReport.Checks.Count = 14 Then AT_Assert result.AllReport.Checks(14).Status = "BLOCKED", "rule presentation overflow explicitly blocked"
    bigRule = raw: bigRule(12) = String$(17000, "c"): bigRule(14) = String$(17000, "i")
    Set loaded = compiler.Compile(AT_TestRulesRaw(Array(raw, bigRule)), AT_TestRuleProtocols())
    AT_Assert loaded.State = atEvalOK, "individual fields within bounds compile before combined message limit"
    Set report = checker.Run(source, loaded)
    AT_Assert report.State = atEvalError And report.Issues.Count = 1 And report.Issues(1).Code = "RULEBOOK_MESSAGE_LIMIT", "combined message overflow removes partial hits only in new check"
    Set result = runner.RunWithRules(source, catalog, actor, loaded)
    AT_Assert result.AllReport.Checks.Count = 14 And result.AllReport.Checks(14).Status = "BLOCKED" And result.AllReport.ErrorCount = 1 And result.UserErrorCount = 0, "new rule limit preserves prior thirteen without partial rule hits"
    ws.Range("A15:D15").ClearContents: ws.Range("E15").MergeArea.ClearContents
    Set loaded = compiler.Compile(AT_TestRulesRaw(Array(raw)), AT_TestRuleProtocols())
    Set report = checker.Run(reader.Read(ws, "EMPTY_TECHNICAL", "2026-10-16", "UTC"), loaded)
    AT_Assert report.Checks(1).Status = "SKIPPED", "empty technical table does not report successful rule evaluation"
    book.Close False: Set book = Nothing
    Exit Sub
Failed:
    n = Err.Number: description = Err.Description
    On Error Resume Next
    If Not output Is Nothing Then output.Close False
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise n, "AT_TestRuleChecks", description
End Sub
