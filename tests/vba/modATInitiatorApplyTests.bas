Attribute VB_Name = "modATInitiatorApplyTests"
Option Explicit
Public Sub AT_TestInitiatorApply()
    Dim book As Workbook, ws As Worksheet, reader As New ATNetworkRequestReader, source As ATNetworkReadResult
    Dim actorCompiler As New ATInitiatorPolicyCompiler, actor As ATInitiatorPolicyResult, actorRunner As New ATNetworkInitiatorChecks
    Dim catalogCompiler As New ATNetworkPolicyCompiler, catalog As ATNetworkPolicyResult, runner As New ATNetworkAccessChecks
    Dim report As ATCheckReport, result As ATNetworkPolicyReport, identities As ATEvalResult, ip As ATEvalResult
    Dim issue As ATCheckIssue, found As Boolean, check As ATCheckOutcome, number As Long, description As String
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    ws.Range("A5").Value2 = "Пользователь": ws.Range("D5").Value2 = "Отдел": ws.Range("E5").Value2 = "Должность": ws.Range("F5").Value2 = "user@example.test"
    ws.Range("C7").Value2 = "Предоставить доступ": ws.Range("C8").Value2 = "Да": ws.Range("C11").Value2 = "Рабочая задача": ws.Range("C12").Value2 = "Бессрочно"
    ws.Range("A15:D15").Value2 = Array("10.0.0.1", "10.0.0.2", "TCP", "443"): ws.Range("E15").Value2 = "reason"
    ws.Range("A17").Value2 = "system": ws.Range("B17").Value2 = "10.0.0.1": ws.Range("C17").Value2 = "host1.example.test": ws.Range("E17").Value2 = "role"
    ws.Range("A18").Value2 = "system": ws.Range("B18").Value2 = "10.0.0.2": ws.Range("C18").Value2 = "host2.example.test": ws.Range("E18").Value2 = "role"
    Set source = reader.Read(ws, "ACTOR", "2026-10-16", "UTC")
    Set identities = AT_TestInitiatorRaw("INITIATORS", Array(Array("Да", "example.test", "Владелец", "Ошибка", "", "")))
    Set actor = actorCompiler.Compile(identities, AT_TestInitiatorRaw("INITIATOR_IP", Array()), AT_NETWORK_NO_IP_POLICY_ACTION)
    Set catalog = catalogCompiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP"))), AT_TestPolicyRaw("PROBLEMS", Array()), AT_TestPolicyRaw("ROUTING", Array()), AT_TestPolicyRaw("COMMON_TEXT", Array()))
    Set result = runner.Run(source, catalog, actor)
    AT_Assert result.AllReport.State = atEvalOK And result.AllReport.Checks.Count = 13, "ten local plus three independent actor checks"
    AT_Assert result.AllReport.ErrorCount = 1 And result.UserErrorCount = 0 And result.UserIssues.Count = 0, "approval error not a user-input error by fallback"
    AT_Assert result.RequestIssues.Count = 1 And result.SupersetIssues.Count = 0, "approval fallback visible only in request channel"
    Set catalog = catalogCompiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP"))), AT_TestPolicyRaw("PROBLEMS", Array(Array("Да", "NET.INITIATOR_APPROVAL_RULES|ERROR|Дополнительное согласование", "Ошибка", "Шаблон: {message}", "", "Нет", "Да"))), AT_TestPolicyRaw("ROUTING", Array()), AT_TestPolicyRaw("COMMON_TEXT", Array()))
    Set result = runner.Run(source, catalog, actor)
    AT_Assert InStr(result.AllReport.Issues(1).Message, "Шаблон: ") = 1 And InStr(2, result.AllReport.Issues(1).Message, "Шаблон: ") = 0, "presentation applied exactly once to merged checks"
    Set identities = AT_TestInitiatorRaw("INITIATORS", Array(Array("Да", "example.test", "", "", "", "")))
    Set ip = AT_TestInitiatorRaw("INITIATOR_IP", Array(Array("Да", "example.test", "Любой", "10.0.0.1", "", "", "")))
    Set actor = actorCompiler.Compile(identities, ip, "ALLOW"): Set report = actorRunner.Run(source, actor): found = False
    For Each issue In report.Issues
        If issue.Code = "INITIATOR_IP_DENIED" And issue.CellRef = "'Лист1'!B15" Then found = True
    Next issue
    AT_Assert found And report.Checks(3).Status = "FAIL", "denied destination points to actual technical cell"
    Set ip = AT_TestInitiatorRaw("INITIATOR_IP", Array(Array("Да", "example.test", "Любой", "0.0.0.0/0", "", "", "Запрещено")))
    Set actor = actorCompiler.Compile(identities, ip, "ALLOW")
    ws.Range("C10").Value2 = "test-account"
    Set report = actorRunner.Run(reader.Read(ws, "DB", "2026-10-16", "UTC"), actor): found = False
    For Each issue In report.Issues
        If issue.Code = "INITIATOR_DB_ACCOUNT_FORBIDDEN" And issue.CellRef = "'Лист1'!C10" Then found = True
    Next issue
    AT_Assert found, "account restriction targets value C10, not label C9"
    ws.Range("C10").MergeArea.ClearContents: ws.Range("A15").Value2 = "group-alpha"
    Set ip = AT_TestInitiatorRaw("INITIATOR_IP", Array(Array("Да", "example.test", "Любой", "10.0.0.2", "", "", "")))
    Set actor = actorCompiler.Compile(identities, ip, "ALLOW")
    Set report = actorRunner.Run(reader.Read(ws, "GROUP", "2026-10-16", "UTC"), actor): found = False
    For Each issue In report.Issues
        If issue.Code = "INITIATOR_IP_DENIED" And issue.CellRef = "'Лист1'!B17" Then found = True
    Next issue
    AT_Assert found And report.Checks(3).Status = "FAIL", "group does not hide denied actual lower IP"
    ws.Range("A15").Value2 = "10.0.0.1": ws.Range("B18").Value2 = "192.0.2.1"
    Set ip = AT_TestInitiatorRaw("INITIATOR_IP", Array(Array("Да", "example.test", "Источник", "10.0.0.0/8", "", "", "")))
    Set actor = actorCompiler.Compile(identities, ip, "ALLOW")
    Set report = actorRunner.Run(reader.Read(ws, "DIRECTION", "2026-10-16", "UTC"), actor)
    AT_Assert report.Checks(3).Status = "SKIPPED" And report.WarningCount > 0, "new lower IP has no guessed direction under directional policy"
    ws.Range("B18").Value2 = "10.0.0.2"
    Set report = actorRunner.Run(reader.Read(ws, "KNOWN_DIRECTION", "2026-10-16", "UTC"), actor)
    AT_Assert report.Checks(3).Status = "PASS", "lower IP already checked in upper rows needs no guessed binding"
    ws.Range("F5").ClearContents
    Set report = actorRunner.Run(reader.Read(ws, "EMPTY_EMAIL", "2026-10-16", "UTC"), actor)
    AT_Assert report.SkippedCount = 3 And report.Issues.Count = 0, "missing email skips actor checks without duplicate issues"
    Set actor = actorCompiler.Compile(Nothing, Nothing, "ALLOW")
    Set result = runner.Run(source, catalog, actor)
    AT_Assert result.AllReport.State = atEvalError And result.AllReport.Checks.Count = 1 And result.UserIssues.Count = 1, "invalid actor policy leaves no partial successes"
    book.Close False: Set book = Nothing
    Exit Sub
CleanupError:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise number, "AT_TestInitiatorApply", description
End Sub
