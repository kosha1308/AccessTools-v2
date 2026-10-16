Attribute VB_Name = "modATPolicyApplyTests"
Option Explicit

Public Sub AT_TestPolicyApply()
    Dim book As Workbook, ws As Worksheet, reader As New ATNetworkRequestReader, source As ATNetworkReadResult
    Dim runner As New ATNetworkPolicyChecks, loaded As ATNetworkPolicyResult, compiler As New ATNetworkPolicyCompiler
    Dim result As ATNetworkPolicyReport, saved As ATNetworkPolicyReport, item As ATNetworkPolicyDecision
    Dim p As ATEvalResult, m As ATEvalResult, r As ATEvalResult, c As ATEvalResult, items As Collection
    Dim key As String, number As Long, description As String, found As Boolean, check As ATCheckOutcome
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    ws.Range("A5").Value2 = "Тестовый пользователь": ws.Range("D5").Value2 = "Отдел": ws.Range("E5").Value2 = "Должность"
    ws.Range("F5").Value2 = "user@example.test": ws.Range("C7").Value2 = "Предоставить доступ": ws.Range("C8").Value2 = "Да"
    ws.Range("C11").Value2 = "Рабочая задача": ws.Range("C12").Value2 = "Бессрочно"
    ws.Range("A15:D15").Value2 = Array("192.0.2.1", "192.0.2.1", "UDP", "443")
    ws.Range("E15").Value2 = "reason": ws.Range("A17").Value2 = "system"
    ws.Range("B17").Value2 = "192.0.2.1": ws.Range("C17").Value2 = "host.example.test": ws.Range("E17").Value2 = "role"
    Set source = reader.Read(ws, "POLICY", "2026-10-16", "UTC")
    key = "NET.TECHNICAL_TABLE|ERROR|Протокол назначения"
    Set p = AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP")))
    Set m = AT_TestPolicyRaw("PROBLEMS", Array(Array("Да", key, "Информация", "Нужен иной протокол", "Проверить координатору", "Нет", "Да")))
    Set r = AT_TestPolicyRaw("ROUTING", Array(Array("Да", key, "Нет", "Да", "Да")))
    Set c = AT_TestPolicyRaw("COMMON_TEXT", Array())
    Set loaded = compiler.Compile(p, m, r, c): Set result = runner.Run(source, loaded)
    AT_Assert result.AllReport.State = atEvalOK And result.AllReport.Checks.Count = 10, "policy wrapper keeps all ten completed checks"
    AT_Assert result.AllReport.Issues.Count = 1 And result.AllReport.ErrorCount = 0, "configured protocol severity applied"
    AT_Assert result.UserIssues.Count = 0 And result.UserErrorCount = 0, "hidden informational issue not a user error"
    AT_Assert result.RequestIssues.Count = 1 And result.SupersetIssues.Count = 1, "explicit routing channels retained"
    AT_Assert result.RequestIssues(1).Message = "Нужен иной протокол" And result.SupersetIssues(1).Message = "Проверить координатору", "separate user/coordinator messages"
    AT_Assert result.AllReport.Issues(1).CheckId = "NET.PROTOCOLS" And result.AllReport.Issues(1).CellRef = "'Лист1'!C15", "binding keeps original identity and ref"
    For Each check In result.AllReport.Checks
        If check.CheckId = "NET.PROTOCOLS" Then AT_Assert check.Status = "PASS", "outcome recalculated after severity change"
    Next check
    Set items = result.Decisions: items.Remove 1
    AT_Assert result.Decisions.Count = 1, "policy decision collection copied"
    Set saved = result
    Set m = AT_TestPolicyRaw("PROBLEMS", Array(Array("Да", key, "Ошибка", "", "", "Нет", "Нет")))
    Set r = AT_TestPolicyRaw("ROUTING", Array(Array("Да", key, "Да", "Да", "Нет")))
    Set loaded = compiler.Compile(p, m, r, c): Set result = runner.Run(source, loaded)
    AT_Assert result.AllReport.ErrorCount = 1 And result.UserIssues.Count = 1 And result.UserErrorCount = 0, "all errors and classified user errors remain distinct"
    AT_Assert result.RequestIssues.Count = 0, "request route also requires highlight permission"
    Set m = AT_TestPolicyRaw("PROBLEMS", Array())
    Set r = AT_TestPolicyRaw("ROUTING", Array())
    Set loaded = compiler.Compile(p, m, r, c): Set result = runner.Run(source, loaded)
    AT_Assert result.UserErrorCount = 1 And result.AllReport.ErrorCount = 1, "unconfigured issue preserves builtin behavior"
    Set loaded = compiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array()), m, r, c)
    Set result = runner.Run(source, loaded): found = False
    For Each check In result.AllReport.Checks
        If check.CheckId = "NET.PROTOCOLS" Then found = (check.Status = "SKIPPED")
    Next check
    AT_Assert found And result.AllReport.WarningCount = 1, "empty policy catalog explicitly skips protocol membership"
    Set loaded = compiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "UDP"))), m, r, c)
    Set result = runner.Run(source, loaded)
    AT_Assert result.AllReport.Issues.Count = 0 And result.AllReport.Checks.Count = 10, "changed active catalog used without stale cache"
    ws.Range("A1").Value2 = "wrong template"
    Set result = runner.Run(reader.Read(ws, "BAD", "2026-10-16", "UTC"), loaded)
    AT_Assert result.AllReport.State = atEvalError And result.UserIssues.Count > 0, "blocked source cannot be hidden by policy"
    AT_Assert result.AllReport.Checks(1).Status = "BLOCKED", "blocked execution status retained"
    Set loaded = compiler.Compile(Nothing, Nothing, Nothing, Nothing)
    Set result = runner.Run(source, loaded)
    AT_Assert result.AllReport.State = atEvalError And result.AllReport.Checks.Count = 1, "invalid policy publishes no partial successes"
    AT_Assert result.AllReport.Checks(1).CheckId = "NET.POLICY" And result.AllReport.Issues(1).Code = "POLICY_INPUT_INVALID", "invalid policy visible with stable reason"
    Set result = runner.Run(source, Nothing)
    AT_Assert result.AllReport.State = atEvalError And result.UserIssues.Count = 1, "missing compiled policy visible"
    book.Close False: Set book = Nothing
    AT_Assert saved.SupersetIssues(1).Message = "Проверить координатору", "policy output survives source closure"
    Exit Sub
CleanupError:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise number, "AT_TestPolicyApply", description
End Sub
