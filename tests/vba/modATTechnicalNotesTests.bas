Attribute VB_Name = "modATTechnicalNotesTests"
Option Explicit
Public Sub AT_TestTechnicalNotes()
    Dim book As Workbook, ws As Worksheet, reader As New ATNetworkRequestReader, source As ATNetworkReadResult
    Dim planner As New ATNetworkRegistryPlanner, registry As ATNetworkRegistryResult, evaluator As New ATCoordinatorTechnicalNotes
    Dim ruleCompiler As New ATNetworkRulebookCompiler, rules As ATNetworkRulebookResult, ruleRunner As New ATNetworkRuleChecks
    Dim presenter As New ATNetworkPolicyChecks, policy As ATNetworkPolicyReport, raw As ATCheckReport
    Dim result As ATCoordinatorNotesResult, rows As ATRuntimeValue, number As Long, description As String
    Dim writer As New ATCoordinatorWorkbookWriter, output As Workbook, budget As ATEvalBudget, other As ATNetworkReadResult
    Dim otherRegistry As ATNetworkRegistryResult, caught As Boolean, code As String, count As Long, key As String
    Dim issues As Collection, checks As Collection, issue As ATCheckIssue, outcome As ATCheckOutcome, globalPolicy As ATNetworkPolicyReport
    On Error GoTo Failed
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    ws.Range("A15:D15").Value2 = Array("192.0.2.0/24", "198.51.100.1", "TCP", "443")
    ws.Range("B17").Value2 = "192.0.2.42": ws.Range("B18").Value2 = "203.0.113.1"
    Set source = reader.Read(ws, "NOTES", "2026-10-17", "UTC"): Set registry = planner.Run(source, Nothing)
    Set rules = ruleCompiler.Compile(AT_TestRulesRaw(Array(Array("Да", "Техническое правило", "198.51.100.1", "Назначение", "", "", "443", "TCP", "", "Сервис назначения", "", "Согласование", "Комментарий", "Информация", "Проверить основание"))), AT_TestRuleProtocols())
    Set raw = ruleRunner.Run(source, rules): Set policy = presenter.Apply(raw, AT_TestTechnicalCatalog())
    Set result = evaluator.Run(source, registry, policy)
    AT_Assert Not result Is Nothing, "technical notes adapter returns result without rerunning checks"
    If result Is Nothing Then GoTo Cleanup
    Set rows = result.Rows
    AT_Assert result.State = atEvalOK And rows.Count = 4, "technical notes preserve every registry entry"
    AT_Assert InStr(AT_ViewText(rows.Row(1).ResolveField("INSTRUCTION")), "Проверить основание") > 0 And InStr(AT_ViewText(rows.Row(2).ResolveField("INSTRUCTION")), "Проверить основание") > 0, "connection instruction attached to both participating endpoints"
    AT_Assert InStr(AT_ViewText(rows.Row(3).ResolveField("INSTRUCTION")), "Проверить основание") > 0 And Len(AT_ViewText(rows.Row(4).ResolveField("INSTRUCTION"))) = 0, "lower contained address gets connection context but unrelated address does not"
    AT_Assert InStr(AT_ViewText(rows.Row(3).ResolveField("INSTRUCTION")), "443") > 0 And InStr(AT_ViewText(rows.Row(3).ResolveField("INSTRUCTION")), "198.51.100.1") > 0, "technical annotation retains full connection not bare service attribution"
    AT_Assert raw.Checks.Count = 1 And raw.Checks(1).Status = "PASS" And result.GlobalCount = 0, "finished technical check unchanged and all routed notes assigned"
    Set output = writer.Create(registry, Nothing, result)
    AT_Assert InStr(output.Worksheets(1).Cells(8, 6).Value2, "Проверить основание") > 0 And Len(output.Worksheets(1).Cells(8, 5).Value2) = 0, "writer appends technical connection notes without attributing service to endpoint"
    AT_Assert InStr(output.Worksheets(1).Cells(3, 1).Value2, "Технические") > 0, "writer banner reflects transferred technical notes"
    output.Close False: Set output = Nothing
    AT_TestNotesOverflow writer, registry, result
    Set other = reader.Read(ws, "ANOTHER", "2026-10-17", "UTC")
    Set otherRegistry = planner.Run(other, Nothing): count = Application.Workbooks.Count
    On Error Resume Next
    Set output = writer.Create(otherRegistry, Nothing, result): caught = Err.Number <> 0: code = Err.Description: Err.Clear
    On Error GoTo Failed
    AT_Assert caught And code = "COORDINATOR_NOTES_BINDING" And Application.Workbooks.Count = count, "technical notes cannot be displayed for another registry"
    If Not output Is Nothing Then output.Close False: Set output = Nothing
    Set result = evaluator.Run(other, registry, policy)
    AT_Assert result.State = atEvalError And result.ErrorCode = "COORDINATOR_NOTES_SOURCE" And result.Rows.Count = 0, "another request snapshot cannot be correlated with registry"
    Set budget = AT_TestInitiatorBudget(): budget.Step 200000
    Set result = evaluator.Run(source, registry, policy, budget)
    AT_Assert result.State = atEvalError And result.Rows.Count = 0, "exhausted budget never leaks partial technical notes"
    Set output = writer.Create(registry, Nothing, result)
    AT_Assert InStr(output.Worksheets(1).Cells(6, 6).Value2, "не перенесены") > 0 And output.Worksheets(1).Cells(6, 2).Value2 = "192.0.2.0/24", "notes failure preserves base workbook with explicit fallback"
    output.Close False: Set output = Nothing
    key = "NET.IP_EXCEPTIONS|INFO|Правило / исключение"
    Set policy = presenter.Apply(raw, AT_TestTechnicalCatalog(Array(Array("Да", key, "Информация", "Пользователь: {message}", "Координатор: {message}", "Нет", "Нет"))))
    Set result = evaluator.Run(source, registry, policy)
    description = AT_ViewText(result.Rows.Row(1).ResolveField("INSTRUCTION"))
    AT_Assert InStr(description, "Координатор:") > 0 And InStr(description, "Пользователь:") = 0 And InStr(InStr(description, "Координатор:") + 1, description, "Координатор:") = 0, "already formatted coordinator message used once without reapplying policy"
    Set policy = presenter.Apply(raw, AT_TestTechnicalCatalog(Array(), Array(Array("Да", key, "Нет", "Нет", "Нет"))))
    Set result = evaluator.Run(source, registry, policy)
    AT_Assert result.State = atEvalOK And Len(AT_ViewText(result.Rows.Row(1).ResolveField("INSTRUCTION"))) = 0, "disabled Superset route prevents technical note export"
    Set issues = New Collection: Set checks = New Collection: Set issue = New ATCheckIssue: Set outcome = New ATCheckOutcome
    issue.Initialize "NET.IP_EXCEPTIONS", "RULEBOOK_MATCH", "INFO", "RULE_RESULT", "", "Общая инструкция": issues.Add issue
    outcome.Initialize "NET.IP_EXCEPTIONS", "PASS", "": checks.Add outcome
    Set raw = New ATCheckReport: raw.Initialize atEvalOK, "NETWORK_RULES", issues, checks
    Set globalPolicy = presenter.Apply(raw, AT_TestTechnicalCatalog())
    Set result = evaluator.Run(source, registry, globalPolicy)
    AT_Assert result.GlobalCount = 1 And result.Incomplete And Len(AT_ViewText(result.Rows.Row(1).ResolveField("INSTRUCTION"))) = 0, "unbound note remains global instead of being assigned to arbitrary address"
    Set output = writer.Create(registry, Nothing, result)
    AT_Assert InStr(output.Worksheets(1).Cells(3, 1).Value2, "Непривязанных: 1") > 0, "global technical issue count points to existing report"
Cleanup:
    If Not output Is Nothing Then output.Close False
    book.Close False: Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not output Is Nothing Then output.Close False
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise number, "AT_TestTechnicalNotes", description
End Sub
Private Sub AT_TestNotesOverflow(ByVal writer As ATCoordinatorWorkbookWriter, ByVal registry As ATNetworkRegistryResult, ByVal notes As ATCoordinatorNotesResult)
    Dim annotations As New ATNetworkSupersetResult, rows As New Collection, row As ATRuntimeRow, entry As ATNetworkRegistryEntry
    Dim report As New ATCheckReport, issues As New Collection, checks As New Collection, output As Workbook
    For Each entry In registry.Entries
        Set row = New ATRuntimeRow: row.Initialize "SUPERSET_ANNOTATION", entry.Id
        row.AddField "SERVICE_RESULT", AT_RuntimeOK(AT_RuntimeScalar("TEXT", "Сервис Superset"))
        row.AddField "INSTRUCTION", AT_RuntimeOK(AT_RuntimeScalar("TEXT", String$(32767, "x")))
        row.AddField "STATUS", AT_RuntimeOK(AT_RuntimeScalar("TEXT", "PASS")): rows.Add row
    Next entry
    report.Initialize atEvalOK, "SUPERSET_RULES", issues, checks
    annotations.Initialize atEvalOK, registry, rows, report
    Set output = writer.Create(registry, annotations, notes)
    AT_Assert Len(output.Worksheets(1).Cells(6, 6).Value2) = 32767 And output.Worksheets(1).Cells(6, 5).Value2 = "Сервис Superset", "technical append overflow preserves complete existing Superset annotations"
    AT_Assert InStr(output.Worksheets(1).Cells(3, 1).Value2, "Не перенесено из-за длины: 3") > 0, "technical append overflow explicitly counted without truncating old instructions"
    output.Close False
End Sub
Private Function AT_TestTechnicalCatalog(Optional ByVal problems As Variant, Optional ByVal routing As Variant) As ATNetworkPolicyResult
    Dim compiler As New ATNetworkPolicyCompiler
    If IsMissing(problems) Then problems = Array()
    If IsMissing(routing) Then routing = Array()
    Set AT_TestTechnicalCatalog = compiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", "TCP"))), AT_TestPolicyRaw("PROBLEMS", problems), AT_TestPolicyRaw("ROUTING", routing), AT_TestPolicyRaw("COMMON_TEXT", Array()))
End Function
