VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkInitiatorChecks"
Option Explicit
Public Function Run(ByVal source As ATNetworkReadResult, ByVal loaded As ATInitiatorPolicyResult) As ATCheckReport
    Dim budget As New ATEvalBudget, limits As New ATEvalLimits, context As ATEvaluationContext
    Dim selection As ATInitiatorSelection, identity As ATCheckReport, report As New ATCheckReport
    Dim issues As Collection, checks As Collection, outcome As ATCheckOutcome, issue As ATCheckIssue
    Dim rows As ATRuntimeValue, row As ATRuntimeRow, field As ATEvalResult, checkedUpper As New ATIntervalSet
    Dim email As String, text As String, manual As Boolean, hasInput As Boolean, status As String, i As Long, code As String
    On Error GoTo Failed
    If loaded Is Nothing Then Err.Raise vbObjectError + 2970, , "INITIATOR_INPUT_INVALID"
    If loaded.State <> atEvalOK Then Err.Raise vbObjectError + 2970, , loaded.Diagnostics(1).Code
    If source Is Nothing Then Err.Raise vbObjectError + 2970, , "INITIATOR_INPUT_INVALID"
    If source.State <> atEvalOK Then Err.Raise vbObjectError + 2970, , source.Diagnostics(1).Code
    budget.Initialize limits: Set context = source.Context: context.Validate
    Set field = context.ResolveField("EMAIL_RAW", budget)
    If Not RawText(field, email, budget) Then email = ""
    Set selection = loaded.Policy.ForEmail(email, budget)
    Set identity = selection.IdentityReport(context.CellRef("EMAIL_RAW")): Set issues = identity.Issues: Set checks = identity.Checks
    checkedUpper.Initialize 4294967295#
    status = "SKIPPED"
    If AT_EmailSyntaxValid(email) Then
        Set rows = AT_InitiatorRows(context.ResolveField("TECHNICAL_ROWS_RAW"), "TECHNICAL_TABLE", budget)
        For i = 1 To rows.Count
            Set row = rows.Row(i)
            CheckCell row.ResolveField("SOURCE_RAW", budget), row.CellRef("SOURCE_RAW"), "TECH_SOURCE_IP", "SOURCE", False, selection, checkedUpper, issues, manual, hasInput, budget
            CheckCell row.ResolveField("DESTINATION_RAW", budget), row.CellRef("DESTINATION_RAW"), "TECH_DEST_IP", "DESTINATION", False, selection, checkedUpper, issues, manual, hasInput, budget
        Next i
        Set rows = AT_InitiatorRows(context.ResolveField("ADDRESS_ROWS_RAW"), "ADDRESS_TABLE", budget)
        For i = 1 To rows.Count
            Set row = rows.Row(i)
            CheckCell row.ResolveField("ADDRESS_RAW", budget), row.CellRef("ADDRESS_RAW"), "ADDRESS_IP", "ANY", True, selection, checkedUpper, issues, manual, hasInput, budget
        Next i
        If selection.AccountForbidden Then
            Set field = context.ResolveField("DB_ACCOUNT_RAW", budget)
            If Not RawText(field, text, budget) Then
                manual = True
                AT_InitiatorIssue issues, "NET.INITIATOR_IP_ACCESS", "INITIATOR_ACCOUNT_UNCHECKED", "ERROR", "DB_ACCOUNT", context.CellRef("DB_ACCOUNT_RAW"), "Не удалось проверить учётную запись: в ячейке ошибка Excel."
            ElseIf Not AT_ExcelBlank(text) Then
                AT_InitiatorIssue issues, "NET.INITIATOR_IP_ACCESS", "INITIATOR_DB_ACCOUNT_FORBIDDEN", "ERROR", "DB_ACCOUNT", context.CellRef("DB_ACCOUNT_RAW"), "Для этого инициатора поле «Учетная запись» должно оставаться пустым."
            End If
        End If
        If hasInput And Not manual Then status = "PASS"
        For Each issue In issues
            If issue.CheckId = "NET.INITIATOR_IP_ACCESS" And issue.Severity = "ERROR" Then status = "FAIL"
        Next issue
    End If
    Set outcome = New ATCheckOutcome
    text = "": If status = "SKIPPED" Then text = "Проверка IP-прав выполнена не полностью: проверьте почту, адреса и применимость области."
    outcome.Initialize "NET.INITIATOR_IP_ACCESS", status, text: checks.Add outcome
    report.Initialize atEvalOK, "NETWORK_INITIATOR", issues, checks: Set Run = report
    Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "INITIATOR_INPUT_INVALID")
    Set issues = New Collection: Set checks = New Collection: Set outcome = New ATCheckOutcome: Set report = New ATCheckReport
    AT_InitiatorIssue issues, "NET.INITIATOR_POLICY", code, "ERROR", "", "", "Проверки инициатора остановлены. Проверьте книгу правил, заявку и объём данных."
    outcome.Initialize "NET.INITIATOR_POLICY", "BLOCKED", "Проверки инициатора не завершены.": checks.Add outcome
    report.Initialize atEvalError, "NETWORK_INITIATOR", issues, checks: Set Run = report
End Function

Private Function RawText(ByVal field As ATEvalResult, ByRef text As String, ByVal budget As ATEvalBudget) As Boolean
    Dim value As ATRuntimeValue
    text = "": budget.Step
    If field.State <> atEvalOK Then
        If field.State <> atEvalError Then GoTo Invalid
        If field.Diagnostics.Count <> 1 Then GoTo Invalid
        If field.Diagnostics(1).Code <> "EXCEL_CELL_ERROR" Then GoTo Invalid
        Exit Function
    End If
    Set value = field.Value
    If value.DataType <> "TEXT" Then GoTo Invalid
    If value.ValueState = atValuePresent Then
        text = CStr(value.Scalar): budget.CheckText Len(text): budget.Step Len(text)
    ElseIf value.ValueState <> atValueEmpty Then
        GoTo Invalid
    End If
    RawText = True
    Exit Function
Invalid:
    Err.Raise vbObjectError + 2970, , "INITIATOR_INPUT_INVALID"
End Function

Private Sub CheckCell(ByVal field As ATEvalResult, ByVal ref As String, ByVal fieldId As String, ByVal scope As String, ByVal lower As Boolean, ByVal selection As ATInitiatorSelection, ByVal checkedUpper As ATIntervalSet, ByVal issues As Collection, ByRef manual As Boolean, ByRef hasInput As Boolean, ByVal budget As ATEvalBudget)
    Dim text As String, parsed As ATEvalResult, answer As ATEvalResult, value As ATRuntimeValue, unchecked As Boolean
    If Not RawText(field, text, budget) Then GoTo ManualCell
    If AT_ExcelBlank(text) Then Exit Sub
    hasInput = True
    Set parsed = AT_ParseDomain("IP_EXPRESSION", Replace(Replace(text, "х", "x"), "Х", "x"), Nothing, budget)
    If parsed.State <> atEvalOK Then Err.Raise vbObjectError + 2970, , parsed.Diagnostics(1).Code
    Set value = parsed.Value
    If value.ValueState <> atValuePresent Then
        If Not selection.HasRegionPolicy(scope) And selection.NoPolicyAction = "ALLOW" Then
            If Not lower Or Not selection.HasDirectionalRules Then Exit Sub
        End If
        GoTo ManualCell
    End If
    If lower Then
        If value.Domain.IsSubsetOf(checkedUpper, budget) Then Exit Sub
    End If
    Set answer = selection.Evaluate(value, scope, budget)
    If answer.State = atEvalOK Then
        If Not CBool(answer.Value.Scalar) Then
            AT_InitiatorIssue issues, "NET.INITIATOR_IP_ACCESS", "INITIATOR_IP_DENIED", "ERROR", fieldId, ref, "IP-политика инициатора не разрешает значение целиком: " & text
        ElseIf lower And selection.HasDirectionalRules Then
            GoTo ManualCell
        ElseIf Not lower Then
            AT_InitiatorUnion checkedUpper, value.Domain, budget
        End If
    ElseIf answer.State = atEvalSkipped Or answer.State = atEvalManualReview Then
        GoTo ManualCell
    Else
        Err.Raise vbObjectError + 2970, , answer.Diagnostics(1).Code
    End If
    Exit Sub
ManualCell:
    manual = True
    AT_InitiatorIssue issues, "NET.INITIATOR_IP_ACCESS", "INITIATOR_IP_UNCHECKED", "WARNING", fieldId, ref, "IP-права требуют ручной проверки: не удалось подтвердить адрес или определить сторону доступа."
End Sub
