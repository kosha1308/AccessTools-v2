VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkRuleChecks"
Option Explicit
Public Function Run(ByVal source As ATNetworkReadResult, ByVal rules As ATNetworkRulebookResult) As ATCheckReport
    Dim report As New ATCheckReport, issues As New Collection, checks As New Collection, outcome As New ATCheckOutcome
    Dim budget As New ATEvalBudget, limits As New ATEvalLimits, context As ATEvaluationContext, rows As ATRuntimeValue
    Dim row As ATRuntimeRow, fields As Variant, texts(0 To 3) As String, valid As Boolean, rowManual As Boolean, manual As Boolean
    Dim tokens As Collection, token As Variant, seen As Object, answer As ATEvalResult, i As Long, j As Long
    Dim issue As ATCheckIssue, status As String, reason As String, code As String, ref As String
    On Error GoTo Failed
    If rules Is Nothing Then
        AddIssue issues, "RULEBOOK_MISSING", "WARNING", "TECHNICAL_ROWS", "", "Нет книги общих правил. Проверка правил и исключений пропущена."
        manual = True: GoTo Finish
    End If
    If rules.State <> atEvalOK Then Err.Raise vbObjectError + 2990, , rules.Diagnostics(1).Code
    If source Is Nothing Then Err.Raise vbObjectError + 2990, , "RULEBOOK_INPUT_INVALID"
    If source.State <> atEvalOK Then Err.Raise vbObjectError + 2990, , source.Diagnostics(1).Code
    budget.Initialize limits: Set context = source.Context: context.Validate
    Set rows = AT_InitiatorRows(context.ResolveField("TECHNICAL_ROWS_RAW"), "TECHNICAL_TABLE", budget)
    fields = Array("SOURCE_RAW", "DESTINATION_RAW", "PORT_RAW", "PROTOCOL_RAW")
    If rows.Count = 0 Then manual = True
    For i = 1 To rows.Count
        Set row = rows.Row(i): valid = True: rowManual = False: ref = row.CellRef("SOURCE_RAW")
        If Len(ref) = 0 Then Err.Raise vbObjectError + 2990, , "RULEBOOK_INPUT_INVALID"
        For j = 0 To 3
            If Not ReadText(row.ResolveField(CStr(fields(j)), budget), texts(j), budget) Then valid = False
        Next j
        If valid Then
            Set tokens = AT_InitiatorTokens(texts(3), budget): Set seen = CreateObject("Scripting.Dictionary")
            If tokens.Count = 0 Then rowManual = True
            For Each token In tokens
                code = AT_AsciiCase(CStr(token), True)
                If Not seen.Exists(code) Then
                    seen.Add code, True
                    Set answer = rules.Policy.Technical(texts(0), texts(1), texts(2), CStr(token), budget)
                    Select Case answer.State
                        Case atEvalOK: AddMatches answer.Value, texts(0), texts(1), texts(2), CStr(token), ref, issues, budget
                        Case atEvalManualReview, atEvalSkipped: rowManual = True
                        Case Else: Err.Raise vbObjectError + 2990, , answer.Diagnostics(1).Code
                    End Select
                End If
            Next token
        Else
            rowManual = True
        End If
        If rowManual Then
            manual = True
            AddIssue issues, "RULEBOOK_CONTEXT_UNCHECKED", "WARNING", "TECHNICAL_ROWS", ref, "Правила для связанности проверены не полностью. Уточните IP, порты, протоколы или hostname; ошибка Excel также требует исправления."
        End If
    Next i
Finish:
    status = "PASS": reason = "Проверены технические условия. Правила, зависящие от Superset, ещё не вычислялись."
    If manual Then status = "SKIPPED": reason = "Правила и исключения проверены не полностью. Требуется справочник или ручная проверка."
    For Each issue In issues
        If issue.Severity = "ERROR" Then status = "FAIL"
    Next issue
    outcome.Initialize "NET.IP_EXCEPTIONS", status, reason: checks.Add outcome
    report.Initialize atEvalOK, "NETWORK_RULES", issues, checks: Set Run = report: Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "RULEBOOK_INPUT_INVALID")
    Set issues = New Collection: Set checks = New Collection: Set outcome = New ATCheckOutcome: Set report = New ATCheckReport
    AddIssue issues, code, "ERROR", "TECHNICAL_ROWS", "", "Проверка общих правил остановлена. Проверьте книгу правил, заявку и объём данных. Остальные проверки выполняются отдельно."
    outcome.Initialize "NET.IP_EXCEPTIONS", "BLOCKED", "Правила и исключения не проверены.": checks.Add outcome
    report.Initialize atEvalError, "NETWORK_RULES", issues, checks: Set Run = report
End Function
Private Function ReadText(ByVal field As ATEvalResult, ByRef text As String, ByVal budget As ATEvalBudget) As Boolean
    Dim value As ATRuntimeValue
    text = "": budget.Step
    If field.State <> atEvalOK Then
        If field.State = atEvalError Then
            If field.Diagnostics.Count = 1 Then
                If field.Diagnostics(1).Code = "EXCEL_CELL_ERROR" Then Exit Function
            End If
        End If
        Err.Raise vbObjectError + 2990, , "RULEBOOK_INPUT_INVALID"
    End If
    Set value = field.Value
    If value.DataType <> "TEXT" Then Err.Raise vbObjectError + 2990, , "RULEBOOK_INPUT_INVALID"
    Select Case value.ValueState
        Case atValuePresent
            text = CStr(value.Scalar): budget.CheckText Len(text): budget.Step Len(text)
        Case atValueEmpty
        Case Else: Err.Raise vbObjectError + 2990, , "RULEBOOK_INPUT_INVALID"
    End Select
    ReadText = True
End Function
Private Sub AddMatches(ByVal matches As ATRuntimeValue, ByVal source As String, ByVal destination As String, ByVal ports As String, ByVal protocol As String, ByVal ref As String, ByVal issues As Collection, ByVal budget As ATEvalBudget)
    Dim i As Long, row As ATRuntimeRow, text As String, prefix As String, field As Variant, value As String, names As String
    Dim severity As String, conflict As Boolean
    prefix = "Связанность: " & source & " -> " & destination & "; порт: " & ports & "; протокол: " & protocol & "."
    budget.CheckText Len(prefix): budget.Step Len(prefix) + 1
    For i = 1 To matches.Count
        Set row = matches.Row(i): text = prefix
        For Each field In Array(Array("RULE_NAME", "Правило"), Array("SERVICE_RESULT", "Сервис"), Array("DEST_GROUP", "Группа назначения"), Array("APPROVAL", "Согласование"), Array("COMMENT", "Комментарий"), Array("INSTRUCTION", "Инструкция"))
            value = AT_PolicyCell(row, CStr(field(0)), budget)
            If Not AT_ExcelBlank(value) Then text = text & vbLf & CStr(field(1)) & ": " & value
        Next field
        text = text & vbLf & "Источник правила: " & row.CellRef("RULE_NAME")
        severity = AT_PolicyCell(row, "SEVERITY", budget)
        AddIssue issues, "RULEBOOK_MATCH", severity, "RULE_RESULT", ref, text
        If row.ResolveField("CONFLICT", budget).Value.Scalar Then conflict = True
        value = AT_PolicyCell(row, "RULE_NAME", budget)
        If AT_ExcelBlank(value) Then value = row.CellRef("RULE_NAME")
        names = names & vbLf & value
        If Len(names) > 32767 Then Err.Raise vbObjectError + 2990, , "RULEBOOK_MESSAGE_LIMIT"
    Next i
    If conflict Then AddIssue issues, "RULEBOOK_CONFLICT", "ERROR", "RULE_CONFLICT", ref, prefix & vbLf & "Найдены конфликтующие сервисы или группы. Проверьте правила:" & names
End Sub
Private Sub AddIssue(ByVal issues As Collection, ByVal code As String, ByVal severity As String, ByVal field As String, ByVal ref As String, ByVal message As String)
    Dim issue As New ATCheckIssue
    If Len(message) > 32767 Then Err.Raise vbObjectError + 2990, , "RULEBOOK_MESSAGE_LIMIT"
    If issues.Count >= 990 Then Err.Raise vbObjectError + 2990, , "CHECK_REPORT_LIMIT"
    issue.Initialize "NET.IP_EXCEPTIONS", code, severity, field, ref, message: issues.Add issue
End Sub
