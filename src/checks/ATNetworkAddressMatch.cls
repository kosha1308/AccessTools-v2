VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkAddressMatch"
Option Explicit
Public Function Run(ByVal source As ATNetworkReadResult) As ATCheckReport
    Dim context As ATEvaluationContext, upper As New Collection, lower As New Collection
    Dim upperKeys As New Collection, lowerKeys As New Collection, issues As New Collection, checks As New Collection
    Dim budget As New ATEvalBudget, limits As New ATEvalLimits, report As New ATCheckReport, outcome As New ATCheckOutcome
    Dim manual As Boolean, cellError As Boolean, reason As String, status As String, code As String
    On Error GoTo Failed
    If source Is Nothing Then BadInput
    If source.State <> atEvalOK Then
        If source.Diagnostics.Count = 0 Then BadInput
        Set Run = Blocked(source.Diagnostics(1).Code): Exit Function
    End If
    Set context = source.Context: context.Validate: budget.Initialize limits
    CollectRows context, "TECHNICAL_ROWS_RAW", "TECHNICAL_TABLE_RAW", Array(Array("SOURCE_RAW", "TECH_SOURCE_IP"), Array("DESTINATION_RAW", "TECH_DEST_IP")), upper, upperKeys, manual, cellError, budget
    CollectRows context, "ADDRESS_ROWS_RAW", "ADDRESS_TABLE_RAW", Array(Array("ADDRESS_RAW", "ADDRESS_IP")), lower, lowerKeys, manual, cellError, budget
    status = "PASS"
    If cellError Then
        status = "SKIPPED": reason = "Сопоставление пропущено: исправьте ошибки Excel в адресных ячейках."
    ElseIf manual Then
        status = "SKIPPED": reason = "Есть текстовые группы или неподдержанные объявления адресов. Требуется ручное сопоставление."
    ElseIf upperKeys.Count = 0 Then
        status = "SKIPPED": reason = "В технической таблице нет объявлений адресов для сопоставления."
    Else
        CompareSide lower, lowerKeys, upper, "ADDRESS_EXTRA", "Адрес указан в нижней таблице, но отсутствует сверху в том же виде.", issues, budget
        CompareSide upper, upperKeys, lower, "ADDRESS_MISSING", "Повторите это объявление адреса в нижней таблице в том же виде.", issues, budget
        If issues.Count > 0 Then status = "FAIL"
    End If
    outcome.Initialize "NET.ADDRESS_MATCH", status, reason: checks.Add outcome
    report.Initialize atEvalOK, "NETWORK_ADDRESS_MATCH", issues, checks: Set Run = report
    Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "NETWORK_MATCH_INPUT_INVALID")
    If Left$(code, 14) <> "RUNTIME_LIMIT_" And code <> "CHECK_REPORT_LIMIT" Then code = "NETWORK_MATCH_INPUT_INVALID"
    Set Run = Blocked(code)
End Function
Private Sub CollectRows(ByVal context As ATEvaluationContext, ByVal id As String, ByVal tableId As String, ByVal fields As Variant, ByVal index As Collection, ByVal keys As Collection, ByRef manual As Boolean, ByRef cellError As Boolean, ByVal budget As ATEvalBudget)
    Dim result As ATEvalResult, rows As ATRuntimeValue, row As ATRuntimeRow, field As Variant, value As ATEvalResult
    Dim tokens As Collection, token As Variant, prefix As String, key As String, i As Long, ref As String
    Set result = context.ResolveField(id)
    If result.State <> atEvalOK Then BadInput
    Set rows = result.Value
    If rows.DataType <> "COLLECTION<ROW<" & tableId & ">>" Or rows.ValueState <> atValuePresent Then BadInput
    budget.CheckList rows.Count
    For i = 1 To rows.Count
        Set row = rows.Row(i): budget.Step
        For Each field In fields
            Set value = row.ResolveField(CStr(field(0))): ref = row.CellRef(CStr(field(0)))
            If Len(ref) = 0 Then BadInput
            If value.State <> atEvalOK Then
                If value.State <> atEvalError Or value.Diagnostics.Count <> 1 Then BadInput
                If value.Diagnostics(1).Code <> "EXCEL_CELL_ERROR" Then BadInput
                cellError = True
            Else
                If value.Value.DataType <> "TEXT" Then BadInput
                If value.Value.ValueState = atValuePresent Then
                    Set tokens = AT_DomainTokens(CStr(value.Value.Scalar), False, budget): prefix = ""
                    For Each token In tokens
                        budget.Step
                        key = DeclarationKey(CStr(token), prefix, manual, budget)
                        If Len(key) > 0 Then AddDeclarationRef index, keys, key, ref, CStr(field(1)), budget
                    Next token
                ElseIf value.Value.ValueState <> atValueEmpty Then
                    BadInput
                End If
            End If
        Next field
    Next i
End Sub
Private Function DeclarationKey(ByVal token As String, ByRef prefix As String, ByRef manual As Boolean, ByVal budget As ATEvalBudget) As String
    Dim intervals As New ATIntervalSet, declarations As New Collection, candidatePrefix As String, code As String
    On Error GoTo Invalid
    If NeedsManual(token) Then manual = True: prefix = "": Exit Function
    intervals.Initialize 4294967295#: candidatePrefix = prefix
    AT_ParseIPToken token, intervals, declarations, budget, candidatePrefix
    intervals.Freeze: prefix = candidatePrefix
    If InStr(token, "/") > 0 Then
        DeclarationKey = CStr(declarations(1))
    ElseIf InStr(token, "-") > 0 Then
        DeclarationKey = "RANGE:" & AT_IPText(intervals.Item(1).Low) & "-" & AT_IPText(intervals.Item(1).High)
    Else
        DeclarationKey = "IP:" & AT_IPText(intervals.Item(1).Low)
    End If
    Exit Function
Invalid:
    code = AT_StableErrorCode(Err.Description, "NETWORK_MATCH_INPUT_INVALID")
    If Left$(code, 7) <> "DOMAIN_" Then Err.Raise vbObjectError + 2970, , code
    prefix = "": DeclarationKey = "INVALID:" & AT_AsciiCase(token, True)
End Function
Private Function NeedsManual(ByVal token As String) As Boolean
    Dim first As String, i As Long, ch As String, numeric As Boolean
    If InStr(token, ":") > 0 Or IsWildcardMask(token) Then NeedsManual = True: Exit Function
    first = Left$(token, 1)
    If first >= "0" And first <= "9" Then
        If InStr(token, ".") > 0 Or InStr(token, "/") > 0 Then Exit Function
    End If
    numeric = True
    For i = 1 To Len(token)
        ch = Mid$(token, i, 1)
        If InStr(1, "0123456789./-", ch, vbBinaryCompare) = 0 Then numeric = False
    Next i
    NeedsManual = Not numeric
End Function
Private Function IsWildcardMask(ByVal token As String) As Boolean
    Dim parts As Variant, part As Variant, found As Boolean
    parts = Split(AT_AsciiCase(token, False), ".")
    If UBound(parts) <> 3 Then Exit Function
    For Each part In parts
        If CStr(part) = "x" Then
            found = True
        Else
            If Not AT_Digits(CStr(part), 3) Then Exit Function
            If Val(CStr(part)) > 255# Then Exit Function
        End If
    Next part
    IsWildcardMask = found
End Function
Private Sub AddDeclarationRef(ByVal index As Collection, ByVal keys As Collection, ByVal key As String, ByVal ref As String, ByVal id As String, ByVal budget As ATEvalBudget)
    Dim refs As Collection
    budget.Step: Set refs = FindRefs(index, key)
    If refs Is Nothing Then
        budget.CheckList keys.Count + 1
        Set refs = New Collection: index.Add refs, key: keys.Add key
    End If
    If Not HasRef(refs, ref) Then refs.Add Array(ref, id), ref
End Sub
Private Function FindRefs(ByVal index As Collection, ByVal key As String) As Collection
    On Error Resume Next
    Set FindRefs = index(key)
    On Error GoTo 0
End Function
Private Function HasRef(ByVal refs As Collection, ByVal ref As String) As Boolean
    Dim item As Variant
    On Error Resume Next
    item = refs(ref): HasRef = (Err.Number = 0)
    On Error GoTo 0
End Function
Private Sub CompareSide(ByVal index As Collection, ByVal keys As Collection, ByVal other As Collection, ByVal code As String, ByVal message As String, ByVal issues As Collection, ByVal budget As ATEvalBudget)
    Dim key As Variant, refs As Collection, matched As Collection, entry As Variant, issue As ATCheckIssue
    For Each key In keys
        budget.Step: Set matched = FindRefs(other, CStr(key))
        If matched Is Nothing Then
            Set refs = FindRefs(index, CStr(key))
            For Each entry In refs
                budget.Step
                If issues.Count >= 990 Then Err.Raise vbObjectError + 2970, , "CHECK_REPORT_LIMIT"
                Set issue = New ATCheckIssue
                issue.Initialize "NET.ADDRESS_MATCH", code, "ERROR", CStr(entry(1)), CStr(entry(0)), message
                issues.Add issue
            Next entry
        End If
    Next key
End Sub
Private Function Blocked(ByVal code As String) As ATCheckReport
    Dim report As New ATCheckReport, issues As New Collection, checks As New Collection, issue As New ATCheckIssue, outcome As New ATCheckOutcome
    issue.Initialize "NET.ADDRESS_MATCH", code, "ERROR", "", "", "Не удалось завершить сопоставление адресов. Проверьте данные или объём заявки.": issues.Add issue
    outcome.Initialize "NET.ADDRESS_MATCH", "BLOCKED", "Сопоставление адресов остановлено.": checks.Add outcome
    report.Initialize atEvalError, "NETWORK_ADDRESS_MATCH", issues, checks: Set Blocked = report
End Function
Private Sub BadInput()
    Err.Raise vbObjectError + 2970, , "NETWORK_MATCH_INPUT_INVALID"
End Sub
