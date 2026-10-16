VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkTableChecks"
Option Explicit
Public Function Run(ByVal source As ATNetworkReadResult, ByVal allowedProtocols As Collection) As ATCheckReport
    Dim normalizer As New ATNetworkRequestNormalizer, prepared As ATNetworkReadResult, context As ATEvaluationContext
    Dim issues As New Collection, checks As New Collection, report As New ATCheckReport
    Dim budget As New ATEvalBudget, limits As New ATEvalLimits, catalog As Collection
    Dim raw As ATRuntimeValue, typed As ATRuntimeValue, row As ATRuntimeRow, normalizedRow As ATRuntimeRow
    Dim seen As New Collection, key As String, first As String, i As Long, entry As Variant, value As ATEvalResult
    Dim before As Long, manual As Boolean, code As String, tokens As Collection, token As Variant
    On Error GoTo Failed
    Set prepared = normalizer.Normalize(source, "DMY_TEXT", "Бессрочно")
    If prepared.State <> atEvalOK Then
        Set Run = Blocked(prepared.Diagnostics(1).Code, prepared.Diagnostics(1).Message): Exit Function
    End If
    Set context = prepared.Context: budget.Initialize limits
    Set catalog = ProtocolCatalog(allowedProtocols)
    Set raw = context.ResolveField("TECHNICAL_ROWS_RAW").Value
    Set typed = context.ResolveField("TECHNICAL_ROWS").Value
    RequireRef context.CellRef("TECHNICAL_ROWS_RAW")
    If raw.Count = 0 Then AddIssue issues, "NET.TECHNICAL_TABLE", "TABLE_EMPTY", "ERROR", "TECHNICAL_ROWS", context.CellRef("TECHNICAL_ROWS_RAW"), "Заполните техническую таблицу: источник, назначение, протокол и порт."
    For i = 1 To raw.Count
        Set row = raw.Row(i): Set normalizedRow = typed.Row(i)
        For Each entry In Array(Array("SOURCE_RAW", "TECH_SOURCE_IP"), Array("DESTINATION_RAW", "TECH_DEST_IP"))
            CheckIP row, normalizedRow, CStr(entry(0)), CStr(entry(1)), "NET.TECHNICAL_TABLE", issues, manual, budget
        Next entry
        Set value = row.ResolveField("PROTOCOL_RAW")
        CheckRequired value, "NET.TECHNICAL_TABLE", "TECH_PROTOCOL", row.CellRef("PROTOCOL_RAW"), issues
        Set value = row.ResolveField("PORT_RAW")
        If CheckRequired(value, "NET.TECHNICAL_TABLE", "TECH_DEST_PORT", row.CellRef("PORT_RAW"), issues) Then
            Set value = normalizedRow.ResolveField("TECH_DEST_PORT")
            If value.Value.ValueState = atValueInvalid Then
                AddIssue issues, "NET.TECHNICAL_TABLE", "PORT_INVALID", "ERROR", "TECH_DEST_PORT", row.CellRef("PORT_RAW"), "Укажите порты от 1 до 65535, список или диапазон."
            ElseIf value.Value.Domain.Item(1).Low < 1# Then
                AddIssue issues, "NET.TECHNICAL_TABLE", "PORT_INVALID", "ERROR", "TECH_DEST_PORT", row.CellRef("PORT_RAW"), "Порт 0 не допускается в заявке. Укажите порт от 1 до 65535."
            End If
        End If
        CheckOptional row.ResolveField("REASON_RAW"), "NET.TECHNICAL_TABLE", "TECH_REASON", row.CellRef("REASON_RAW"), "Добавьте обоснование связанности.", issues
        key = ""
        For Each entry In Array("SOURCE_RAW", "DESTINATION_RAW", "PROTOCOL_RAW", "PORT_RAW")
            Set value = row.ResolveField(CStr(entry)): budget.Step
            key = key & DuplicatePart(value)
        Next entry
        first = SeenRow(seen, key)
        If Len(first) > 0 Then
            AddIssue issues, "NET.TECHNICAL_TABLE", "ROW_DUPLICATE", "WARNING", "TECH_SOURCE_IP", row.CellRef("SOURCE_RAW"), "Связанность повторяет строку " & first & ". Проверьте необходимость дубликата."
        Else
            seen.Add row.RowId, key
        End If
    Next i
    AddOutcome checks, issues, "NET.TECHNICAL_TABLE", 0, manual
    before = issues.Count
    If catalog Is Nothing Then
        AddIssue issues, "NET.PROTOCOLS", "PROTOCOL_CATALOG_MISSING", "WARNING", "TECH_PROTOCOL", context.CellRef("TECHNICAL_ROWS_RAW"), "Нет справочника разрешённых протоколов. Проверка допустимости протоколов пропущена."
        AddOutcome checks, issues, "NET.PROTOCOLS", before, True
    Else
        For i = 1 To raw.Count
            Set row = raw.Row(i): Set value = row.ResolveField("PROTOCOL_RAW")
            If value.State = atEvalOK Then
                If value.Value.ValueState = atValuePresent Then
                    Set tokens = ProtocolTokens(CStr(value.Value.Scalar), budget)
                    If tokens.Count = 0 Then AddIssue issues, "NET.PROTOCOLS", "PROTOCOL_INVALID", "ERROR", "TECH_PROTOCOL", row.CellRef("PROTOCOL_RAW"), "Укажите протокол, а не только разделители."
                    For Each token In tokens
                        If Not HasProtocol(catalog, AT_AsciiCase(CStr(token), True)) Then
                            AddIssue issues, "NET.PROTOCOLS", "PROTOCOL_INVALID", "ERROR", "TECH_PROTOCOL", row.CellRef("PROTOCOL_RAW"), "Указан протокол, отсутствующий в справочнике разрешённых протоколов."
                            Exit For
                        End If
                    Next token
                End If
            End If
        Next i
        AddOutcome checks, issues, "NET.PROTOCOLS", before, False
    End If
    before = issues.Count: manual = False
    Set raw = context.ResolveField("ADDRESS_ROWS_RAW").Value: Set typed = context.ResolveField("ADDRESS_ROWS").Value
    RequireRef context.CellRef("ADDRESS_ROWS_RAW")
    If raw.Count = 0 Then AddIssue issues, "NET.ADDRESS_TABLE", "TABLE_EMPTY", "ERROR", "ADDRESS_ROWS", context.CellRef("ADDRESS_ROWS_RAW"), "Заполните нижнюю таблицу адресов и имён хостов."
    For i = 1 To raw.Count
        Set row = raw.Row(i): Set normalizedRow = typed.Row(i)
        CheckIP row, normalizedRow, "ADDRESS_RAW", "ADDRESS_IP", "NET.ADDRESS_TABLE", issues, manual, budget
        CheckRequired row.ResolveField("HOSTNAME_RAW"), "NET.ADDRESS_TABLE", "HOSTNAME", row.CellRef("HOSTNAME_RAW"), issues
        CheckOptional row.ResolveField("SYSTEM_RAW"), "NET.ADDRESS_TABLE", "SYSTEM", row.CellRef("SYSTEM_RAW"), "Укажите название системы.", issues
        CheckOptional row.ResolveField("ROLE_RAW"), "NET.ADDRESS_TABLE", "ROLE", row.CellRef("ROLE_RAW"), "Добавьте роль или назначение адреса.", issues
    Next i
    AddOutcome checks, issues, "NET.ADDRESS_TABLE", before, manual
    report.Initialize atEvalOK, "NETWORK_TABLE_INPUT", issues, checks: Set Run = report
    Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "NETWORK_TABLE_INPUT_INVALID")
    If Left$(code, 7) <> "CONFIG_" And Left$(code, 14) <> "RUNTIME_LIMIT_" And code <> "CHECK_REPORT_LIMIT" Then code = "NETWORK_TABLE_INPUT_INVALID"
    Set Run = Blocked(code, "Не удалось завершить проверки таблиц. Проверьте данные и справочник протоколов.")
End Function
Private Function ProtocolTokens(ByVal text As String, ByVal budget As ATEvalBudget) As Collection
    Dim result As New Collection, i As Long, start As Long, ch As String, separator As Boolean
    budget.CheckText Len(text)
    For i = 1 To Len(text) + 1
        budget.Step: ch = Mid$(text, i, 1)
        separator = (Len(ch) = 0 Or InStr(1, ",;| " & vbTab & vbCr & vbLf, ch, vbBinaryCompare) > 0)
        If separator Then
            If start > 0 Then
                budget.CheckList result.Count + 1: result.Add Mid$(text, start, i - start): start = 0
            End If
        ElseIf start = 0 Then
            start = i
        End If
    Next i
    Set ProtocolTokens = result
End Function
Private Sub CheckIP(ByVal raw As ATRuntimeRow, ByVal typed As ATRuntimeRow, ByVal rawId As String, ByVal id As String, ByVal checkId As String, ByVal issues As Collection, ByRef manual As Boolean, ByVal budget As ATEvalBudget)
    Dim value As ATEvalResult, text As String, i As Long, ch As String, unusual As Boolean
    Set value = raw.ResolveField(rawId)
    If Not CheckRequired(value, checkId, id, raw.CellRef(rawId), issues) Then Exit Sub
    text = CStr(value.Value.Scalar): Set value = typed.ResolveField(id)
    If value.Value.ValueState <> atValueInvalid Then Exit Sub
    For i = 1 To Len(text)
        budget.Step: ch = Mid$(text, i, 1)
        If InStr(1, "0123456789.xX/,-;| " & vbTab & vbCr & vbLf, ch, vbBinaryCompare) = 0 Then unusual = True
    Next i
    If unusual Then
        AddIssue issues, checkId, "ADDRESS_MANUAL_REQUIRED", "WARNING", id, raw.CellRef(rawId), "Возможно, указана текстовая группа адресов. Состав требует ручной проверки; автоматическое подтверждение не выполнено."
        manual = True
    Else
        AddIssue issues, checkId, "IP_INVALID", "ERROR", id, raw.CellRef(rawId), "Некорректная IP-запись. Проверьте адрес, сеть или диапазон целиком."
    End If
End Sub
Private Function CheckRequired(ByVal value As ATEvalResult, ByVal checkId As String, ByVal id As String, ByVal ref As String, ByVal issues As Collection) As Boolean
    RequireRef ref
    If value.State <> atEvalOK Then
        If value.State <> atEvalError Or value.Diagnostics.Count <> 1 Then BadInput
        If value.Diagnostics(1).Code <> "EXCEL_CELL_ERROR" Then BadInput
        AddIssue issues, checkId, "CELL_ERROR", "ERROR", id, ref, "В ячейке ошибка Excel. Исправьте значение и повторите проверку."
    ElseIf value.Value.ValueState = atValueEmpty Then
        AddIssue issues, checkId, "FIELD_REQUIRED", "ERROR", id, ref, "Заполните обязательное поле."
    Else
        If value.Value.ValueState <> atValuePresent Then BadInput
        CheckRequired = True
    End If
End Function
Private Sub CheckOptional(ByVal value As ATEvalResult, ByVal checkId As String, ByVal id As String, ByVal ref As String, ByVal message As String, ByVal issues As Collection)
    RequireRef ref
    If value.State <> atEvalOK Then
        CheckRequired value, checkId, id, ref, issues
    ElseIf value.Value.ValueState = atValueEmpty Then
        AddIssue issues, checkId, "FIELD_RECOMMENDED", "WARNING", id, ref, message
    ElseIf value.Value.ValueState <> atValuePresent Then
        BadInput
    End If
End Sub
Private Function ProtocolCatalog(ByVal source As Collection) As Collection
    Dim result As New Collection, entry As Variant, value As String, i As Long, ch As String
    If source Is Nothing Then Exit Function
    If source.Count = 0 Then Exit Function
    If source.Count > 1000 Then Err.Raise vbObjectError + 2950, , "CONFIG_PROTOCOL_CATALOG"
    For Each entry In source
        If VarType(entry) <> vbString Then Err.Raise vbObjectError + 2950, , "CONFIG_PROTOCOL_CATALOG"
        value = AT_AsciiCase(Trim$(CStr(entry)), True)
        If Len(value) = 0 Or Len(value) > 128 Then Err.Raise vbObjectError + 2950, , "CONFIG_PROTOCOL_CATALOG"
        For i = 1 To Len(value)
            ch = Mid$(value, i, 1)
            If Not ((ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or InStr(1, "_-./+", ch, vbBinaryCompare) > 0) Then Err.Raise vbObjectError + 2950, , "CONFIG_PROTOCOL_CATALOG"
        Next i
        If Not HasProtocol(result, value) Then result.Add True, "P" & value
    Next entry
    Set ProtocolCatalog = result
End Function
Private Function HasProtocol(ByVal catalog As Collection, ByVal value As String) As Boolean
    Dim found As Boolean
    On Error Resume Next
    found = catalog("P" & value): HasProtocol = (Err.Number = 0)
    On Error GoTo 0
End Function
Private Function DuplicatePart(ByVal value As ATEvalResult) As String
    Dim text As String
    If value.State <> atEvalOK Then
        DuplicatePart = "E;": Exit Function
    End If
    If value.Value.ValueState = atValuePresent Then text = AT_ExcelHeader(CStr(value.Value.Scalar))
    DuplicatePart = "T" & CStr(Len(text)) & ":" & text
End Function
Private Function SeenRow(ByVal seen As Collection, ByVal key As String) As String
    On Error Resume Next
    SeenRow = seen(key)
    On Error GoTo 0
End Function
Private Sub AddIssue(ByVal issues As Collection, ByVal checkId As String, ByVal code As String, ByVal severity As String, ByVal id As String, ByVal ref As String, ByVal message As String)
    Dim issue As New ATCheckIssue
    If issues.Count >= 990 Then Err.Raise vbObjectError + 2950, , "CHECK_REPORT_LIMIT"
    issue.Initialize checkId, code, severity, id, ref, message: issues.Add issue
End Sub
Private Sub AddOutcome(ByVal checks As Collection, ByVal issues As Collection, ByVal id As String, ByVal before As Long, ByVal manual As Boolean)
    Dim outcome As New ATCheckOutcome, status As String, reason As String, i As Long
    status = "PASS"
    If manual Then status = "SKIPPED": reason = "Автоматическая проверка выполнена не полностью; требуется справочник или ручная проверка."
    For i = before + 1 To issues.Count
        If issues(i).Severity = "ERROR" Then status = "FAIL": reason = ""
    Next i
    outcome.Initialize id, status, reason: checks.Add outcome
End Sub
Private Function Blocked(ByVal code As String, ByVal message As String) As ATCheckReport
    Dim report As New ATCheckReport, issues As New Collection, checks As New Collection
    Dim issue As New ATCheckIssue, outcome As New ATCheckOutcome
    issue.Initialize "NET.TABLE_INPUT", code, "ERROR", "", "", "Проверки таблиц остановлены. " & message: issues.Add issue
    outcome.Initialize "NET.TABLE_INPUT", "BLOCKED", "Проверки таблиц не завершены.": checks.Add outcome
    report.Initialize atEvalError, "NETWORK_TABLE_INPUT", issues, checks: Set Blocked = report
End Function
Private Sub RequireRef(ByVal ref As String)
    If Len(ref) = 0 Then BadInput
End Sub
Private Sub BadInput()
    Err.Raise vbObjectError + 2950, , "NETWORK_TABLE_INPUT_INVALID"
End Sub
