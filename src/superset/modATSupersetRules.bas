Attribute VB_Name = "modATSupersetRules"
Option Explicit
Public Sub AT_SupersetRuleIssue(ByVal issues As Collection, ByVal code As String, ByVal severity As String, ByVal ref As String, ByVal message As String)
    Dim issue As New ATCheckIssue, field As String
    If issues.Count >= 990 Then Err.Raise vbObjectError + 3140, , "CHECK_REPORT_LIMIT"
    If Len(message) > 32767 Then Err.Raise vbObjectError + 3140, , "RULEBOOK_MESSAGE_LIMIT"
    field = "HOSTNAME"
    If code = "RULEBOOK_CONFLICT" Then field = "RULE_CONFLICT"
    issue.Initialize "NET.HOSTNAME_RULES", code, severity, field, ref, message: issues.Add issue
End Sub
Public Function AT_SupersetIntersection(ByVal entry As ATNetworkRegistryEntry, ByVal row As ATRuntimeRow, ByVal budget As ATEvalBudget, ByRef manual As Boolean) As String
    Dim region As New ATIntervalSet, field As Variant, text As String, valid As Boolean, parsed As ATEvalResult
    Dim source As ATIntervalSet, query As ATIntervalSet, a As Long, b As Long, low As Double, high As Double, piece As String
    region.Initialize 4294967295#: Set query = entry.Domain
    For Each field In Array("SOURCE_TELEGRAPH", "SOURCE_CXD", "SOURCE_AWX")
        text = AT_RegistryText(row.ResolveField(CStr(field), budget), valid, budget)
        If Not valid Then
            manual = True
        ElseIf Not AT_ExcelBlank(text) Then
            Set parsed = AT_ParseDomain("IP_EXPRESSION", Replace(Replace(text, "х", "x"), "Х", "x"), Nothing, budget)
            If parsed.State <> atEvalOK Then Err.Raise vbObjectError + 3140, , parsed.Diagnostics(1).Code
            If parsed.Value.ValueState <> atValuePresent Then
                manual = True
            Else
                Set source = parsed.Value.Domain: a = 1: b = 1
                Do While a <= source.Count And b <= query.Count
                    budget.Step
                    low = source.Item(a).Low: If query.Item(b).Low > low Then low = query.Item(b).Low
                    high = source.Item(a).High: If query.Item(b).High < high Then high = query.Item(b).High
                    If low <= high Then region.Add low, high, budget
                    If source.Item(a).High < query.Item(b).High Then a = a + 1 Else b = b + 1
                Loop
            End If
        End If
    Next field
    If region.Count = 0 Then Err.Raise vbObjectError + 3140, , "SUPERSET_RULES_BINDING"
    For a = 1 To region.Count
        piece = AT_IPText(region.Item(a).Low)
        If region.Item(a).Low <> region.Item(a).High Then piece = piece & "-" & AT_IPText(region.Item(a).High)
        If Len(AT_SupersetIntersection) > 0 Then AT_SupersetIntersection = AT_SupersetIntersection & ","
        budget.CheckText Len(AT_SupersetIntersection) + Len(piece): budget.Step Len(piece)
        AT_SupersetIntersection = AT_SupersetIntersection & piece
    Next a
End Function
Public Function AT_SupersetRuleBlock(ByVal row As ATRuntimeRow, ByVal budget As ATEvalBudget) As String
    Dim field As Variant, value As String, text As String
    For Each field In Array(Array("RULE_NAME", "Правило"), Array("SERVICE_RESULT", "Сервис"), Array("DEST_GROUP", "Группа назначения"), Array("APPROVAL", "Согласование"), Array("COMMENT", "Комментарий"), Array("INSTRUCTION", "Инструкция"))
        value = AT_PolicyCell(row, CStr(field(0)), budget)
        If Not AT_ExcelBlank(value) Then
            If Len(text) > 0 Then text = text & vbLf
            text = text & CStr(field(1)) & ": " & value
        End If
    Next field
    text = text & vbLf & "Источник правила: " & row.CellRef("RULE_NAME")
    If Len(text) > 32767 Then Err.Raise vbObjectError + 3140, , "RULEBOOK_MESSAGE_LIMIT"
    budget.Step Len(text): AT_SupersetRuleBlock = text
End Function
Public Sub AT_SupersetAppend(ByRef target As String, ByVal block As String)
    If Len(target) + Len(block) + 2 > 32767 Then Err.Raise vbObjectError + 3140, , "RULEBOOK_MESSAGE_LIMIT"
    If Len(target) > 0 Then target = target & vbLf & vbLf
    target = target & block
End Sub
