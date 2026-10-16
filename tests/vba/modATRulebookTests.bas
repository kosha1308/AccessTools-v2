Attribute VB_Name = "modATRulebookTests"
Option Explicit
Public Sub AT_TestRulebook()
    Dim c As New ATNetworkRulebookCompiler, loaded As ATNetworkRulebookResult, r As ATEvalResult
    Dim row As Variant, rows As Variant, v As Variant, book As ATNetworkRulebook, b As ATEvalBudget
    rows = Array(Array("Да", "Web", "192.0.2.0/24", "Назначение", "", "", "443,8443", "TCP", "", "Web service", "", "Approval A", "Comment A", "Информация", "Instruction A"))
    Set loaded = c.Compile(AT_TestRulesRaw(rows), AT_TestRuleProtocols())
    AT_Assert loaded.State = atEvalOK, "rules compile real RAW contract"
    If loaded.State <> atEvalOK Then Exit Sub
    Set book = loaded.Policy
    Set r = book.Technical("198.51.100.1", "192.0.2.1", "443", "tcp", AT_TestInitiatorBudget())
    AT_Assert r.State = atEvalOK And r.Value.Count = 1, "technical scoped IP and full ports match"
    AT_Assert AT_PolicyCell(r.Value.Row(1), "SERVICE_RESULT", AT_TestInitiatorBudget()) = "Web service", "separate service output preserved"
    AT_Assert AT_PolicyCell(r.Value.Row(1), "INSTRUCTION", AT_TestInitiatorBudget()) = "Instruction A", "instruction output preserved"
    AT_Assert r.Value.Row(1).CellRef("RULE_NAME") = "'Правила и исключения'!B2", "rule source ref retained"
    For Each v In Array("443,9000", "0", "443,bad")
        Set r = book.Technical("198.51.100.1", "192.0.2.1", CStr(v), "TCP", AT_TestInitiatorBudget())
        If CStr(v) = "443,9000" Then
            AT_Assert r.State = atEvalOK And r.Value.Count = 0, "partial port coverage is not a match"
        Else
            AT_Assert r.State = atEvalManualReview, "invalid requested ports cannot succeed"
        End If
    Next v
    Set r = book.Technical("192.0.2.1", "198.51.100.1", "443", "TCP", AT_TestInitiatorBudget())
    AT_Assert r.State = atEvalOK And r.Value.Count = 0, "destination scope does not match source"
    Set r = book.Technical("198.51.100.1", "192.0.2.1", "443", "UDP", AT_TestInitiatorBudget())
    AT_Assert r.State = atEvalOK And r.Value.Count = 0, "protocol exact match"
    Set r = book.Technical("group A", "192.0.2.1", "443", "TCP", AT_TestInitiatorBudget())
    AT_Assert r.State = atEvalManualReview, "unknown request IP cannot report completed rule evaluation"
    Set r = book.Superset("192.0.2.1", "web.test", "old", AT_TestInitiatorBudget())
    AT_Assert r.State = atEvalOK And r.Value.Count = 0, "Superset does not guess technical ports"
    rows = Array(AT_TestHostRule("broad", "web", "Содержит", "Broad"), AT_TestHostRule("prefix", "web", "Начинается с", "Prefix"), AT_TestHostRule("exact", "web.test.", "Точное", "Exact"), AT_TestHostRule("tie", "WEB.TEST", "exact", "Other"), AT_TestHostRule("independent", "", "", "Independent"))
    Set loaded = c.Compile(AT_TestRulesRaw(rows), AT_TestRuleProtocols()): Set book = loaded.Policy
    Set r = book.Superset("192.0.2.1", "WEB.TEST.", "", AT_TestInitiatorBudget())
    AT_Assert r.State = atEvalOK And r.Value.Count = 3, "best hostname ties plus non-host rules retained"
    AT_Assert r.Value.Row(1).ResolveField("CONFLICT", AT_TestInitiatorBudget()).Value.Scalar, "different best service outputs flag conflict"
    Set r = book.Technical("192.0.2.1", "192.0.2.2", "443", "TCP", AT_TestInitiatorBudget())
    AT_Assert r.State = atEvalManualReview, "technical hostname context unresolved explicitly"
    rows = Array(Array("Да", "service", "192.0.2.0/24", "", "web.test", "", "", "", "Service With Spaces;Other", "New", "", "", "", "", "Do this"))
    Set loaded = c.Compile(AT_TestRulesRaw(rows), AT_TestRuleProtocols()): Set book = loaded.Policy
    Set r = book.Superset("192.0.2.1", "web.test", "SERVICE WITH SPACES", AT_TestInitiatorBudget())
    AT_Assert r.State = atEvalOK And r.Value.Count = 1, "service selector preserves spaces and compares case insensitive"
    Set r = book.Superset("192.0.2.1", "web.test", "Service", AT_TestInitiatorBudget())
    AT_Assert r.State = atEvalOK And r.Value.Count = 0, "service condition exact, not partial"
    Set r = book.Superset("192.0.2.1", "", "Service With Spaces", AT_TestInitiatorBudget())
    AT_Assert r.State = atEvalManualReview, "missing hostname does not silently complete eligible hostname rule"
    Set loaded = c.Compile(AT_TestRulesRaw(Array(Array("", "draft", "bad", "bad", "", "bad", "bad", "bad", "bad", "", "", "", "", "bad", ""))), AT_TestRuleProtocols())
    AT_Assert loaded.State = atEvalOK, "blank ACTIVE is inactive per real header comment"
    Set loaded = c.Compile(AT_TestRulesRaw(Array()), AT_TestRuleProtocols())
    AT_Assert loaded.State = atEvalOK, "empty supplied rulebook is valid"
    Set r = loaded.Policy.Superset("192.0.2.1", "", "", AT_TestInitiatorBudget())
    AT_Assert r.State = atEvalOK And r.Value.Count = 0, "empty rulebook has no invented findings"
    For Each v In Array("bad", "443tail", "0", "65536")
        row = Array("Да", "bad port", "", "", "", "", v, "", "", "result", "", "", "", "", "")
        Set loaded = c.Compile(AT_TestRulesRaw(Array(row)), AT_TestRuleProtocols())
        AT_Assert loaded.State = atEvalError, "malformed active port blocks entire ruleset"
    Next v
    rows = Array(Array("Да", "no conditions", "", "", "", "", "", "", "", "output", "", "", "", "", ""), Array("Да", "bad IP", "192.0.2.1,bad", "", "", "", "", "", "", "output", "", "", "", "", ""), Array("Да", "bad combo", "", "", "", "", "443", "", "svc", "output", "", "", "", "", ""), Array("Да", "orphan scope", "", "Источник", "", "", "", "", "", "output", "", "", "", "", ""), Array("Да", "no output", "192.0.2.1", "", "", "", "", "", "", "", "", "", "", "", ""), Array("Да", "bad protocol", "", "", "", "", "", "BOGUS", "", "output", "", "", "", "", ""), Array("Да", "bad mode", "", "", "web", "regex", "", "", "", "output", "", "", "", "", ""))
    For Each row In rows
        Set loaded = c.Compile(AT_TestRulesRaw(Array(row)), AT_TestRuleProtocols())
        AT_Assert loaded.State = atEvalError And loaded.Diagnostics.Count > 0, "invalid active rules fail closed"
    Next row
    Set loaded = c.Compile(Nothing, AT_TestRuleProtocols())
    AT_Assert loaded.State = atEvalError, "missing rule source explicit error"
    Set loaded = c.Compile(AT_TestRulesRaw(Array(AT_TestHostRule("valid", "web", "exact", "svc"))), AT_TestRuleProtocols())
    Set b = AT_TestInitiatorBudget(): b.Step 200000
    Set r = loaded.Policy.Superset("192.0.2.1", "web", "", b)
    AT_Assert r.State = atEvalError And r.Diagnostics(1).Code = "RUNTIME_LIMIT_STEPS", "evaluation exhausted budget gives no partial matches"
    AT_TestRulebookEdges
End Sub
Private Sub AT_TestRulebookEdges()
    Dim c As New ATNetworkRulebookCompiler, loaded As ATNetworkRulebookResult, r As ATEvalResult, row As Variant, rows As Variant
    Dim raw As ATEvalResult, b As ATEvalBudget, v As Variant, rule As Variant
    Set loaded = c.Compile(AT_TestRulesRaw(Array(AT_TestHostRule("valid", "web", "exact", "svc"))), AT_TestRuleProtocols())
    Set r = loaded.Policy.Superset("192.0.2.1", "web", "", Nothing)
    AT_Assert r.State = atEvalError, "missing evaluation budget is system error, not manual business context"
    Set r = loaded.Policy.Technical("192.0.2.1", "192.0.2.2", "443", "TCP", Nothing)
    AT_Assert r.State = atEvalError, "technical missing budget is system error"
    rows = Array(AT_TestHostRule("short", "we", "prefix", "Short"), AT_TestHostRule("long", "web", "prefix", "Long"), Array("Да", "wrong base exact", "203.0.113.1", "", "web.test", "exact", "", "", "", "Wrong", "", "", "", "", ""))
    Set loaded = c.Compile(AT_TestRulesRaw(rows), AT_TestRuleProtocols())
    Set r = loaded.Policy.Superset("192.0.2.1", "web.test", "", AT_TestInitiatorBudget())
    AT_Assert r.State = atEvalOK And r.Value.Count = 1, "specificity uses eligible base rules and longer prefix"
    AT_Assert AT_PolicyCell(r.Value.Row(1), "SERVICE_RESULT", AT_TestInitiatorBudget()) = "Long", "unrelated exact hostname cannot suppress matching prefix"
    rows = Array(AT_TestHostRule("same A", "web", "exact", "Service"), AT_TestHostRule("same B", "WEB.", "", "service"))
    Set loaded = c.Compile(AT_TestRulesRaw(rows), AT_TestRuleProtocols())
    Set r = loaded.Policy.Superset("192.0.2.1", "web", "", AT_TestInitiatorBudget())
    AT_Assert r.Value.Count = 2 And Not r.Value.Row(1).ResolveField("CONFLICT").Value.Scalar, "equal case-insensitive outputs are not a conflict"
    AT_Assert AT_PolicyCell(r.Value.Row(1), "SEVERITY", AT_TestInitiatorBudget()) = "INFO", "blank severity legacy info default"
    rows = Array(Array("Да", "instruction only", "192.0.2.0/25", "Любой", "", "", "1-65535", "TCP;UDP", "", "", "", "", "", "Ошибка", "Action"))
    Set loaded = c.Compile(AT_TestRulesRaw(rows), AT_TestRuleProtocols())
    AT_Assert loaded.State = atEvalOK, "instruction only is useful rule output"
    For Each v In Array("1", "65535", "1-65535", "443,8443")
        Set r = loaded.Policy.Technical("192.0.2.0/24", "198.51.100.1", CStr(v), "UDP", AT_TestInitiatorBudget())
        AT_Assert r.State = atEvalOK And r.Value.Count = 1, "rule IP trigger overlap and complete port boundaries"
    Next v
    rows = Array(Array("Да", "first", "", "", "", "", "443", "", "", "A", "", "", "", "", ""), Array("Да", "second", "", "", "", "", "8443", "", "", "B", "", "", "", "", ""))
    Set loaded = c.Compile(AT_TestRulesRaw(rows), AT_TestRuleProtocols())
    Set r = loaded.Policy.Technical("192.0.2.1", "198.51.100.1", "443,8443", "TCP", AT_TestInitiatorBudget())
    AT_Assert r.State = atEvalOK And r.Value.Count = 0, "different rule rows not silently unioned for port coverage"
    For Each v In Array(CVErr(xlErrValue), "maybe")
        row = AT_TestHostRule("invalid active", "web", "", "svc"): row(0) = v
        Set loaded = c.Compile(AT_TestRulesRaw(Array(row)), AT_TestRuleProtocols())
        AT_Assert loaded.State = atEvalError, "active flag error cannot be ignored"
    Next v
    For Each v In Array(2, 4, 6, 8, 9, 14)
        row = AT_TestHostRule("cell error", "web", "", "svc"): row(CLng(v)) = CVErr(xlErrValue)
        Set loaded = c.Compile(AT_TestRulesRaw(Array(row)), AT_TestRuleProtocols())
        AT_Assert loaded.State = atEvalError, "active condition/output cell error blocks complete compile"
    Next v
    row = AT_TestHostRule("large", "web", "", "svc"): row(14) = String$(32768, "x")
    Set loaded = c.Compile(AT_TestRulesRaw(Array(row)), AT_TestRuleProtocols())
    AT_Assert loaded.State = atEvalError, "Excel output limit explicit"
    row = AT_TestHostRule("large", "web", "", "svc"): row(14) = String$(65537, "x")
    Set loaded = c.Compile(AT_TestRulesRaw(Array(row)), AT_TestRuleProtocols())
    AT_Assert loaded.State = atEvalError And loaded.Diagnostics(1).Code = "RUNTIME_LIMIT_TEXT", "compile text budget preserved"
    row = AT_TestHostRule("large", "web", "", "svc"): row(14) = String$(30000, "x")
    Set loaded = c.Compile(AT_TestRulesRaw(Array(row, row, row, row, row, row, row)), AT_TestRuleProtocols())
    AT_Assert loaded.State = atEvalError And loaded.Diagnostics(1).Code = "RUNTIME_LIMIT_STEPS", "compile shared budget across rows"
End Sub
Public Function AT_TestHostRule(ByVal name As String, ByVal host As String, ByVal mode As String, ByVal service As String) As Variant
    AT_TestHostRule = Array("Да", name, "192.0.2.0/24", "", host, mode, "", "", "", service, "", "", "", "", "")
End Function
Public Function AT_TestRuleProtocols() As Collection
    Dim r As New Collection
    r.Add "TCP": r.Add "UDP": Set AT_TestRuleProtocols = r
End Function
Public Function AT_TestRulesRaw(ByVal records As Variant) As ATEvalResult
    Dim fields As Variant, record As Variant, rows As New Collection, row As ATRuntimeRow, value As New ATRuntimeValue
    Dim field As ATEvalResult, i As Long, number As Long
    fields = Array("ACTIVE", "RULE_NAME", "IP", "IP_SCOPE", "HOSTNAME", "HOSTNAME_MATCH", "DEST_PORT", "DEST_PROTOCOL", "SERVICE_CONDITION", "SERVICE_RESULT", "DEST_GROUP", "APPROVAL", "COMMENT", "SEVERITY", "INSTRUCTION")
    For Each record In records
        number = number + 1: Set row = New ATRuntimeRow: row.Initialize "RULES_RAW", CStr(number + 1)
        For i = 0 To UBound(fields)
            If IsError(record(i)) Then
                Set field = AT_RuntimeFailure(atEvalError, "EXCEL_CELL_ERROR", Nothing)
            ElseIf Len(CStr(record(i))) = 0 Then
                Set field = AT_RuntimeOK(AT_RuntimeState("TEXT", atValueEmpty))
            Else
                Set field = AT_RuntimeOK(AT_RuntimeScalar("TEXT", CStr(record(i))))
            End If
            row.AddField CStr(fields(i)), field, "'Правила и исключения'!" & Chr$(65 + i) & CStr(number + 1)
        Next i
        rows.Add row
    Next record
    value.InitializeRows "RULES_RAW", rows: Set AT_TestRulesRaw = AT_RuntimeOK(value)
End Function
