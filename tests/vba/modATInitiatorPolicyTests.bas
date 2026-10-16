Attribute VB_Name = "modATInitiatorPolicyTests"
Option Explicit

Public Sub AT_TestInitiatorPolicy()
    Dim compiler As New ATInitiatorPolicyCompiler, loaded As ATInitiatorPolicyResult, policy As ATInitiatorPolicy
    Dim selection As ATInitiatorSelection, budget As ATEvalBudget, report As ATCheckReport, result As ATEvalResult
    Dim identities As ATEvalResult, ip As ATEvalResult, row As Variant, value As Variant, action As Variant
    Set identities = AT_TestInitiatorRaw("INITIATORS", Array(Array("Да", "example.test; НЕ:special@example.test", "Согласование A", "Предупреждение", "Комментарий", ""), Array("Да", "special@example.test", "Согласование B", "Ошибка", "", "APPROVAL_ONLY")))
    Set ip = AT_TestInitiatorRaw("INITIATOR_IP", Array(Array("", "@example.test", "", "10.0.0.0/8; НЕ:10.2.x.x", "10.3.0.0/16", "", "")))
    Set loaded = compiler.Compile(identities, ip, AT_NETWORK_NO_IP_POLICY_ACTION)
    AT_Assert loaded.State = atEvalOK, "actor policy compiles explicit published defaults"
    Set policy = loaded.Policy: Set budget = AT_TestInitiatorBudget(): Set selection = policy.ForEmail("USER@EXAMPLE.TEST", budget)
    Set report = selection.IdentityReport("'Лист1'!F5")
    AT_Assert report.Checks.Count = 2 And report.Checks(1).Status = "PASS", "domain admission case-insensitive"
    AT_Assert report.Issues.Count = 1 And report.Issues(1).CheckId = "NET.INITIATOR_APPROVAL_RULES", "approval remains independent issue"
    AT_Assert report.Issues(1).CellRef = "'Лист1'!F5" And InStr(report.Issues(1).Message, "Согласование A") > 0, "approval has actual email ref"
    Set selection = policy.ForEmail("special@example.test", AT_TestInitiatorBudget())
    Set report = selection.IdentityReport("'Лист1'!F5")
    AT_Assert report.Checks(1).Status = "FAIL" And report.Issues.Count = 2, "negative selector excludes admission, approval-only row does not admit"
    Set selection = policy.ForEmail("user@evil-example.test", AT_TestInitiatorBudget())
    AT_Assert selection.IdentityReport("F5").Checks(1).Status = "FAIL", "domain is exact, not suffix match"
    Set selection = policy.ForEmail("bad email", AT_TestInitiatorBudget())
    AT_Assert selection.IdentityReport("F5").SkippedCount = 2, "invalid email does not cause duplicate admission/approval errors"
    AT_TestInitiatorAllows ip, "user@example.test", "10.1.0.1", "SOURCE", True
    AT_TestInitiatorAllows ip, "user@example.test", "10.2.0.1", "SOURCE", False
    AT_TestInitiatorAllows ip, "user@example.test", "10.3.0.1", "DESTINATION", False
    AT_TestInitiatorAllows ip, "user@example.test", "11.0.0.1", "DESTINATION", False
    AT_TestInitiatorAllows ip, "outside@other.test", "11.0.0.1", "ANY", True
    Set ip = AT_TestInitiatorRaw("INITIATOR_IP", Array(Array("Да", "example.test", "Любой", "10.0.0.0/8", "", "", ""), Array("Да", "user@example.test", "Любой", "192.0.2.1", "", "", "Запрещено")))
    AT_TestInitiatorAllows ip, "USER@example.test", "10.1.0.1", "SOURCE", False
    AT_TestInitiatorAllows ip, "USER@example.test", "192.0.2.1", "SOURCE", True
    Set loaded = compiler.Compile(AT_TestInitiatorRaw("INITIATORS", Array()), ip, "ALLOW")
    Set selection = loaded.Policy.ForEmail("user@example.test", AT_TestInitiatorBudget())
    AT_Assert selection.AccountForbidden, "personal matching policy forbids DB account"
    Set selection = loaded.Policy.ForEmail("other@example.test", AT_TestInitiatorBudget())
    AT_Assert Not selection.AccountForbidden, "personal account restriction not applied to another email"
    Set ip = AT_TestInitiatorRaw("INITIATOR_IP", Array(Array("Да", "example.test", "Любой", "10.81.х.х", "", "", "")))
    For Each value In Array("10.81.0.0", "10.81.255.255", "10.81.0.0/16", "10.81.0.1-10.81.0.9")
        AT_TestInitiatorAllows ip, "user@example.test", CStr(value), "SOURCE", True
    Next value
    AT_TestInitiatorAllows ip, "user@example.test", "10.82.0.0", "SOURCE", False
    AT_TestInitiatorAllows ip, "user@example.test", "10.81.255.255-10.82.0.0", "SOURCE", False
    Set ip = AT_TestInitiatorRaw("INITIATOR_IP", Array(Array("Да", "example.test", "Любой", "192.0.2.0/25", "", "", ""), Array("Да", "example.test", "Любой", "192.0.2.128/25", "", "", "")))
    AT_TestInitiatorAllows ip, "user@example.test", "192.0.2.0/24", "ANY", True
    AT_TestInitiatorAllows ip, "user@example.test", "192.0.2.0/23", "ANY", False
    Set ip = AT_TestInitiatorRaw("INITIATOR_IP", Array(Array("Да", "example.test", "Источник", "192.0.2.1", "", "", "")))
    AT_TestInitiatorAllows ip, "user@example.test", "192.0.2.2", "SOURCE", False
    AT_TestInitiatorAllows ip, "user@example.test", "192.0.2.2", "DESTINATION", True
    Set loaded = compiler.Compile(AT_TestInitiatorRaw("INITIATORS", Array()), ip, "ALLOW")
    AT_Assert loaded.Policy.ForEmail("user@example.test", AT_TestInitiatorBudget()).HasDirectionalRules, "directional policy recorded for lower-table binding"
    For Each action In Array("ALLOW", "DENY", "WARNING", "SKIPPED")
        Set loaded = compiler.Compile(AT_TestInitiatorRaw("INITIATORS", Array()), AT_TestInitiatorRaw("INITIATOR_IP", Array()), CStr(action))
        Set budget = AT_TestInitiatorBudget(): Set selection = loaded.Policy.ForEmail("user@example.test", budget)
        Set result = selection.Evaluate(AT_ParseDomain("IP_EXPRESSION", "192.0.2.1", Nothing).Value, "ANY", budget)
        Select Case CStr(action)
            Case "ALLOW": AT_Assert result.State = atEvalOK And result.Value.Scalar, "explicit no-policy ALLOW"
            Case "DENY": AT_Assert result.State = atEvalOK And Not result.Value.Scalar, "explicit no-policy DENY"
            Case "WARNING": AT_Assert result.State = atEvalManualReview, "explicit no-policy warning"
            Case "SKIPPED": AT_Assert result.State = atEvalSkipped, "explicit no-policy skipped"
        End Select
    Next action
    AT_TestInitiatorBad compiler.Compile(identities, ip, ""), "INITIATOR_NO_POLICY_ACTION"
    Set ip = AT_TestInitiatorRaw("INITIATOR_IP", Array(Array("Нет", "bad", "bad", "bad", "bad", "", "bad")))
    Set loaded = compiler.Compile(identities, ip, "ALLOW")
    AT_Assert loaded.State = atEvalOK, "inactive IP draft ignored"
    For Each row In Array(Array("Да", "bad email", "Любой", "10.0.0.0/8", "", "", ""), Array("Да", "NOT:example.test", "Любой", "10.0.0.0/8", "", "", ""))
        AT_TestInitiatorBad compiler.Compile(identities, AT_TestInitiatorRaw("INITIATOR_IP", Array(row)), "ALLOW"), "INITIATOR_SELECTOR_INVALID"
    Next row
    AT_TestInitiatorBad compiler.Compile(identities, AT_TestInitiatorRaw("INITIATOR_IP", Array(Array("Да", "example.test", "Other", "10.0.0.0/8", "", "", ""))), "ALLOW"), "INITIATOR_SCOPE_INVALID"
    AT_TestInitiatorBad compiler.Compile(identities, AT_TestInitiatorRaw("INITIATOR_IP", Array(Array("Да", "example.test", "Любой", "10.0.0.999tail", "", "", ""))), "ALLOW"), "INITIATOR_IP_INVALID"
    AT_TestInitiatorBad compiler.Compile(identities, AT_TestInitiatorRaw("INITIATOR_IP", Array(Array("Да", "example.test", "Любой", "10.0.0.0/8", "", "", "someone"))), "ALLOW"), "INITIATOR_DB_INVALID"
    AT_TestInitiatorBad compiler.Compile(Nothing, Nothing, "ALLOW"), "INITIATOR_INPUT_INVALID"
End Sub

Private Sub AT_TestInitiatorAllows(ByVal ip As ATEvalResult, ByVal email As String, ByVal expression As String, ByVal scope As String, ByVal expected As Boolean)
    Dim compiler As New ATInitiatorPolicyCompiler, loaded As ATInitiatorPolicyResult, selection As ATInitiatorSelection
    Dim budget As ATEvalBudget, parsed As ATEvalResult, result As ATEvalResult
    Set loaded = compiler.Compile(AT_TestInitiatorRaw("INITIATORS", Array()), ip, "ALLOW")
    AT_Assert loaded.State = atEvalOK, "IP corpus compiles " & expression
    If loaded.State <> atEvalOK Then Exit Sub
    Set budget = AT_TestInitiatorBudget(): Set selection = loaded.Policy.ForEmail(email, budget)
    Set parsed = AT_ParseDomain("IP_EXPRESSION", expression, Nothing, budget)
    Set result = selection.Evaluate(parsed.Value, scope, budget)
    AT_Assert result.State = atEvalOK And result.Value.Scalar = expected, "IP full coverage/deny " & expression & "/" & scope
End Sub

Public Function AT_TestInitiatorBudget() As ATEvalBudget
    Dim limits As New ATEvalLimits, budget As New ATEvalBudget
    budget.Initialize limits: Set AT_TestInitiatorBudget = budget
End Function

Public Function AT_TestInitiatorRaw(ByVal kind As String, ByVal records As Variant) As ATEvalResult
    Dim fields As Variant, record As Variant, rows As New Collection, row As ATRuntimeRow, value As New ATRuntimeValue
    Dim field As ATEvalResult, i As Long, number As Long
    If kind = "INITIATORS" Then
        fields = Array("ACTIVE", "EMAIL_DOMAINS", "APPROVAL", "SEVERITY", "COMMENT", "SERVICE_MODE")
    Else
        fields = Array("ACTIVE", "EMAIL_DOMAIN", "IP_SCOPE", "ALLOWED_IP", "DENIED_IP", "COMMENT", "DB_ACCOUNT")
    End If
    For Each record In records
        number = number + 1: Set row = New ATRuntimeRow: row.Initialize kind & "_RAW", CStr(number + 1)
        For i = 0 To UBound(fields)
            If IsError(record(i)) Then
                Set field = AT_RuntimeFailure(atEvalError, "EXCEL_CELL_ERROR", Nothing)
            ElseIf Len(CStr(record(i))) = 0 Then
                Set field = AT_RuntimeOK(AT_RuntimeState("TEXT", atValueEmpty))
            Else
                Set field = AT_RuntimeOK(AT_RuntimeScalar("TEXT", CStr(record(i))))
            End If
            row.AddField CStr(fields(i)), field, "'actor policy'!A" & CStr(number + 1)
        Next i
        rows.Add row
    Next record
    value.InitializeRows kind & "_RAW", rows: Set AT_TestInitiatorRaw = AT_RuntimeOK(value)
End Function

Private Sub AT_TestInitiatorBad(ByVal result As ATInitiatorPolicyResult, ByVal code As String)
    AT_Assert result.State = atEvalError And result.Diagnostics.Count = 1, "actor policy blocks " & code
    If result.State = atEvalError Then AT_Assert result.Diagnostics(1).Code = code, "actor policy stable diagnostic " & code
End Sub
