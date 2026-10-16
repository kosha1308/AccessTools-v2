Attribute VB_Name = "modATPolicyCatalogTests"
Option Explicit

Public Sub AT_TestPolicyCatalog()
    Dim compiler As New ATNetworkPolicyCompiler, loaded As ATNetworkPolicyResult, policy As ATNetworkPolicy
    Dim protocols As ATEvalResult, problems As ATEvalResult, routes As ATEvalResult, common As ATEvalResult
    Dim issue As New ATCheckIssue, decision As ATNetworkPolicyDecision, items As Collection, key As String, value As Variant
    Dim caught As Boolean
    key = "NET.REQUIRED_FIELDS|ERROR|Электронная почта"
    Set protocols = AT_TestPolicyRaw("PROTOCOLS", Array(Array("Да", " tcp "), Array("TRUE", "TCP"), Array("Нет", CVErr(xlErrNA))))
    Set problems = AT_TestPolicyRaw("PROBLEMS", Array(Array("Да", key, "Предупреждение", "Поле {field}: {message} / строка {row}", "", "Нет", "Да")))
    Set routes = AT_TestPolicyRaw("ROUTING", Array(Array("Да", key, "Нет", "Да", "Нет")))
    Set common = AT_TestPolicyRaw("COMMON_TEXT", Array(Array("NO_PROBLEMS", "Готово")))
    Set loaded = compiler.Compile(protocols, problems, routes, common)
    AT_Assert loaded.State = atEvalOK And loaded.Diagnostics.Count = 0, "policy compiles four typed sources"
    Set policy = loaded.Policy
    Set items = policy.Protocols
    AT_Assert items.Count = 1 And items(1) = "TCP", "active protocols canonical and deduplicated"
    items.Remove 1
    AT_Assert policy.Protocols.Count = 1, "policy protocols returned as copied collection"
    AT_Assert policy.CommonText("NO_PROBLEMS", "fallback") = "Готово", "configured common text"
    AT_Assert policy.CommonText("UNKNOWN", "fallback") = "fallback", "explicit common-text fallback"
    issue.Initialize "NET.REQUIRED_FIELDS", "FIELD_REQUIRED", "ERROR", "EMAIL", "'Лист1'!F5", "{field} исходное"
    Set decision = policy.Decide(issue, key)
    AT_Assert decision.Issue.Severity = "WARNING" And issue.Severity = "ERROR", "configured severity without mutating original"
    AT_Assert decision.Issue.Message = "Поле электронная почта: {field} исходное / строка 5", "template expands once without interpreting inserted message"
    AT_Assert decision.CoordinatorIssue.Message = decision.Issue.Message, "coordinator falls back to selected user text"
    AT_Assert Not decision.InUserReport And decision.InRequest And Not decision.InSuperset And Not decision.UserError, "routing uses original severity key"
    AT_Assert decision.Issue.CellRef = issue.CellRef And decision.Issue.Code = issue.Code, "policy retains original trace and issue code"
    Set decision = policy.Decide(issue, "")
    AT_Assert decision.Issue.Message = issue.Message And decision.InUserReport And Not decision.InRequest And Not decision.InSuperset, "unmapped issue retains existing report behavior without new side effects"
    AT_Assert decision.UserError, "unmapped ERROR retains user-error classification"
    Set loaded = compiler.Compile(AT_TestPolicyRaw("PROTOCOLS", Array()), AT_TestPolicyRaw("PROBLEMS", Array()), AT_TestPolicyRaw("ROUTING", Array()), AT_TestPolicyRaw("COMMON_TEXT", Array()))
    AT_Assert loaded.State = atEvalOK And loaded.Policy.Protocols.Count = 0, "empty protocol table has no invented defaults"
    Set problems = AT_TestPolicyRaw("PROBLEMS", Array(Array("Нет", "bad", "bad", "{unknown}", "", "bad", "bad")))
    Set loaded = compiler.Compile(protocols, problems, routes, common)
    AT_Assert loaded.State = atEvalOK, "inactive draft row does not override policy"
    AT_TestPolicyCompileBad compiler, "PROTOCOLS", Array(Array("maybe", "TCP")), "POLICY_FLAG_INVALID"
    AT_TestPolicyCompileBad compiler, "PROTOCOLS", Array(Array("", "TCP")), "POLICY_FLAG_INVALID"
    For Each value In Array("", "TCP UDP", "TCP;UDP", "TCP|UDP")
        AT_TestPolicyCompileBad compiler, "PROTOCOLS", Array(Array("Да", value)), "POLICY_PROTOCOL_INVALID"
    Next value
    AT_TestPolicyCompileBad compiler, "PROTOCOLS", Array(Array("Да", CVErr(xlErrNA))), "POLICY_INPUT_INVALID"
    AT_TestPolicyCompileBad compiler, "PROTOCOLS", Array(Array("Да", String$(65537, "x"))), "RUNTIME_LIMIT_TEXT"
    AT_TestPolicyCompileBad compiler, "PROBLEMS", Array(Array("Да", key, "fatal", "", "", "Да", "Да")), "POLICY_SEVERITY_INVALID"
    AT_TestPolicyCompileBad compiler, "PROBLEMS", Array(Array("Да", "ERROR|*", "Ошибка", "", "", "Да", "Да")), "POLICY_KEY_INVALID"
    AT_TestPolicyCompileBad compiler, "PROBLEMS", Array(Array("Да", key, "Ошибка", "{unknown}", "", "Да", "Да")), "POLICY_TEMPLATE_INVALID"
    AT_TestPolicyCompileBad compiler, "PROBLEMS", Array(Array("Да", key, "Ошибка", "broken {message", "", "Да", "Да")), "POLICY_TEMPLATE_INVALID"
    AT_TestPolicyCompileBad compiler, "PROBLEMS", Array(Array("Да", key, "Ошибка", String$(32768, "x"), "", "Да", "Да")), "POLICY_MESSAGE_LIMIT"
    AT_TestPolicyCompileBad compiler, "PROBLEMS", Array(Array("Да", key, "Ошибка", "", "", "Да", "Да"), Array("Да", LCase$(key), "Ошибка", "", "", "Да", "Да")), "POLICY_DUPLICATE_KEY"
    AT_TestPolicyCompileBad compiler, "ROUTING", Array(Array("Да", key, "", "Да", "Нет")), "POLICY_FLAG_INVALID"
    AT_TestPolicyCompileBad compiler, "ROUTING", Array(Array("Да", key, "Да", "Да", "Нет"), Array("Да", key, "Да", "Да", "Нет")), "POLICY_DUPLICATE_KEY"
    AT_TestPolicyCompileBad compiler, "COMMON_TEXT", Array(Array("NO_PROBLEMS", "one"), Array("NO_PROBLEMS", "two")), "POLICY_DUPLICATE_KEY"
    Set loaded = compiler.Compile(Nothing, Nothing, Nothing, Nothing)
    AT_TestPolicyLoadFailure loaded, "POLICY_INPUT_INVALID"
    Set problems = AT_TestPolicyRaw("PROBLEMS", Array(Array("Да", key, "INFO", "{message}{message}", "Координатор: {check}", "Да", "Нет")))
    Set loaded = compiler.Compile(protocols, problems, routes, common)
    Set issue = New ATCheckIssue
    issue.Initialize "NET.REQUIRED_FIELDS", "FIELD_REQUIRED", "ERROR", "EMAIL", "", String$(20000, "x")
    On Error Resume Next
    Set decision = loaded.Policy.Decide(issue, key)
    caught = Err.Description = "POLICY_MESSAGE_LIMIT": Err.Clear
    On Error GoTo 0
    AT_Assert caught, "expanded message limit checked before publishing decision"
End Sub

Public Function AT_TestPolicyRaw(ByVal kind As String, ByVal records As Variant) As ATEvalResult
    Dim fields As Variant, record As Variant, row As ATRuntimeRow, rows As New Collection, value As New ATRuntimeValue
    Dim i As Long, number As Long, field As ATEvalResult
    Select Case kind
        Case "PROTOCOLS": fields = Array("ACTIVE", "PROTOCOL")
        Case "PROBLEMS": fields = Array("ACTIVE", "KEY", "SEVERITY", "USER_MESSAGE", "COORDINATOR_MESSAGE", "USER_ERROR", "HIGHLIGHT")
        Case "ROUTING": fields = Array("ACTIVE", "KEY", "IN_USER_REPORT", "IN_REQUEST", "IN_SUPERSET")
        Case "COMMON_TEXT": fields = Array("KEY", "VALUE")
    End Select
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
            row.AddField CStr(fields(i)), field, "'policy'!A" & CStr(number + 1)
        Next i
        rows.Add row
    Next record
    value.InitializeRows kind & "_RAW", rows: Set AT_TestPolicyRaw = AT_RuntimeOK(value)
End Function

Private Sub AT_TestPolicyCompileBad(ByVal compiler As ATNetworkPolicyCompiler, ByVal kind As String, ByVal records As Variant, ByVal code As String)
    Dim p As ATEvalResult, m As ATEvalResult, r As ATEvalResult, c As ATEvalResult
    Set p = AT_TestPolicyRaw("PROTOCOLS", Array()): Set m = AT_TestPolicyRaw("PROBLEMS", Array())
    Set r = AT_TestPolicyRaw("ROUTING", Array()): Set c = AT_TestPolicyRaw("COMMON_TEXT", Array())
    Select Case kind
        Case "PROTOCOLS": Set p = AT_TestPolicyRaw(kind, records)
        Case "PROBLEMS": Set m = AT_TestPolicyRaw(kind, records)
        Case "ROUTING": Set r = AT_TestPolicyRaw(kind, records)
        Case "COMMON_TEXT": Set c = AT_TestPolicyRaw(kind, records)
    End Select
    AT_TestPolicyLoadFailure compiler.Compile(p, m, r, c), code
End Sub

Private Sub AT_TestPolicyLoadFailure(ByVal loaded As ATNetworkPolicyResult, ByVal code As String)
    Dim partial As ATNetworkPolicy, number As Long
    AT_Assert loaded.State = atEvalError And loaded.Diagnostics.Count = 1, "policy fails closed " & code
    AT_Assert loaded.Diagnostics(1).Code = code, "policy stable error " & code
    On Error Resume Next
    Set partial = loaded.Policy: number = Err.Number
    On Error GoTo 0
    AT_Assert number <> 0 And partial Is Nothing, "invalid policy inaccessible " & code
End Sub
