VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkRulebook"
Option Explicit
Private mReady As Boolean, mRules As Collection
Friend Sub Initialize(ByVal raw As ATEvalResult, ByVal protocols As Collection)
    Dim budget As New ATEvalBudget, limits As New ATEvalLimits, rows As ATRuntimeValue, row As ATRuntimeRow
    Dim allowed As Object, token As Variant, rule As ATNetworkRule, i As Long, active As String
    If mReady Then Err.Raise vbObjectError + 2980, , "RULEBOOK_FROZEN"
    budget.Initialize limits: Set mRules = New Collection
    If protocols Is Nothing Then Err.Raise vbObjectError + 2980, , "RULEBOOK_PROTOCOL_INVALID"
    If protocols.Count > 1000 Then Err.Raise vbObjectError + 2980, , "RULEBOOK_PROTOCOL_INVALID"
    Set allowed = CreateObject("Scripting.Dictionary")
    For Each token In protocols
        If VarType(token) <> vbString Then Err.Raise vbObjectError + 2980, , "RULEBOOK_PROTOCOL_INVALID"
        budget.CheckText Len(CStr(token)): budget.Step Len(CStr(token)) + 1
        allowed(AT_RuleProtocol(CStr(token))) = True
    Next token
    Set rows = AT_InitiatorRows(raw, "RULES", budget)
    For i = 1 To rows.Count
        Set row = rows.Row(i): active = AT_PolicyCell(row, "ACTIVE", budget)
        If Not AT_ExcelBlank(active) Then
            If AT_PolicyBoolean(active) Then
                Set rule = New ATNetworkRule: rule.Initialize row, allowed, budget: mRules.Add rule
            End If
        End If
    Next i
    mReady = True
End Sub
Public Function Technical(ByVal source As String, ByVal destination As String, ByVal ports As String, ByVal protocol As String, ByVal budget As ATEvalBudget) As ATEvalResult
    Dim src As ATIntervalSet, dst As ATIntervalSet, portSet As ATIntervalSet, rule As ATNetworkRule, matches As New Collection, parsing As Boolean
    On Error GoTo Failed
    RequireReady
    If budget Is Nothing Then Err.Raise vbObjectError + 2980, , "RULEBOOK_BUDGET_REQUIRED"
    parsing = True
    Set src = AT_RuleDomain(source, False, budget): Set dst = AT_RuleDomain(destination, False, budget)
    Set portSet = AT_RuleDomain(ports, True, budget): protocol = AT_RuleProtocol(protocol)
    budget.CheckText Len(protocol): budget.Step Len(protocol) + 1
    parsing = False
    For Each rule In mRules
        If rule.TechnicalBase(src, dst, portSet, protocol, budget) Then
            If rule.HasHostname Then
                Set Technical = AT_RuntimeFailure(atEvalManualReview, "RULEBOOK_HOSTNAME_REQUIRED", Nothing): Exit Function
            End If
            matches.Add rule
        End If
    Next rule
    Set Technical = Outputs(matches, budget): Exit Function
Failed:
    Set Technical = AT_RuleFailure(Err.Description, parsing)
End Function
Public Function Superset(ByVal ip As String, ByVal hostname As String, ByVal service As String, ByVal budget As ATEvalBudget) As ATEvalResult
    Dim addresses As ATIntervalSet, rule As ATNetworkRule, candidates As New Collection, matches As New Collection
    Dim score As Long, best As Long, pair As Variant, parsing As Boolean
    On Error GoTo Failed
    RequireReady
    If budget Is Nothing Then Err.Raise vbObjectError + 2980, , "RULEBOOK_BUDGET_REQUIRED"
    parsing = True
    budget.CheckText Len(hostname): budget.CheckText Len(service): budget.Step Len(hostname) + Len(service) + 1
    Set addresses = AT_RuleDomain(ip, False, budget)
    parsing = False
    For Each rule In mRules
        If rule.SupersetBase(addresses, service, budget) Then
            score = 0
            If rule.HasHostname Then
                If AT_ExcelBlank(hostname) Then
                    Set Superset = AT_RuntimeFailure(atEvalManualReview, "RULEBOOK_HOSTNAME_REQUIRED", Nothing): Exit Function
                End If
                score = rule.Specificity(hostname, budget)
                If score > best Then best = score
                If score > 0 Then candidates.Add Array(rule, score)
            Else
                candidates.Add Array(rule, 0)
            End If
        End If
    Next rule
    For Each pair In candidates
        budget.Step
        If CLng(pair(1)) = 0 Or CLng(pair(1)) = best Then matches.Add pair(0)
    Next pair
    Set Superset = Outputs(matches, budget): Exit Function
Failed:
    Set Superset = AT_RuleFailure(Err.Description, parsing)
End Function
Private Sub RequireReady()
    If Not mReady Then Err.Raise vbObjectError + 2980, , "RULEBOOK_ACCESS"
End Sub
Private Function Outputs(ByVal matches As Collection, ByVal budget As ATEvalBudget) As ATEvalResult
    Dim rows As New Collection, value As New ATRuntimeValue, rule As ATNetworkRule, field As Variant
    Dim first As String, text As String, conflict As Boolean
    For Each field In Array("SERVICE_RESULT", "DEST_GROUP")
        first = ""
        For Each rule In matches
            text = rule.ConflictValue(CStr(field), budget)
            If Len(text) > 0 Then
                If Len(first) > 0 And text <> first Then conflict = True
                first = text
            End If
        Next rule
    Next field
    For Each rule In matches
        budget.CheckList rows.Count + 1: rows.Add rule.Output(conflict, budget)
    Next rule
    value.InitializeRows "RULE_MATCH", rows: Set Outputs = AT_RuntimeOK(value)
End Function
