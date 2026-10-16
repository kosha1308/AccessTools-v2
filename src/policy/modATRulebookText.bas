Attribute VB_Name = "modATRulebookText"
Option Explicit
Public Function AT_RuleTokens(ByVal text As String, ByVal host As Boolean, ByVal budget As ATEvalBudget) As Collection
    Dim r As New Collection, token As Variant, s As String, i As Long
    budget.CheckText Len(text): budget.Step Len(text) + 1
    AT_RuleSeparators text
    text = Replace(Replace(Replace(Replace(text, ";", ","), "|", ","), vbCr, ","), vbLf, ",")
    For Each token In Split(text, ",")
        s = Trim$(CStr(token))
        If Len(s) > 0 Then
            s = LCase$(s)
            If host Then
                s = AT_RuleHostname(s)
            End If
            budget.CheckList r.Count + 1: r.Add s
        End If
    Next token
    If Not AT_ExcelBlank(text) And r.Count = 0 Then Err.Raise vbObjectError + 2980, , "RULEBOOK_LIST_INVALID"
    Set AT_RuleTokens = r
End Function
Public Sub AT_RuleSeparators(ByVal text As String)
    If InStr(1, text, "|", vbBinaryCompare) > 0 Then Err.Raise vbObjectError + 2980, , "RULEBOOK_SEPARATOR_UNSUPPORTED"
End Sub
Public Function AT_RuleHostname(ByVal text As String) As String
    Dim i As Long, ch As String
    text = LCase$(Trim$(text))
    If Right$(text, 1) = "." Then text = Left$(text, Len(text) - 1)
    If Len(text) = 0 Or Len(text) > 253 Then Err.Raise vbObjectError + 2980, , "RULEBOOK_HOSTNAME_INVALID"
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If AscW(ch) >= 0 And AscW(ch) <= 32 Then Err.Raise vbObjectError + 2980, , "RULEBOOK_HOSTNAME_INVALID"
        If InStr(1, ",;|", ch, vbBinaryCompare) > 0 Then Err.Raise vbObjectError + 2980, , "RULEBOOK_HOSTNAME_INVALID"
    Next i
    AT_RuleHostname = text
End Function
Public Function AT_RuleDomain(ByVal text As String, ByVal ports As Boolean, ByVal budget As ATEvalBudget) As ATIntervalSet
    Dim r As ATEvalResult, i As Long
    AT_RuleSeparators text
    Set r = AT_ParseDomain(IIf(ports, "PORT_SET", "IP_EXPRESSION"), text, Nothing, budget)
    If r.State <> atEvalOK Then Err.Raise vbObjectError + 2980, , r.Diagnostics(1).Code
    If r.Value.ValueState <> atValuePresent Then Err.Raise vbObjectError + 2980, , "RULEBOOK_DOMAIN_INVALID"
    If ports Then
        For i = 1 To r.Value.Domain.Count
            If r.Value.Domain.Item(i).Low < 1 Then Err.Raise vbObjectError + 2980, , "RULEBOOK_PORT_INVALID"
        Next i
    End If
    Set AT_RuleDomain = r.Value.Domain
End Function
Public Function AT_RuleProtocol(ByVal text As String) As String
    Dim i As Long, ch As String
    text = AT_AsciiCase(Trim$(text), True)
    If Len(text) = 0 Or Len(text) > 128 Then GoTo Invalid
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If Not ((ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or InStr(1, "_-./+", ch, vbBinaryCompare) > 0) Then GoTo Invalid
    Next i
    AT_RuleProtocol = text: Exit Function
Invalid:
    Err.Raise vbObjectError + 2980, , "RULEBOOK_PROTOCOL_INVALID"
End Function
Public Function AT_RuleFailure(ByVal description As String, Optional ByVal invalidContext As Boolean = False) As ATEvalResult
    Dim code As String
    code = AT_StableErrorCode(description, "RULEBOOK_INPUT_INVALID")
    If Left$(code, 14) = "RUNTIME_LIMIT_" Then
        Set AT_RuleFailure = AT_RuntimeFailure(atEvalError, code, Nothing)
    ElseIf invalidContext Then
        Set AT_RuleFailure = AT_RuntimeFailure(atEvalManualReview, "RULEBOOK_CONTEXT_INCOMPLETE", Nothing)
    Else
        If Left$(code, 9) <> "RULEBOOK_" Then code = "RULEBOOK_INPUT_INVALID"
        Set AT_RuleFailure = AT_RuntimeFailure(atEvalError, code, Nothing)
    End If
End Function
