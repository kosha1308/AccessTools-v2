Attribute VB_Name = "modATRegistry"
Option Explicit
Public Function AT_RegistryText(ByVal field As ATEvalResult, ByRef valid As Boolean, ByVal budget As ATEvalBudget) As String
    Dim value As ATRuntimeValue
    budget.Step: valid = False
    If field.State <> atEvalOK Then
        If field.Diagnostics.Count = 1 Then
            If field.Diagnostics(1).Code = "EXCEL_CELL_ERROR" Then Exit Function
        End If
        Err.Raise vbObjectError + 3010, , "REGISTRY_INPUT_INVALID"
    End If
    Set value = field.Value
    If value.DataType <> "TEXT" Then Err.Raise vbObjectError + 3010, , "REGISTRY_INPUT_INVALID"
    Select Case value.ValueState
        Case atValuePresent
            AT_RegistryText = CStr(value.Scalar): budget.CheckText Len(AT_RegistryText): budget.Step Len(AT_RegistryText)
        Case atValueEmpty
        Case Else: Err.Raise vbObjectError + 3010, , "REGISTRY_INPUT_INVALID"
    End Select
    valid = True
End Function
Public Function AT_RegistryHost(ByVal text As String) As String
    text = LCase$(Trim$(text))
    If Right$(text, 1) = "." Then text = Left$(text, Len(text) - 1)
    AT_RegistryHost = text
End Function
Public Function AT_RegistryCandidate(ByVal token As String) As Boolean
    Dim i As Long, ch As String, first As String, numeric As Boolean
    first = Left$(token, 1)
    If first >= "0" And first <= "9" Then
        If InStr(token, ".") > 0 Or InStr(token, "/") > 0 Then AT_RegistryCandidate = True: Exit Function
    End If
    If Len(token) - Len(Replace(token, ":", "")) >= 2 Then AT_RegistryCandidate = True: Exit Function
    numeric = Len(token) > 0 And (InStr(token, ".") > 0 Or (first >= "0" And first <= "9"))
    For i = 1 To Len(token)
        ch = Mid$(token, i, 1)
        If InStr(1, "0123456789.xX/-", ch, vbBinaryCompare) = 0 Then numeric = False
    Next i
    AT_RegistryCandidate = numeric
End Function
Public Function AT_RegistryToken(ByVal token As String, ByRef prefix As String, ByRef manual As Boolean, ByVal budget As ATEvalBudget) As ATRuntimeValue
    Dim domain As New ATIntervalSet, declarations As New Collection, value As New ATRuntimeValue, nextPrefix As String, code As String
    On Error GoTo Failed
    domain.Initialize 4294967295#: nextPrefix = prefix
    AT_ParseIPToken token, domain, declarations, budget, nextPrefix
    domain.Freeze: value.InitializeDomain "IP_EXPRESSION", domain, declarations
    prefix = nextPrefix: Set AT_RegistryToken = value: Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "REGISTRY_INPUT_INVALID")
    If Left$(code, 7) <> "DOMAIN_" Then Err.Raise vbObjectError + 3010, , code
    manual = True: prefix = ""
End Function
Public Function AT_RegistryKey(ByVal domain As ATIntervalSet, ByVal budget As ATEvalBudget) As String
    Dim i As Long, piece As String
    For i = 1 To domain.Count
        piece = CStr(domain.Item(i).Low) & ":" & CStr(domain.Item(i).High) & ";"
        budget.Step Len(piece) + 1: budget.CheckText Len(AT_RegistryKey) + Len(piece)
        AT_RegistryKey = AT_RegistryKey & piece
    Next i
End Function
Public Function AT_RegistryDisplay(ByVal token As String, ByVal value As ATRuntimeValue) As String
    If InStr(token, "/") > 0 Then
        AT_RegistryDisplay = Mid$(CStr(value.DeclarationKeys(1)), 6)
    ElseIf InStr(AT_AsciiCase(token, False), "x") > 0 Then
        AT_RegistryDisplay = AT_AsciiCase(token, False)
    ElseIf InStr(token, "-") > 0 Then
        AT_RegistryDisplay = AT_IPText(value.Domain.Item(1).Low) & "-" & AT_IPText(value.Domain.Item(1).High)
    Else
        AT_RegistryDisplay = AT_IPText(value.Domain.Item(1).Low)
    End If
End Function
Public Sub AT_RegistryIssue(ByVal issues As Collection, ByVal code As String, ByVal severity As String, ByVal field As String, ByVal ref As String, ByVal message As String)
    Dim issue As New ATCheckIssue
    If issues.Count >= 990 Then Err.Raise vbObjectError + 3010, , "CHECK_REPORT_LIMIT"
    If Len(message) > 32767 Then Err.Raise vbObjectError + 3010, , "REGISTRY_MESSAGE_LIMIT"
    issue.Initialize "NET.IP_REGISTRY", code, severity, field, ref, message: issues.Add issue
End Sub
