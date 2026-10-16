Attribute VB_Name = "modATDomainParse"
Option Explicit

Public Function AT_RuntimeDomainType(ByVal typeName As String) As Boolean
    Select Case typeName
        Case "IP", "IP_EXPRESSION", "PORT", "PORT_SET": AT_RuntimeDomainType = True
    End Select
End Function

Public Function AT_ParseDomain(ByVal typeName As String, ByVal raw As String, ByVal node As ATAstNode, Optional ByVal budget As ATEvalBudget = Nothing) As ATEvalResult
    Dim limits As New ATEvalLimits, intervals As New ATIntervalSet, keys As New Collection, tokens As Collection
    Dim token As Variant, prefix As String, value As New ATRuntimeValue, code As String, ports As Boolean
    On Error GoTo Failed
    If budget Is Nothing Then
        Set budget = New ATEvalBudget
        budget.Initialize limits
    End If
    budget.CheckText Len(raw)
    If Not AT_RuntimeDomainType(typeName) Then
        Set AT_ParseDomain = AT_RuntimeFailure(atEvalError, "RUNTIME_UNSUPPORTED_NORMALIZER", node)
        Exit Function
    End If
    ports = (typeName = "PORT" Or typeName = "PORT_SET")
    intervals.Initialize IIf(ports, 65535#, 4294967295#)
    If Len(Trim$(raw)) = 0 Then
        Set AT_ParseDomain = AT_RuntimeOK(AT_RuntimeState(typeName, atValueEmpty))
        Exit Function
    End If
    Set tokens = AT_DomainTokens(raw, ports, budget)
    If tokens.Count = 0 Then Err.Raise vbObjectError + 2610, , "DOMAIN_EMPTY_TOKENS"
    If typeName = "IP" Then
        If tokens.Count <> 1 Then Err.Raise vbObjectError + 2611, , "DOMAIN_INVALID_IPV4"
        AT_AddDomainRange intervals, keys, AT_ParseIPv4(CStr(tokens(1))), AT_ParseIPv4(CStr(tokens(1))), budget
    ElseIf typeName = "PORT" Then
        If tokens.Count <> 1 Then Err.Raise vbObjectError + 2612, , "DOMAIN_INVALID_PORT"
        AT_AddDomainRange intervals, keys, PortNumber(CStr(tokens(1))), PortNumber(CStr(tokens(1))), budget
    Else
        For Each token In tokens
            budget.Step
            If ports Then
                ParsePortToken CStr(token), intervals, keys, budget
            Else
                AT_ParseIPToken CStr(token), intervals, keys, budget, prefix
            End If
        Next token
    End If
    intervals.Freeze
    value.InitializeDomain typeName, intervals, keys
    Set AT_ParseDomain = AT_RuntimeOK(value)
    Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "RUNTIME_INTERNAL_ERROR")
    If Left$(code, 7) = "DOMAIN_" Then
        Set AT_ParseDomain = AT_RuntimeInvalid(typeName, code, node)
    Else
        Set AT_ParseDomain = AT_RuntimeFailure(atEvalError, code, node)
    End If
End Function

Public Function AT_DomainTokens(ByVal raw As String, ByVal ports As Boolean, ByVal budget As ATEvalBudget) As Collection
    Dim normalized As String, i As Long, ch As String, previous As String, output As String, length As Long, pending As Boolean
    Dim result As New Collection, token As Variant
    normalized = raw
    For i = 1 To Len(raw)
        budget.Step
        ch = Mid$(raw, i, 1)
        Select Case AscW(ch)
            Case &H2010 To &H2015, &H2212: Mid$(normalized, i, 1) = "-"
        End Select
    Next i
    raw = normalized
    output = Space$(Len(raw))
    For i = 1 To Len(raw)
        budget.Step
        ch = Mid$(raw, i, 1)
        If ch = vbTab Or ch = vbCr Or ch = vbLf Or ch = " " Then
            pending = True
        Else
            If ch = ";" Or ch = "|" Then ch = ","
            If ports And (ch = "/" Or ch = "\") Then ch = ","
            previous = ""
            If length > 0 Then previous = Mid$(output, length, 1)
            If pending And length > 0 And previous <> "-" And ch <> "-" And previous <> "," And ch <> "," Then
                length = length + 1
                Mid$(output, length, 1) = ","
                previous = ","
            End If
            pending = False
            If ch <> "," Or previous <> "," Then
                length = length + 1
                Mid$(output, length, 1) = ch
            End If
        End If
    Next i
    normalized = Left$(output, length)
    For Each token In Split(normalized, ",")
        If Len(CStr(token)) > 0 Then
            budget.CheckList result.Count + 1
            result.Add CStr(token)
        End If
    Next token
    Set AT_DomainTokens = result
End Function

Public Function AT_Digits(ByVal text As String, ByVal maxLength As Long) As Boolean
    Dim i As Long, code As Long
    If Len(text) < 1 Or Len(text) > maxLength Then Exit Function
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1))
        If code < 48 Or code > 57 Then Exit Function
    Next i
    AT_Digits = True
End Function

Private Function PortNumber(ByVal text As String) As Double
    If Not AT_Digits(text, 5) Then Err.Raise vbObjectError + 2612, , "DOMAIN_INVALID_PORT"
    PortNumber = Val(text)
    If PortNumber > 65535# Then Err.Raise vbObjectError + 2612, , "DOMAIN_INVALID_PORT"
End Function

Private Sub ParsePortToken(ByVal token As String, ByVal intervals As ATIntervalSet, ByVal keys As Collection, ByVal budget As ATEvalBudget)
    Dim parts As Variant, low As Double, high As Double
    parts = Split(token, "-")
    If UBound(parts) > 1 Then Err.Raise vbObjectError + 2612, , "DOMAIN_INVALID_PORT"
    low = PortNumber(CStr(parts(0))): high = low
    If UBound(parts) = 1 Then high = PortNumber(CStr(parts(1)))
    If high < low Then Err.Raise vbObjectError + 2612, , "DOMAIN_INVALID_PORT"
    AT_AddDomainRange intervals, keys, low, high, budget
End Sub

Public Sub AT_AddDomainKey(ByVal keys As Collection, ByVal key As String, ByVal budget As ATEvalBudget)
    budget.Step
    If AT_HasDomainKey(keys, key) Then Exit Sub
    budget.CheckList keys.Count + 1
    keys.Add key, key
End Sub
Public Function AT_HasDomainKey(ByVal keys As Collection, ByVal key As String) As Boolean
    Dim ignored As String
    On Error Resume Next
    ignored = keys(key)
    AT_HasDomainKey = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
Public Sub AT_AddDomainRange(ByVal intervals As ATIntervalSet, ByVal keys As Collection, ByVal low As Double, ByVal high As Double, ByVal budget As ATEvalBudget)
    intervals.Add low, high, budget
    AT_AddDomainKey keys, "RANGE:" & Format$(low, "0") & ":" & Format$(high, "0"), budget
End Sub
