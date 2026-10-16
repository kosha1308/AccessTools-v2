Attribute VB_Name = "modATInitiatorText"
Option Explicit
Public Function AT_InitiatorTokens(ByVal text As String, ByVal budget As ATEvalBudget) As Collection
    Dim result As New Collection, i As Long, ch As String, start As Long
    budget.CheckText Len(text)
    For i = 1 To Len(text) + 1
        budget.Step: ch = Mid$(text, i, 1)
        If Len(ch) = 0 Or InStr(1, ",;| " & vbTab & vbCr & vbLf, ch, vbBinaryCompare) > 0 Then
            If start > 0 Then budget.CheckList result.Count + 1: result.Add Mid$(text, start, i - start): start = 0
        ElseIf start = 0 Then
            start = i
        End If
    Next i
    Set AT_InitiatorTokens = result
End Function
Public Function AT_InitiatorNegative(ByRef token As String) As Boolean
    If AT_AsciiCase(Left$(token, 4), True) = "NOT:" Then
        token = Mid$(token, 5): AT_InitiatorNegative = True
    ElseIf LCase$(Left$(token, 3)) = "не:" Then
        token = Mid$(token, 4): AT_InitiatorNegative = True
    End If
    If Len(token) = 0 Then Err.Raise vbObjectError + 2970, , "INITIATOR_SELECTOR_INVALID"
End Function
Public Function AT_InitiatorScope(ByVal text As String) As String
    Select Case AT_ExcelHeader(text)
        Case "", "любой", "any": AT_InitiatorScope = "ANY"
        Case "источник", "source": AT_InitiatorScope = "SOURCE"
        Case "назначение", "destination": AT_InitiatorScope = "DESTINATION"
        Case Else: Err.Raise vbObjectError + 2970, , "INITIATOR_SCOPE_INVALID"
    End Select
End Function
Public Function AT_InitiatorIPPair(ByVal allowText As String, ByVal denyText As String, ByVal budget As ATEvalBudget) As Variant
    Dim token As Variant, text As String, allowJoined As String, notJoined As String, ignored As Boolean
    Dim allowed As ATIntervalSet, denied As New ATIntervalSet, embedded As ATIntervalSet, explicitDeny As ATIntervalSet
    allowText = Replace(Replace(allowText, "х", "x"), "Х", "x")
    denyText = Replace(Replace(denyText, "х", "x"), "Х", "x")
    For Each token In AT_DomainTokens(allowText, False, budget)
        text = CStr(token)
        If AT_InitiatorNegative(text) Then
            notJoined = notJoined & text & ","
        Else
            allowJoined = allowJoined & text & ","
        End If
    Next token
    Set allowed = AT_InitiatorDomain(allowJoined, budget)
    Set embedded = AT_InitiatorDomain(notJoined, budget): Set explicitDeny = AT_InitiatorDomain(denyText, budget)
    denied.Initialize 4294967295#
    AT_InitiatorUnion denied, embedded, budget: AT_InitiatorUnion denied, explicitDeny, budget
    denied.Freeze
    AT_InitiatorIPPair = Array(allowed, denied)
End Function
Public Function AT_InitiatorDomain(ByVal text As String, ByVal budget As ATEvalBudget) As ATIntervalSet
    Dim result As ATEvalResult, emptySet As New ATIntervalSet
    Set result = AT_ParseDomain("IP_EXPRESSION", text, Nothing, budget)
    If result.State <> atEvalOK Then Err.Raise vbObjectError + 2970, , result.Diagnostics(1).Code
    If result.Value.ValueState = atValueEmpty Then
        emptySet.Initialize 4294967295#: emptySet.Freeze: Set AT_InitiatorDomain = emptySet
    ElseIf result.Value.ValueState = atValuePresent Then
        Set AT_InitiatorDomain = result.Value.Domain
    Else
        Err.Raise vbObjectError + 2970, , "INITIATOR_IP_INVALID"
    End If
End Function
Public Sub AT_InitiatorUnion(ByVal target As ATIntervalSet, ByVal source As ATIntervalSet, ByVal budget As ATEvalBudget)
    Dim i As Long
    For i = 1 To source.Count
        budget.Step: target.Add source.Item(i).Low, source.Item(i).High, budget
    Next i
End Sub
Public Function AT_InitiatorRows(ByVal raw As ATEvalResult, ByVal kind As String, ByVal budget As ATEvalBudget) As ATRuntimeValue
    Dim value As ATRuntimeValue
    If raw Is Nothing Then GoTo Invalid
    If raw.State <> atEvalOK Then GoTo Invalid
    Set value = raw.Value
    If value.DataType <> "COLLECTION<ROW<" & kind & "_RAW>>" Then GoTo Invalid
    budget.CheckList value.Count: budget.Step value.Count
    Set AT_InitiatorRows = value
    Exit Function
Invalid:
    Err.Raise vbObjectError + 2970, , "INITIATOR_INPUT_INVALID"
End Function
Public Sub AT_InitiatorIssue(ByVal issues As Collection, ByVal check As String, ByVal code As String, ByVal severity As String, ByVal field As String, ByVal ref As String, ByVal message As String)
    Dim issue As New ATCheckIssue
    If Len(message) > 32767 Then Err.Raise vbObjectError + 2970, , "INITIATOR_MESSAGE_LIMIT"
    If issues.Count >= 990 Then Err.Raise vbObjectError + 2970, , "CHECK_REPORT_LIMIT"
    issue.Initialize check, code, severity, field, ref, message: issues.Add issue
End Sub
