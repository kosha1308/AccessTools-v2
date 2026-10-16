VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkPolicy"
Option Explicit
Private mReady As Boolean, mProtocols As Collection, mProblems As Object, mRoutes As Object, mText As Object

Friend Sub Initialize(ByVal protocols As ATEvalResult, ByVal problems As ATEvalResult, ByVal routing As ATEvalResult, ByVal commonText As ATEvalResult)
    Dim budget As New ATEvalBudget, limits As New ATEvalLimits, source As ATRuntimeValue, row As ATRuntimeRow
    Dim seen As Object, i As Long, j As Long, key As String, text As String, values As Variant, ignored As String
    If mReady Then Err.Raise vbObjectError + 2960, , "POLICY_FROZEN"
    budget.Initialize limits
    Set mProtocols = New Collection: Set seen = CreateObject("Scripting.Dictionary")
    Set mProblems = CreateObject("Scripting.Dictionary"): Set mRoutes = CreateObject("Scripting.Dictionary"): Set mText = CreateObject("Scripting.Dictionary")
    Set source = SourceRows(protocols, "PROTOCOLS", budget)
    For i = 1 To source.Count
        Set row = source.Row(i)
        If AT_PolicyBoolean(AT_PolicyCell(row, "ACTIVE", budget)) Then
            text = AT_AsciiCase(Trim$(AT_PolicyCell(row, "PROTOCOL", budget)), True)
            If Len(text) = 0 Or Len(text) > 128 Then Err.Raise vbObjectError + 2960, , "POLICY_PROTOCOL_INVALID"
            For j = 1 To Len(text)
                If AscW(Mid$(text, j, 1)) < 33 Or AscW(Mid$(text, j, 1)) > 126 Or InStr(",;|", Mid$(text, j, 1)) > 0 Then Err.Raise vbObjectError + 2960, , "POLICY_PROTOCOL_INVALID"
            Next j
            If Not seen.Exists(text) Then seen.Add text, True: mProtocols.Add text
        End If
    Next i
    Set source = SourceRows(problems, "PROBLEMS", budget)
    For i = 1 To source.Count
        Set row = source.Row(i)
        If AT_PolicyBoolean(AT_PolicyCell(row, "ACTIVE", budget)) Then
            key = AT_PolicyKey(AT_PolicyCell(row, "KEY", budget))
            values = Array(AT_PolicyCell(row, "USER_MESSAGE", budget), AT_PolicyCell(row, "COORDINATOR_MESSAGE", budget), AT_PolicySeverity(AT_PolicyCell(row, "SEVERITY", budget)), AT_PolicyBoolean(AT_PolicyCell(row, "USER_ERROR", budget)), AT_PolicyBoolean(AT_PolicyCell(row, "HIGHLIGHT", budget)))
            ignored = AT_PolicyExpand(CStr(values(0)), Nothing, ""): ignored = AT_PolicyExpand(CStr(values(1)), Nothing, "")
            AddUnique mProblems, key, values
        End If
    Next i
    Set source = SourceRows(routing, "ROUTING", budget)
    For i = 1 To source.Count
        Set row = source.Row(i)
        If AT_PolicyBoolean(AT_PolicyCell(row, "ACTIVE", budget)) Then
            key = AT_PolicyKey(AT_PolicyCell(row, "KEY", budget))
            values = Array(AT_PolicyBoolean(AT_PolicyCell(row, "IN_USER_REPORT", budget)), AT_PolicyBoolean(AT_PolicyCell(row, "IN_REQUEST", budget)), AT_PolicyBoolean(AT_PolicyCell(row, "IN_SUPERSET", budget)))
            AddUnique mRoutes, key, values
        End If
    Next i
    Set source = SourceRows(commonText, "COMMON_TEXT", budget)
    For i = 1 To source.Count
        Set row = source.Row(i): text = Trim$(AT_PolicyCell(row, "KEY", budget))
        If Len(text) = 0 Or Len(text) > 128 Then Err.Raise vbObjectError + 2960, , "POLICY_KEY_INVALID"
        key = AT_CanonicalId(text): text = AT_PolicyCell(row, "VALUE", budget)
        If Len(text) > 32767 Then Err.Raise vbObjectError + 2960, , "POLICY_MESSAGE_LIMIT"
        AddUnique mText, key, text
    Next i
    mReady = True
End Sub

Private Function SourceRows(ByVal raw As ATEvalResult, ByVal kind As String, ByVal budget As ATEvalBudget) As ATRuntimeValue
    Dim value As ATRuntimeValue
    If raw Is Nothing Then GoTo Invalid
    If raw.State <> atEvalOK Then GoTo Invalid
    Set value = raw.Value
    If value.DataType <> "COLLECTION<ROW<" & kind & "_RAW>>" Then GoTo Invalid
    budget.CheckList value.Count: budget.Step value.Count
    Set SourceRows = value
    Exit Function
Invalid:
    Err.Raise vbObjectError + 2960, , "POLICY_INPUT_INVALID"
End Function
Private Sub AddUnique(ByVal target As Object, ByVal key As String, ByVal value As Variant)
    If target.Exists(key) Then Err.Raise vbObjectError + 2960, , "POLICY_DUPLICATE_KEY"
    target.Add key, value
End Sub
Private Sub RequireReady()
    If Not mReady Then Err.Raise vbObjectError + 2960, , "POLICY_ACCESS"
End Sub
Public Property Get Protocols() As Collection
    Dim copy As New Collection, item As Variant
    RequireReady
    For Each item In mProtocols: copy.Add CStr(item): Next item
    Set Protocols = copy
End Property
Public Function CommonText(ByVal key As String, ByVal fallback As String) As String
    RequireReady
    key = AT_CanonicalId(key)
    CommonText = fallback
    If mText.Exists(key) Then CommonText = CStr(mText(key))
End Function

Public Function Decide(ByVal original As ATCheckIssue, ByVal key As String) As ATNetworkPolicyDecision
    Dim result As New ATNetworkPolicyDecision, issue As New ATCheckIssue, coordinator As New ATCheckIssue
    Dim message As String, coordinatorText As String, severity As String, values As Variant, field As String
    Dim user As Boolean, request As Boolean, superset As Boolean, userError As Boolean, highlight As Boolean
    RequireReady
    If original Is Nothing Then Err.Raise vbObjectError + 2960, , "POLICY_INPUT_INVALID"
    message = original.Message: severity = original.Severity
    user = True: userError = (severity = "ERROR")
    If Len(key) > 0 Then
        key = AT_PolicyKey(key): values = Split(key, "|"): field = CStr(values(2))
        If CStr(values(1)) <> original.Severity Then Err.Raise vbObjectError + 2960, , "POLICY_KEY_INVALID"
        If mProblems.Exists(key) Then
            values = mProblems(key)
            If Not AT_ExcelBlank(values(0)) Then message = AT_PolicyExpand(CStr(values(0)), original, field)
            severity = CStr(values(2)): userError = CBool(values(3)): highlight = CBool(values(4))
            If Not AT_ExcelBlank(values(1)) Then coordinatorText = AT_PolicyExpand(CStr(values(1)), original, field)
        End If
        If mRoutes.Exists(key) Then
            values = mRoutes(key): user = CBool(values(0)): request = CBool(values(1)) And highlight: superset = CBool(values(2))
        End If
    End If
    If Len(coordinatorText) = 0 Then coordinatorText = message
    issue.Initialize original.CheckId, original.Code, severity, original.FieldId, original.CellRef, message
    coordinator.Initialize original.CheckId, original.Code, severity, original.FieldId, original.CellRef, coordinatorText
    result.Initialize issue, coordinator, key, user, request, superset, userError
    Set Decide = result
End Function
