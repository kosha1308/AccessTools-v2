VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkRule"
Option Explicit
Private mReady As Boolean, mRaw As ATRuntimeRow, mIP As ATIntervalSet, mPorts As ATIntervalSet
Private mScope As String, mMode As Long, mHosts As Collection, mServices As Collection, mProtocols As Object
Private mOutputs As Variant, mSeverity As String
Friend Sub Initialize(ByVal raw As ATRuntimeRow, ByVal protocols As Object, ByVal budget As ATEvalBudget)
    Dim ip As String, scope As String, host As String, mode As String, port As String, protocol As String, service As String
    Dim id As Variant, value As String, hasOutput As Boolean, token As Variant
    If mReady Then Err.Raise vbObjectError + 2980, , "RULEBOOK_FROZEN"
    Set mRaw = raw.Snapshot
    ip = AT_PolicyCell(raw, "IP", budget): scope = AT_PolicyCell(raw, "IP_SCOPE", budget)
    host = AT_PolicyCell(raw, "HOSTNAME", budget): mode = AT_PolicyCell(raw, "HOSTNAME_MATCH", budget)
    port = AT_PolicyCell(raw, "DEST_PORT", budget): protocol = AT_PolicyCell(raw, "DEST_PROTOCOL", budget)
    service = AT_PolicyCell(raw, "SERVICE_CONDITION", budget)
    If AT_ExcelBlank(ip) And AT_ExcelBlank(host) And AT_ExcelBlank(port) And AT_ExcelBlank(protocol) And AT_ExcelBlank(service) Then Err.Raise vbObjectError + 2980, , "RULEBOOK_CONDITION_REQUIRED"
    If AT_ExcelBlank(ip) And Not AT_ExcelBlank(scope) Then Err.Raise vbObjectError + 2980, , "RULEBOOK_SCOPE_INVALID"
    mScope = AT_InitiatorScope(scope)
    If Not AT_ExcelBlank(ip) Then Set mIP = AT_RuleDomain(ip, False, budget)
    If Not AT_ExcelBlank(port) Then Set mPorts = AT_RuleDomain(port, True, budget)
    Set mProtocols = CreateObject("Scripting.Dictionary")
    For Each token In AT_InitiatorTokens(protocol, budget)
        value = AT_RuleProtocol(CStr(token))
        If Not protocols.Exists(value) Then Err.Raise vbObjectError + 2980, , "RULEBOOK_PROTOCOL_INVALID"
        mProtocols(value) = True
    Next token
    If Not AT_ExcelBlank(protocol) And mProtocols.Count = 0 Then Err.Raise vbObjectError + 2980, , "RULEBOOK_PROTOCOL_INVALID"
    Set mHosts = AT_RuleTokens(host, True, budget): Set mServices = AT_RuleTokens(service, False, budget)
    If mHosts.Count = 0 And Not AT_ExcelBlank(mode) Then Err.Raise vbObjectError + 2980, , "RULEBOOK_HOSTNAME_MODE_INVALID"
    Select Case AT_ExcelHeader(mode)
        Case "", "точное", "exact": mMode = 3
        Case "начинается с", "starts", "startswith", "prefix": mMode = 2
        Case "содержит", "contains": mMode = 1
        Case Else: Err.Raise vbObjectError + 2980, , "RULEBOOK_HOSTNAME_MODE_INVALID"
    End Select
    If mServices.Count > 0 Then
        If Not mPorts Is Nothing Or mProtocols.Count > 0 Then Err.Raise vbObjectError + 2980, , "RULEBOOK_CONTEXT_COMBINATION"
    End If
    mOutputs = Array("RULE_NAME", "SERVICE_RESULT", "DEST_GROUP", "APPROVAL", "COMMENT", "SEVERITY", "INSTRUCTION")
    For Each id In mOutputs
        value = AT_PolicyCell(raw, CStr(id), budget)
        If Len(value) > 32767 Then Err.Raise vbObjectError + 2980, , "RULEBOOK_OUTPUT_LIMIT"
        If CStr(id) <> "RULE_NAME" And CStr(id) <> "SEVERITY" And Not AT_ExcelBlank(value) Then hasOutput = True
        If CStr(id) = "SEVERITY" Then
            If AT_ExcelBlank(value) Then mSeverity = "INFO" Else mSeverity = AT_PolicySeverity(value)
        End If
    Next id
    If Not hasOutput Then Err.Raise vbObjectError + 2980, , "RULEBOOK_OUTPUT_REQUIRED"
    mReady = True
End Sub
Friend Function TechnicalBase(ByVal source As ATIntervalSet, ByVal destination As ATIntervalSet, ByVal ports As ATIntervalSet, ByVal protocol As String, ByVal budget As ATEvalBudget) As Boolean
    budget.Step
    If mServices.Count > 0 Then Exit Function
    If Not mIP Is Nothing Then
        Select Case mScope
            Case "SOURCE": If Not source.Overlaps(mIP, budget) Then Exit Function
            Case "DESTINATION": If Not destination.Overlaps(mIP, budget) Then Exit Function
            Case Else
                If Not source.Overlaps(mIP, budget) Then
                    If Not destination.Overlaps(mIP, budget) Then Exit Function
                End If
        End Select
    End If
    If Not mPorts Is Nothing Then
        If Not ports.IsSubsetOf(mPorts, budget) Then Exit Function
    End If
    If mProtocols.Count > 0 Then
        If Not mProtocols.Exists(protocol) Then Exit Function
    End If
    TechnicalBase = True
End Function
Friend Function SupersetBase(ByVal ip As ATIntervalSet, ByVal service As String, ByVal budget As ATEvalBudget) As Boolean
    Dim token As Variant, found As Boolean
    budget.Step
    If Not mPorts Is Nothing Or mProtocols.Count > 0 Then Exit Function
    If Not mIP Is Nothing Then
        If Not ip.Overlaps(mIP, budget) Then Exit Function
    End If
    If mServices.Count > 0 Then
        For Each token In mServices
            budget.Step Len(service) + Len(CStr(token)) + 1
            If LCase$(Trim$(service)) = CStr(token) Then found = True
        Next token
        If Not found Then Exit Function
    End If
    SupersetBase = True
End Function
Friend Property Get HasHostname() As Boolean
    HasHostname = mHosts.Count > 0
End Property
Friend Function Specificity(ByVal host As String, ByVal budget As ATEvalBudget) As Long
    Dim token As Variant, matched As Boolean, score As Long
    host = LCase$(Trim$(host))
    If Right$(host, 1) = "." Then host = Left$(host, Len(host) - 1)
    For Each token In mHosts
        budget.Step Len(host) + Len(CStr(token)) + 1
        Select Case mMode
            Case 3: matched = host = CStr(token)
            Case 2: matched = Left$(host, Len(CStr(token))) = CStr(token)
            Case 1: matched = InStr(1, host, CStr(token), vbBinaryCompare) > 0
        End Select
        If matched Then
            score = mMode * 100000 + Len(CStr(token))
            If score > Specificity Then Specificity = score
        End If
    Next token
End Function
Friend Function Output(ByVal conflict As Boolean, ByVal budget As ATEvalBudget) As ATRuntimeRow
    Dim row As New ATRuntimeRow, id As Variant, value As String
    row.Initialize "RULE_MATCH", mRaw.RowId
    For Each id In mOutputs
        value = AT_PolicyCell(mRaw, CStr(id), budget)
        If CStr(id) = "SEVERITY" Then value = mSeverity
        row.AddField CStr(id), AT_RuntimeOK(AT_RuntimeScalar("TEXT", value)), mRaw.CellRef(CStr(id))
    Next id
    row.AddField "CONFLICT", AT_RuntimeOK(AT_RuntimeScalar("BOOL", conflict))
    Set Output = row.Snapshot
End Function
Friend Function ConflictValue(ByVal field As String, ByVal budget As ATEvalBudget) As String
    ConflictValue = AT_ExcelHeader(AT_PolicyCell(mRaw, field, budget))
End Function
