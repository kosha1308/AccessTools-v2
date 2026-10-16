VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATInitiatorPolicy"
Option Explicit
Private mReady As Boolean, mIdentity As Collection, mIP As Collection, mAction As String
Friend Sub Initialize(ByVal identity As ATEvalResult, ByVal ipRules As ATEvalResult, ByVal noPolicyAction As String)
    Dim budget As New ATEvalBudget, limits As New ATEvalLimits, rows As ATRuntimeValue, row As ATRuntimeRow
    Dim selector As ATInitiatorSelector, i As Long, mode As String, approval As String, severity As String, comment As String
    Dim text As String, pair As Variant, scope As String, forbidden As Boolean, ordinary As Boolean
    If mReady Then Err.Raise vbObjectError + 2970, , "INITIATOR_FROZEN"
    If Len(noPolicyAction) > 16 Then Err.Raise vbObjectError + 2970, , "INITIATOR_NO_POLICY_ACTION"
    mAction = UCase$(Trim$(noPolicyAction))
    If mAction <> "ALLOW" And mAction <> "DENY" And mAction <> "WARNING" And mAction <> "SKIPPED" Then Err.Raise vbObjectError + 2970, , "INITIATOR_NO_POLICY_ACTION"
    budget.Initialize limits: Set mIdentity = New Collection: Set mIP = New Collection
    Set rows = AT_InitiatorRows(identity, "INITIATORS", budget)
    For i = 1 To rows.Count
        Set row = rows.Row(i)
        If AT_PolicyBoolean(AT_PolicyCell(row, "ACTIVE", budget)) Then
            Set selector = New ATInitiatorSelector: selector.Initialize AT_PolicyCell(row, "EMAIL_DOMAINS", budget), budget
            mode = AT_ExcelHeader(AT_PolicyCell(row, "SERVICE_MODE", budget)): ordinary = False
            Select Case mode
                Case "": ordinary = True
                Case "approval_only", "только согласование"
                Case Else: Err.Raise vbObjectError + 2970, , "INITIATOR_MODE_INVALID"
            End Select
            approval = AT_PolicyCell(row, "APPROVAL", budget): severity = AT_PolicyCell(row, "SEVERITY", budget)
            If AT_ExcelBlank(severity) Then severity = "WARNING" Else severity = AT_PolicySeverity(severity)
            comment = AT_PolicyCell(row, "COMMENT", budget)
            mIdentity.Add Array(selector, approval, severity, comment, ordinary)
        End If
    Next i
    Set rows = AT_InitiatorRows(ipRules, "INITIATOR_IP", budget)
    For i = 1 To rows.Count
        Set row = rows.Row(i): text = AT_PolicyCell(row, "ACTIVE", budget)
        If Len(AT_ExcelHeader(text)) = 0 Then text = "Да"
        If AT_PolicyBoolean(text) Then
            Set selector = New ATInitiatorSelector: selector.Initialize AT_PolicyCell(row, "EMAIL_DOMAIN", budget), budget
            scope = AT_InitiatorScope(AT_PolicyCell(row, "IP_SCOPE", budget))
            pair = AT_InitiatorIPPair(AT_PolicyCell(row, "ALLOWED_IP", budget), AT_PolicyCell(row, "DENIED_IP", budget), budget)
            forbidden = False
            Select Case AT_ExcelHeader(AT_PolicyCell(row, "DB_ACCOUNT", budget))
                Case "", "не ограничено"
                Case "запрещено": forbidden = True
                Case Else: Err.Raise vbObjectError + 2970, , "INITIATOR_DB_INVALID"
            End Select
            comment = AT_PolicyCell(row, "COMMENT", budget)
            mIP.Add Array(selector, scope, pair(0), pair(1), forbidden, comment)
        End If
    Next i
    mReady = True
End Sub

Public Function ForEmail(ByVal email As String, ByVal budget As ATEvalBudget) As ATInitiatorSelection
    Dim result As New ATInitiatorSelection, approvals As New Collection, matches As New Collection
    Dim entry As Variant, matched As Variant, region As Variant, regions As Object, scope As Variant
    Dim selector As ATInitiatorSelector, allowed As ATIntervalSet, denied As ATIntervalSet
    Dim hasAllowed As Boolean, admitted As Boolean, forbidden As Boolean, directional As Boolean
    Dim rank As Long, best As Long, valid As Boolean
    If Not mReady Then Err.Raise vbObjectError + 2970, , "INITIATOR_ACCESS"
    budget.CheckText Len(email): budget.Step Len(email) + 1
    email = AT_AsciiCase(Trim$(email), False): valid = AT_EmailSyntaxValid(email)
    For Each entry In mIdentity
        budget.Step: Set selector = entry(0)
        If CBool(entry(4)) Then hasAllowed = True
        rank = selector.Rank(email, budget)
        If rank > 0 Then
            If CBool(entry(4)) Then admitted = True
            If Not AT_ExcelBlank(entry(1)) Then approvals.Add Array(CStr(entry(1)), CStr(entry(2)), CStr(entry(3)))
        End If
    Next entry
    For Each entry In mIP
        budget.Step: Set selector = entry(0): rank = selector.Rank(email, budget)
        If rank > 0 Then matches.Add Array(entry, rank)
        If rank > best Then best = rank
    Next entry
    Set regions = CreateObject("Scripting.Dictionary")
    For Each scope In Array("ANY", "SOURCE", "DESTINATION")
        Set allowed = New ATIntervalSet: Set denied = New ATIntervalSet
        allowed.Initialize 4294967295#: denied.Initialize 4294967295#
        regions.Add CStr(scope), Array(allowed, denied, False)
    Next scope
    For Each matched In matches
        If CLng(matched(1)) = best Then
            entry = matched(0)
            If CBool(entry(4)) Then forbidden = True
            If CStr(entry(1)) <> "ANY" Then directional = True
            For Each scope In Array("ANY", "SOURCE", "DESTINATION")
                If CStr(entry(1)) = "ANY" Or CStr(entry(1)) = CStr(scope) Then
                    budget.Step: region = regions(CStr(scope))
                    Set allowed = region(0): Set denied = region(1)
                    AT_InitiatorUnion allowed, entry(2), budget: AT_InitiatorUnion denied, entry(3), budget
                    region(2) = True: regions(CStr(scope)) = region
                End If
            Next scope
        End If
    Next matched
    For Each scope In Array("ANY", "SOURCE", "DESTINATION")
        region = regions(CStr(scope)): Set allowed = region(0): Set denied = region(1)
        allowed.Freeze: denied.Freeze
    Next scope
    result.Initialize email, valid, hasAllowed, admitted, approvals, regions, forbidden, directional, mAction
    Set ForEmail = result
End Function
