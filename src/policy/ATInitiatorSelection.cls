VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATInitiatorSelection"
Option Explicit
Private mReady As Boolean, mEmail As String, mValid As Boolean, mHasAllowed As Boolean, mAdmitted As Boolean
Private mApprovals As Collection, mRegions As Object, mForbidden As Boolean, mDirectional As Boolean, mAction As String
Friend Sub Initialize(ByVal email As String, ByVal valid As Boolean, ByVal hasAllowed As Boolean, ByVal admitted As Boolean, ByVal approvals As Collection, ByVal regions As Object, ByVal forbidden As Boolean, ByVal directional As Boolean, ByVal action As String)
    Dim item As Variant
    If mReady Then Err.Raise vbObjectError + 2970, , "INITIATOR_FROZEN"
    mEmail = email: mValid = valid: mHasAllowed = hasAllowed: mAdmitted = admitted
    Set mApprovals = New Collection
    For Each item In approvals: mApprovals.Add Array(CStr(item(0)), CStr(item(1)), CStr(item(2))): Next item
    Set mRegions = regions: mForbidden = forbidden: mDirectional = directional: mAction = action: mReady = True
End Sub
Private Sub RequireReady()
    If Not mReady Then Err.Raise vbObjectError + 2970, , "INITIATOR_ACCESS"
End Sub
Public Property Get AccountForbidden() As Boolean
    RequireReady: AccountForbidden = mForbidden
End Property
Public Property Get HasDirectionalRules() As Boolean
    RequireReady: HasDirectionalRules = mDirectional
End Property
Public Property Get NoPolicyAction() As String
    RequireReady: NoPolicyAction = mAction
End Property
Public Function HasRegionPolicy(ByVal scope As String) As Boolean
    Dim region As Variant
    RequireReady
    If Not mRegions.Exists(scope) Then Err.Raise vbObjectError + 2970, , "INITIATOR_SCOPE_INVALID"
    region = mRegions(scope): HasRegionPolicy = CBool(region(2))
End Function
Public Function IdentityReport(ByVal ref As String) As ATCheckReport
    Dim issues As New Collection, checks As New Collection, report As New ATCheckReport, outcome As ATCheckOutcome
    Dim status As String, entry As Variant, message As String, reason As String
    RequireReady
    status = "PASS"
    If Not mValid Then
        status = "SKIPPED": reason = "Не указана корректная почта инициатора."
    ElseIf Not mHasAllowed Then
        status = "SKIPPED": reason = "Нет активных правил допуска инициатора."
    ElseIf Not mAdmitted Then
        status = "FAIL"
        AT_InitiatorIssue issues, "NET.INITIATOR_ALLOWED", "INITIATOR_NOT_ALLOWED", "ERROR", "EMAIL", ref, "Почта или домен инициатора отсутствует среди разрешённых."
    End If
    Set outcome = New ATCheckOutcome: outcome.Initialize "NET.INITIATOR_ALLOWED", status, reason: checks.Add outcome
    status = "PASS": reason = ""
    If Not mValid Then status = "SKIPPED": reason = "Не указана корректная почта инициатора."
    For Each entry In mApprovals
        message = "Для инициатора «" & mEmail & "» требуется согласование: " & CStr(entry(0)) & "."
        If Len(CStr(entry(2))) > 0 Then message = message & vbLf & CStr(entry(2))
        AT_InitiatorIssue issues, "NET.INITIATOR_APPROVAL_RULES", "INITIATOR_APPROVAL_REQUIRED", CStr(entry(1)), "APPROVAL", ref, message
        If CStr(entry(1)) = "ERROR" Then status = "FAIL"
    Next entry
    Set outcome = New ATCheckOutcome: outcome.Initialize "NET.INITIATOR_APPROVAL_RULES", status, reason: checks.Add outcome
    report.Initialize atEvalOK, "NETWORK_INITIATOR_IDENTITY", issues, checks: Set IdentityReport = report
End Function
Public Function Evaluate(ByVal expression As ATRuntimeValue, ByVal scope As String, ByVal budget As ATEvalBudget) As ATEvalResult
    Dim region As Variant, allowed As ATIntervalSet, denied As ATIntervalSet, answer As Boolean
    RequireReady: budget.Step
    If Not mValid Then
        Set Evaluate = AT_RuntimeFailure(atEvalSkipped, "INITIATOR_EMAIL_UNAVAILABLE", Nothing): Exit Function
    End If
    If expression Is Nothing Then GoTo Invalid
    expression.ValidateBudget budget
    If expression.DataType <> "IP_EXPRESSION" Or expression.ValueState <> atValuePresent Then GoTo Invalid
    If Not mRegions.Exists(scope) Then Err.Raise vbObjectError + 2970, , "INITIATOR_SCOPE_INVALID"
    region = mRegions(scope)
    If Not CBool(region(2)) Then
        Select Case mAction
            Case "ALLOW": answer = True
            Case "DENY": answer = False
            Case "WARNING": Set Evaluate = AT_RuntimeFailure(atEvalManualReview, "INITIATOR_IP_POLICY_MISSING", Nothing): Exit Function
            Case "SKIPPED": Set Evaluate = AT_RuntimeFailure(atEvalSkipped, "INITIATOR_IP_POLICY_MISSING", Nothing): Exit Function
        End Select
    Else
        Set allowed = region(0): Set denied = region(1)
        answer = expression.Domain.IsSubsetOf(allowed, budget)
        If expression.Domain.Overlaps(denied, budget) Then answer = False
    End If
    Set Evaluate = AT_RuntimeOK(AT_RuntimeScalar("BOOL", answer))
    Exit Function
Invalid:
    Set Evaluate = AT_RuntimeFailure(atEvalManualReview, "INITIATOR_IP_UNCHECKED", Nothing)
End Function
