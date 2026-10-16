VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATCheckReport"
Option Explicit
Private mState As ATEvalState, mScope As String, mFrozen As Boolean
Private mIssues As Collection, mChecks As Collection, mErrors As Long, mWarnings As Long, mSkipped As Long
Friend Sub Initialize(ByVal state As ATEvalState, ByVal scopeId As String, ByVal issues As Collection, ByVal checks As Collection)
    Dim issue As ATCheckIssue, outcome As ATCheckOutcome, copiedIssues As New Collection, copiedChecks As New Collection
    Dim errors As Long, warnings As Long, skipped As Long, scope As String
    If mFrozen Then Err.Raise vbObjectError + 2940, , "CHECK_VALUE_FROZEN"
    If state <> atEvalOK And state <> atEvalError Then BadReport
    If issues Is Nothing Or checks Is Nothing Then BadReport
    If issues.Count > 1000 Or checks.Count > 1000 Then Err.Raise vbObjectError + 2940, , "CHECK_REPORT_LIMIT"
    scope = AT_CanonicalId(scopeId)
    For Each issue In issues
        If issue Is Nothing Then BadReport
        Select Case issue.Severity
            Case "ERROR": errors = errors + 1
            Case "WARNING": warnings = warnings + 1
            Case "INFO"
            Case Else: BadReport
        End Select
        copiedIssues.Add issue
    Next issue
    For Each outcome In checks
        If outcome Is Nothing Then BadReport
        Select Case outcome.Status
            Case "SKIPPED": skipped = skipped + 1
            Case "PASS", "FAIL", "BLOCKED"
            Case Else: BadReport
        End Select
        copiedChecks.Add outcome
    Next outcome
    If state = atEvalError And errors = 0 Then BadReport
    mState = state: mScope = scope: mErrors = errors: mWarnings = warnings: mSkipped = skipped
    Set mIssues = copiedIssues: Set mChecks = copiedChecks: mFrozen = True
End Sub
Private Sub BadReport()
    Err.Raise vbObjectError + 2940, , "CHECK_REPORT_INVALID"
End Sub
Private Sub RequireInitialized()
    If Not mFrozen Then Err.Raise vbObjectError + 2940, , "CHECK_REPORT_REQUIRED"
End Sub
Public Property Get State() As ATEvalState
    RequireInitialized: State = mState
End Property
Public Property Get ScopeId() As String
    RequireInitialized: ScopeId = mScope
End Property
Public Property Get ErrorCount() As Long
    RequireInitialized: ErrorCount = mErrors
End Property
Public Property Get WarningCount() As Long
    RequireInitialized: WarningCount = mWarnings
End Property
Public Property Get SkippedCount() As Long
    RequireInitialized: SkippedCount = mSkipped
End Property
Public Property Get HasErrors() As Boolean
    RequireInitialized: HasErrors = (mState = atEvalError Or mErrors > 0)
End Property
Public Property Get Summary() As String
    RequireInitialized
    Summary = "Проверки полей заявки. Ошибок: " & CStr(mErrors) & "; предупреждений: " & CStr(mWarnings) & "; пропущено: " & CStr(mSkipped) & "."
End Property
Public Property Get Issues() As Collection
    Dim copied As New Collection, issue As ATCheckIssue
    RequireInitialized
    For Each issue In mIssues
        copied.Add issue
    Next issue
    Set Issues = copied
End Property
Public Property Get Checks() As Collection
    Dim copied As New Collection, outcome As ATCheckOutcome
    RequireInitialized
    For Each outcome In mChecks
        copied.Add outcome
    Next outcome
    Set Checks = copied
End Property
