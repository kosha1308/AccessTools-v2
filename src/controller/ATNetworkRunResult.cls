VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkRunResult"
Option Explicit
Private mReady As Boolean, mState As ATEvalState, mCode As String, mPhase As String, mData As ATCheckReport
Private mPolicy As ATNetworkPolicyReport, mReportBook As Workbook, mCoordinatorBook As Workbook
Private mRuleFingerprint As String, mSupersetFingerprint As String, mRequestUnsaved As Boolean
Friend Sub Initialize(ByVal state As ATEvalState, ByVal code As String, ByVal phase As String, ByVal data As ATCheckReport, ByVal policy As ATNetworkPolicyReport, ByVal reportBook As Workbook, ByVal coordinatorBook As Workbook, ByVal ruleFingerprint As String, ByVal supersetFingerprint As String, ByVal requestUnsaved As Boolean)
    If mReady Then Err.Raise vbObjectError + 3180, , "RUN_RESULT_FROZEN"
    If state <> atEvalOK And state <> atEvalError Then Err.Raise vbObjectError + 3180, , "RUN_RESULT_INVALID"
    If state = atEvalOK Then
        If reportBook Is Nothing Or coordinatorBook Is Nothing Or data Is Nothing Then Err.Raise vbObjectError + 3180, , "RUN_RESULT_INVALID"
    End If
    Set mData = data: Set mPolicy = policy: Set mReportBook = reportBook: Set mCoordinatorBook = coordinatorBook
    mState = state: mCode = code: mPhase = phase: mRuleFingerprint = ruleFingerprint
    mSupersetFingerprint = supersetFingerprint: mRequestUnsaved = requestUnsaved: mReady = True
End Sub
Private Sub RequireReady()
    If Not mReady Then Err.Raise vbObjectError + 3180, , "RUN_RESULT_ACCESS"
End Sub
Public Property Get State() As ATEvalState
    RequireReady: State = mState
End Property
Public Property Get ErrorCode() As String
    RequireReady: ErrorCode = mCode
End Property
Public Property Get Phase() As String
    RequireReady: Phase = mPhase
End Property
Public Property Get ReportData() As ATCheckReport
    RequireReady: Set ReportData = mData
End Property
Public Property Get PolicyReport() As ATNetworkPolicyReport
    RequireReady: Set PolicyReport = mPolicy
End Property
Public Property Get ReportBook() As Workbook
    RequireReady: Set ReportBook = mReportBook
End Property
Public Property Get CoordinatorBook() As Workbook
    RequireReady: Set CoordinatorBook = mCoordinatorBook
End Property
Public Property Get RulebookFingerprint() As String
    RequireReady: RulebookFingerprint = mRuleFingerprint
End Property
Public Property Get SupersetFingerprint() As String
    RequireReady: SupersetFingerprint = mSupersetFingerprint
End Property
Public Property Get RequestUnsaved() As Boolean
    RequireReady: RequestUnsaved = mRequestUnsaved
End Property
Public Property Get IsComplete() As Boolean
    RequireReady: IsComplete = mState = atEvalOK And Not mReportBook Is Nothing And Not mCoordinatorBook Is Nothing
End Property
Public Property Get AllChecksPassed() As Boolean
    Dim outcome As ATCheckOutcome
    RequireReady
    If Not IsComplete Or mData Is Nothing Then Exit Property
    If mData.Checks.Count <> 16 Then Exit Property
    For Each outcome In mData.Checks
        If outcome.Status <> "PASS" Then Exit Property
    Next outcome
    AllChecksPassed = True
End Property
