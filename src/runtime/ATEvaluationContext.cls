VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATEvaluationContext"
Option Explicit
Private mRunId As String, mRunDate As String, mTimezone As String, mLocale As String, mComparison As String
Private mFields As Collection, mIds As Collection, mRefs As Collection
Private mFrozen As Boolean, mInitialized As Boolean

Private Sub Class_Initialize()
    Set mFields = New Collection
    Set mIds = New Collection
    Set mRefs = New Collection
End Sub
Public Sub Initialize(ByVal runId As String, ByVal runDate As String, ByVal timezoneId As String, ByVal localeId As String, ByVal comparisonMode As String)
    If mInitialized Then Err.Raise vbObjectError + 2500, , "RUNTIME_VALUE_FROZEN"
    mRunId = runId: mRunDate = runDate: mTimezone = timezoneId: mLocale = localeId: mComparison = comparisonMode
    mInitialized = True
End Sub
Public Sub AddField(ByVal id As String, ByVal result As ATEvalResult, Optional ByVal cellRef As String = "")
    If mFrozen Then Err.Raise vbObjectError + 2500, , "RUNTIME_VALUE_FROZEN"
    id = AT_CanonicalId(id)
    If HasField(id) Then Err.Raise vbObjectError + 2503, , "RUNTIME_DUPLICATE_FIELD"
    mFields.Add result.Clone, id
    mIds.Add id
    mRefs.Add cellRef, id
End Sub
Private Function HasField(ByVal id As String) As Boolean
    Dim result As ATEvalResult
    On Error Resume Next
    Set result = mFields(id)
    HasField = Not result Is Nothing
    On Error GoTo 0
End Function
Public Function ResolveField(ByVal id As String, Optional ByVal budget As ATEvalBudget = Nothing) As ATEvalResult
    If Not HasField(id) Then
        Set ResolveField = AT_RuntimeFailure(atEvalError, "RUNTIME_FIELD_MISSING", Nothing)
    Else
        If Not budget Is Nothing Then mFields(id).ValidateBudget budget
        Set ResolveField = mFields(id).Clone
    End If
End Function
Public Function CellRef(ByVal id As String) As String
    If HasField(id) Then CellRef = mRefs(id)
End Function
Public Function Snapshot() As ATEvaluationContext
    Dim result As New ATEvaluationContext, id As Variant
    result.Initialize mRunId, mRunDate, mTimezone, mLocale, mComparison
    For Each id In mIds
        result.RetainField CStr(id), mFields(CStr(id)), CStr(mRefs(CStr(id)))
    Next id
    result.Freeze
    Set Snapshot = result
End Function
Friend Sub RetainField(ByVal id As String, ByVal result As ATEvalResult, Optional ByVal cellRef As String = "")
    If mFrozen Then Err.Raise vbObjectError + 2500, , "RUNTIME_VALUE_FROZEN"
    mFields.Add result, id
    mIds.Add id
    mRefs.Add cellRef, id
End Sub
Public Sub Freeze()
    mFrozen = True
End Sub
Public Sub Validate()
    Dim normalized As ATEvalResult
    If Not mInitialized Or Len(mRunId) = 0 Or Len(mRunDate) = 0 Or Len(mTimezone) = 0 Or Len(mLocale) = 0 Or Len(mComparison) = 0 Then Err.Raise vbObjectError + 2504, , "CONFIG_CONTEXT_REQUIRED"
    If mLocale <> "INVARIANT" Or (mComparison <> "BINARY" And mComparison <> "ASCII_TRIM_CI") Then Err.Raise vbObjectError + 2504, , "CONFIG_UNSUPPORTED"
    Set normalized = AT_NormalizeScalar("DATE", mRunDate, Nothing)
    If normalized.State <> atEvalOK Then Err.Raise vbObjectError + 2504, , "CONFIG_CONTEXT_REQUIRED"
    If normalized.Value.ValueState <> atValuePresent Then Err.Raise vbObjectError + 2504, , "CONFIG_CONTEXT_REQUIRED"
End Sub
Public Property Get RunId() As String
    RunId = mRunId
End Property
Public Property Get RunDate() As String
    RunDate = mRunDate
End Property
Public Property Get TimezoneId() As String
    TimezoneId = mTimezone
End Property
Public Property Get LocaleId() As String
    LocaleId = mLocale
End Property
Public Property Get ComparisonMode() As String
    ComparisonMode = mComparison
End Property
