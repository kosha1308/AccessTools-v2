VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATEvalResult"
Option Explicit
Private mState As ATEvalState
Private mValue As ATRuntimeValue
Private mDiagnostics As Collection, mTrace As Collection
Private mFrozen As Boolean

Public Sub Initialize(ByVal state As ATEvalState, ByVal value As ATRuntimeValue, ByVal diagnostics As Collection, ByVal trace As Collection)
    Dim entry As ATTraceEntry
    If mFrozen Then Err.Raise vbObjectError + 2500, , "RUNTIME_VALUE_FROZEN"
    If state < atEvalOK Or state > atEvalManualReview Then Err.Raise vbObjectError + 2501, , "RUNTIME_RESULT_STATE"
    If state = atEvalOK Then
        If value Is Nothing Then Err.Raise vbObjectError + 2502, , "RUNTIME_VALUE_ACCESS"
        Set mValue = value.Clone
    End If
    mState = state
    Set mDiagnostics = AT_CopyDiagnostics(diagnostics)
    Set mTrace = New Collection
    If Not trace Is Nothing Then
        For Each entry In trace
            mTrace.Add entry
        Next entry
    End If
    mFrozen = True
End Sub
Public Property Get State() As ATEvalState
    State = mState
End Property
Public Property Get Value() As ATRuntimeValue
    If Not mFrozen Or mState <> atEvalOK Then Err.Raise vbObjectError + 2502, , "RUNTIME_VALUE_ACCESS"
    Set Value = mValue.Clone
End Property
Public Property Get Diagnostics() As Collection
    Set Diagnostics = AT_CopyDiagnostics(mDiagnostics)
End Property
Public Property Get Trace() As Collection
    Dim result As New Collection, entry As ATTraceEntry
    If Not mTrace Is Nothing Then
        For Each entry In mTrace
            result.Add entry
        Next entry
    End If
    Set Trace = result
End Property
Public Function Clone() As ATEvalResult
    Dim result As New ATEvalResult
    result.Initialize mState, mValue, mDiagnostics, mTrace
    Set Clone = result
End Function
Public Sub ValidateBudget(ByVal budget As ATEvalBudget)
    If mState = atEvalOK Then mValue.ValidateBudget budget
End Sub
