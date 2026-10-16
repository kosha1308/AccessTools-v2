VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkReadResult"
Option Explicit
Private mState As ATEvalState, mFrozen As Boolean
Private mContext As ATEvaluationContext, mDiagnostics As Collection
Friend Sub Initialize(ByVal state As ATEvalState, ByVal context As ATEvaluationContext, ByVal diagnostics As Collection)
    If mFrozen Then Err.Raise vbObjectError + 2920, , "NETWORK_RESULT_FROZEN"
    If state <> atEvalOK And state <> atEvalError Then Err.Raise vbObjectError + 2920, , "NETWORK_RESULT_STATE"
    If state = atEvalOK Then
        If context Is Nothing Then Err.Raise vbObjectError + 2920, , "NETWORK_CONTEXT_REQUIRED"
        Set mContext = context.Snapshot
    End If
    mState = state: Set mDiagnostics = AT_CopyDiagnostics(diagnostics): mFrozen = True
End Sub
Public Property Get State() As ATEvalState
    If Not mFrozen Then Err.Raise vbObjectError + 2920, , "NETWORK_RESULT_REQUIRED"
    State = mState
End Property
Public Property Get Context() As ATEvaluationContext
    If Not mFrozen Or mState <> atEvalOK Then Err.Raise vbObjectError + 2920, , "NETWORK_CONTEXT_ACCESS"
    Set Context = mContext.Snapshot
End Property
Public Property Get Diagnostics() As Collection
    Set Diagnostics = AT_CopyDiagnostics(mDiagnostics)
End Property
