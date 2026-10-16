VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATInitiatorPolicyResult"
Option Explicit
Private mReady As Boolean, mState As ATEvalState, mPolicy As ATInitiatorPolicy, mDiagnostics As Collection
Friend Sub Initialize(ByVal state As ATEvalState, ByVal policy As ATInitiatorPolicy, ByVal diagnostics As Collection)
    If mReady Then Err.Raise vbObjectError + 2970, , "INITIATOR_FROZEN"
    If state <> atEvalOK And state <> atEvalError Then Err.Raise vbObjectError + 2970, , "INITIATOR_INPUT_INVALID"
    If state = atEvalOK Then
        If policy Is Nothing Then Err.Raise vbObjectError + 2970, , "INITIATOR_INPUT_INVALID"
        Set mPolicy = policy
    End If
    mState = state: Set mDiagnostics = AT_CopyDiagnostics(diagnostics): mReady = True
End Sub
Private Sub RequireReady()
    If Not mReady Then Err.Raise vbObjectError + 2970, , "INITIATOR_ACCESS"
End Sub
Public Property Get State() As ATEvalState
    RequireReady: State = mState
End Property
Public Property Get Policy() As ATInitiatorPolicy
    RequireReady
    If mState <> atEvalOK Then Err.Raise vbObjectError + 2970, , "INITIATOR_ACCESS"
    Set Policy = mPolicy
End Property
Public Property Get Diagnostics() As Collection
    RequireReady: Set Diagnostics = AT_CopyDiagnostics(mDiagnostics)
End Property
