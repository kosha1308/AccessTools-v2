VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkPolicyResult"
Option Explicit
Private mReady As Boolean, mState As ATEvalState, mPolicy As ATNetworkPolicy, mDiagnostics As Collection
Friend Sub Initialize(ByVal state As ATEvalState, ByVal policy As ATNetworkPolicy, ByVal diagnostics As Collection)
    If mReady Then Err.Raise vbObjectError + 2960, , "POLICY_FROZEN"
    If state <> atEvalOK And state <> atEvalError Then Err.Raise vbObjectError + 2960, , "POLICY_INPUT_INVALID"
    If state = atEvalOK Then
        If policy Is Nothing Then Err.Raise vbObjectError + 2960, , "POLICY_INPUT_INVALID"
        Set mPolicy = policy
    End If
    mState = state: Set mDiagnostics = AT_CopyDiagnostics(diagnostics): mReady = True
End Sub
Public Property Get State() As ATEvalState
    If Not mReady Then Err.Raise vbObjectError + 2960, , "POLICY_ACCESS"
    State = mState
End Property
Public Property Get Policy() As ATNetworkPolicy
    If Not mReady Or mState <> atEvalOK Then Err.Raise vbObjectError + 2960, , "POLICY_ACCESS"
    Set Policy = mPolicy
End Property
Public Property Get Diagnostics() As Collection
    Set Diagnostics = AT_CopyDiagnostics(mDiagnostics)
End Property
