VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATIPSourceResult"
Option Explicit
Private mReady As Boolean, mState As ATEvalState, mIndex As ATIPSourceIndex, mDiagnostics As Collection
Friend Sub Initialize(ByVal state As ATEvalState, ByVal index As ATIPSourceIndex, ByVal diagnostics As Collection)
    If mReady Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_FROZEN"
    If state <> atEvalOK And state <> atEvalError Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_INPUT_INVALID"
    If state = atEvalOK Then
        If index Is Nothing Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_INPUT_INVALID"
        Set mIndex = index
    End If
    mState = state: Set mDiagnostics = AT_CopyDiagnostics(diagnostics): mReady = True
End Sub
Private Sub RequireReady()
    If Not mReady Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_ACCESS"
End Sub
Public Property Get State() As ATEvalState
    RequireReady: State = mState
End Property
Public Property Get Index() As ATIPSourceIndex
    RequireReady
    If mState <> atEvalOK Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_ACCESS"
    Set Index = mIndex
End Property
Public Property Get Diagnostics() As Collection
    RequireReady: Set Diagnostics = AT_CopyDiagnostics(mDiagnostics)
End Property
