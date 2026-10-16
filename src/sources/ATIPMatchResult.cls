VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATIPMatchResult"
Option Explicit
Private mReady As Boolean, mState As ATEvalState, mMatches As ATRuntimeValue, mComplete As Boolean, mDiagnostics As Collection
Friend Sub Initialize(ByVal state As ATEvalState, ByVal matches As ATRuntimeValue, ByVal complete As Boolean, ByVal diagnostics As Collection)
    If mReady Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_FROZEN"
    If state = atEvalOK Then
        If matches Is Nothing Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_INPUT_INVALID"
        Set mMatches = matches.Clone: mComplete = complete
    End If
    mState = state: Set mDiagnostics = AT_CopyDiagnostics(diagnostics): mReady = True
End Sub
Public Property Get State() As ATEvalState
    If Not mReady Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_ACCESS"
    State = mState
End Property
Public Property Get Matches() As ATRuntimeValue
    If Not mReady Or mState <> atEvalOK Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_ACCESS"
    Set Matches = mMatches.Clone
End Property
Public Property Get SourceComplete() As Boolean
    If Not mReady Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_ACCESS"
    SourceComplete = mComplete
End Property
Public Property Get Diagnostics() As Collection
    If Not mReady Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_ACCESS"
    Set Diagnostics = AT_CopyDiagnostics(mDiagnostics)
End Property
