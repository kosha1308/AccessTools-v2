VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATCoordinatorNotesResult"
Option Explicit
Private mReady As Boolean, mState As ATEvalState, mRegistry As ATNetworkRegistryResult
Private mRows As ATRuntimeValue, mGlobal As Long, mIncomplete As Boolean, mCode As String
Friend Sub Initialize(ByVal state As ATEvalState, ByVal registry As ATNetworkRegistryResult, ByVal rows As Collection, ByVal globalCount As Long, ByVal incomplete As Boolean, Optional ByVal code As String = "")
    If mReady Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_FROZEN"
    If state <> atEvalOK And state <> atEvalError Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_INPUT"
    If rows Is Nothing Or globalCount < 0 Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_INPUT"
    If state = atEvalError And rows.Count > 0 Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_INPUT"
    Set mRows = New ATRuntimeValue: mRows.InitializeRows "COORDINATOR_NOTES", rows
    Set mRegistry = registry: mState = state: mGlobal = globalCount: mIncomplete = incomplete: mCode = code: mReady = True
End Sub
Private Sub RequireReady()
    If Not mReady Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_ACCESS"
End Sub
Friend Sub ValidateRegistry(ByVal registry As ATNetworkRegistryResult)
    RequireReady
    If registry Is Nothing Or mRegistry Is Nothing Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_BINDING"
    If Not registry Is mRegistry Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_BINDING"
End Sub
Public Property Get State() As ATEvalState
    RequireReady: State = mState
End Property
Public Property Get Rows() As ATRuntimeValue
    RequireReady: Set Rows = mRows.Clone
End Property
Public Property Get GlobalCount() As Long
    RequireReady: GlobalCount = mGlobal
End Property
Public Property Get Incomplete() As Boolean
    RequireReady: Incomplete = mIncomplete
End Property
Public Property Get ErrorCode() As String
    RequireReady: ErrorCode = mCode
End Property
