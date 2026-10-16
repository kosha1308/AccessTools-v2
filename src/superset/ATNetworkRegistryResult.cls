VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkRegistryResult"
Option Explicit
Private mReady As Boolean, mState As ATEvalState, mEntries As Collection, mReport As ATCheckReport, mSource As ATEvalResult
Private mRequest As ATNetworkReadResult
Friend Sub Initialize(ByVal state As ATEvalState, ByVal entries As Collection, ByVal report As ATCheckReport, Optional ByVal sourceIssues As ATEvalResult = Nothing, Optional ByVal request As ATNetworkReadResult = Nothing)
    Dim item As ATNetworkRegistryEntry, emptyRows As New Collection, emptyValue As New ATRuntimeValue
    If mReady Then Err.Raise vbObjectError + 3010, , "REGISTRY_FROZEN"
    If state <> atEvalOK And state <> atEvalError Then Err.Raise vbObjectError + 3010, , "REGISTRY_INPUT_INVALID"
    If report Is Nothing Or entries Is Nothing Then Err.Raise vbObjectError + 3010, , "REGISTRY_INPUT_INVALID"
    If state = atEvalError And entries.Count > 0 Then Err.Raise vbObjectError + 3010, , "REGISTRY_INPUT_INVALID"
    Set mEntries = New Collection
    For Each item In entries: mEntries.Add item: Next item
    If sourceIssues Is Nothing Then
        emptyValue.InitializeRows "IP_SOURCE_ISSUE", emptyRows: Set mSource = AT_RuntimeOK(emptyValue)
    Else
        Set mSource = sourceIssues.Clone
    End If
    mState = state: Set mReport = report: Set mRequest = request: mReady = True
End Sub
Friend Sub ValidateRequest(ByVal request As ATNetworkReadResult)
    RequireReady
    If request Is Nothing Or mRequest Is Nothing Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_SOURCE"
    If Not request Is mRequest Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_SOURCE"
End Sub
Private Sub RequireReady()
    If Not mReady Then Err.Raise vbObjectError + 3010, , "REGISTRY_ACCESS"
End Sub
Public Property Get State() As ATEvalState
    RequireReady: State = mState
End Property
Public Property Get Entries() As Collection
    Dim r As New Collection, item As ATNetworkRegistryEntry
    RequireReady
    For Each item In mEntries: r.Add item: Next item
    Set Entries = r
End Property
Public Property Get Report() As ATCheckReport
    RequireReady: Set Report = mReport
End Property
Public Property Get SourceIssues() As ATEvalResult
    RequireReady: Set SourceIssues = mSource.Clone
End Property
