VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkSupersetResult"
Option Explicit
Private mReady As Boolean, mState As ATEvalState, mRegistry As ATNetworkRegistryResult
Private mAnnotations As ATRuntimeValue, mReport As ATCheckReport
Friend Sub Initialize(ByVal state As ATEvalState, ByVal registry As ATNetworkRegistryResult, ByVal annotations As Collection, ByVal report As ATCheckReport)
    If mReady Then Err.Raise vbObjectError + 3140, , "SUPERSET_RULES_FROZEN"
    If state <> atEvalOK And state <> atEvalError Then Err.Raise vbObjectError + 3140, , "SUPERSET_RULES_INPUT_INVALID"
    If annotations Is Nothing Or report Is Nothing Then Err.Raise vbObjectError + 3140, , "SUPERSET_RULES_INPUT_INVALID"
    If state = atEvalError And annotations.Count > 0 Then Err.Raise vbObjectError + 3140, , "SUPERSET_RULES_INPUT_INVALID"
    Set mAnnotations = New ATRuntimeValue: mAnnotations.InitializeRows "SUPERSET_ANNOTATION", annotations
    Set mRegistry = registry: Set mReport = report: mState = state: mReady = True
End Sub
Public Property Get State() As ATEvalState
    RequireReady: State = mState
End Property
Public Property Get Annotations() As ATRuntimeValue
    RequireReady: Set Annotations = mAnnotations.Clone
End Property
Public Property Get Report() As ATCheckReport
    RequireReady: Set Report = mReport
End Property
Friend Sub ValidateRegistry(ByVal registry As ATNetworkRegistryResult)
    RequireReady
    If registry Is Nothing Or mRegistry Is Nothing Then Err.Raise vbObjectError + 3140, , "SUPERSET_RULES_BINDING"
    If Not registry Is mRegistry Then Err.Raise vbObjectError + 3140, , "SUPERSET_RULES_BINDING"
End Sub
Private Sub RequireReady()
    If Not mReady Then Err.Raise vbObjectError + 3140, , "SUPERSET_RULES_ACCESS"
End Sub
