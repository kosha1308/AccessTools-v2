VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtEvalResult"
Option Explicit

Private mType As String, mState As String, mValue As String, mDependency As String
Private mDiagnostics As Collection, mRequirements As Collection, mProvenance As Collection, mTrace As Collection

Private Sub Class_Initialize()
    mType = "Text": mState = "ERROR"
    Set mDiagnostics = New Collection: Set mRequirements = New Collection
    Set mProvenance = New Collection: Set mTrace = New Collection
End Sub

Public Property Get ResultType() As String
    ResultType = mType
End Property
Public Property Get State() As String
    State = mState
End Property
Public Property Get Value() As String
    Value = mValue
End Property
Public Property Get HasValue() As Boolean
    HasValue = mState = "OK"
End Property
Public Property Get Dependency() As String
    Dependency = mDependency
End Property
Public Property Get IsComplete() As Boolean
    IsComplete = mState = "OK" And mRequirements.Count = 0
    Dim item As Variant
    For Each item In mDiagnostics
        If item("Severity") = "Error" Then IsComplete = False: Exit Property
    Next item
End Property
Public Property Get Diagnostics() As Collection
    Set Diagnostics = AtRecordsCopy(mDiagnostics)
End Property
Public Property Get Requirements() As Collection
    Set Requirements = AtRecordsCopy(mRequirements)
End Property
Public Property Get Provenance() As Collection
    Set Provenance = AtRecordsCopy(mProvenance)
End Property
Public Property Get Trace() As Collection
    Set Trace = AtRecordsCopy(mTrace)
End Property

Friend Sub SetData(ByVal typeName As String, ByVal state As String, ByVal value As String, Optional ByVal dependency As String = "")
    mType = typeName: mState = state: mValue = value: mDependency = dependency
End Sub

Friend Sub AddDiagnostic(ByVal code As String, ByVal objectId As String, ByVal nodeId As String, ByVal message As String)
    Dim d As Object: Set d = AtDictionary()
    d.Add "DiagnosticId", objectId & "/" & nodeId & "/" & code
    d.Add "Code", code: d.Add "ObjectId", objectId: d.Add "NodeId", nodeId
    d.Add "Message", message: d.Add "Severity", "Error"
    UniqueAdd mDiagnostics, d, "DiagnosticId"
End Sub

Friend Sub AddRequirement(ByVal objectId As String, ByVal nodeId As String, ByVal dependency As String, ByVal state As String)
    Dim d As Object: Set d = AtDictionary()
    d.Add "RequirementId", objectId & "/" & nodeId & "/" & dependency & "/" & state
    d.Add "OriginNodeId", nodeId: d.Add "ObjectId", objectId: d.Add "Dependency", dependency
    d.Add "RequiredBranch", nodeId: d.Add "State", state: d.Add "Disposition", "UNRESOLVED"
    Select Case state
        Case "EMPTY", "INVALID": d.Add "Classification", "Input"
        Case "NOT_APPLICABLE": d.Add "Classification", "Applicability"
        Case Else: d.Add "Classification", "Technical"
    End Select
    UniqueAdd mRequirements, d, "RequirementId"
End Sub

Friend Sub AddProvenance(ByVal kind As String, ByVal id As String)
    Dim d As Object: Set d = AtDictionary()
    d.Add "Kind", kind: d.Add "Id", id: d.Add "Key", kind & "/" & id
    UniqueAdd mProvenance, d, "Key"
End Sub

Friend Sub MergeEffects(ByVal source As CAtEvalResult)
    Dim item As Variant
    For Each item In source.Diagnostics: UniqueAdd mDiagnostics, item, "DiagnosticId": Next item
    For Each item In source.Requirements: UniqueAdd mRequirements, item, "RequirementId": Next item
    For Each item In source.Provenance: UniqueAdd mProvenance, item, "Key": Next item
End Sub

Friend Sub SetTrace(ByVal records As Collection)
    Set mTrace = AtRecordsCopy(records)
End Sub

Friend Sub BoundDiagnostics(ByVal maximum As Long, ByVal objectId As String, ByVal nodeId As String)
    If mDiagnostics.Count <= maximum Then Exit Sub
    Do While mDiagnostics.Count >= maximum: mDiagnostics.Remove mDiagnostics.Count: Loop
    AddDiagnostic "E_LIMIT_DIAGNOSTICS", objectId, nodeId, "Diagnostic budget exhausted; details truncated, requirements retained."
End Sub

Public Function Snapshot() As CAtEvalResult
    Dim r As New CAtEvalResult
    r.SetData mType, mState, mValue, mDependency
    r.MergeEffects Me: r.SetTrace mTrace
    Set Snapshot = r
End Function

Private Sub UniqueAdd(ByVal target As Collection, ByVal record As Object, ByVal key As String)
    Dim item As Variant
    For Each item In target
        If item(key) = record(key) Then Exit Sub
    Next item
    target.Add AtRecordCopy(record)
End Sub
