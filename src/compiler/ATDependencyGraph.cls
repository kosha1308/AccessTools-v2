VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATDependencyGraph"
Option Explicit

Private mDependencies As Collection
Private mEdges As Collection

Private Sub Class_Initialize()
    Set mDependencies = New Collection
    Set mEdges = New Collection
End Sub

Public Sub Add(ByVal dependency As String, Optional ByVal usedBy As String = "RULE")
    If Not Contains(dependency) Then mDependencies.Add dependency, dependency
    On Error Resume Next
    mEdges.Add usedBy & " -> " & dependency, usedBy & "|" & dependency
    On Error GoTo 0
End Sub

Public Function Contains(ByVal dependency As String) As Boolean
    Dim value As String
    On Error Resume Next
    value = mDependencies(dependency)
    Contains = Err.Number = 0
    On Error GoTo 0
End Function

Public Property Get Items() As Collection
    Set Items = mDependencies
End Property

Public Property Get UsedBy() As Collection
    Set UsedBy = mEdges
End Property

