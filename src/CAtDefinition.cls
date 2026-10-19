VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtDefinition"
Option Explicit

' Kind selects a closed schema in modAtValidation, not an executable handler.
Public Kind As String
Public Id As String
Public Version As String
Public DisplayName As String
Public Values As Object
Public References As Collection
Public Dependencies As Collection
Public RootNodeId As String
Public ScopeTemplateId As String
Public Nodes As Collection

Private Sub Class_Initialize()
    Set Values = AtDictionary()
    Set References = New Collection
    Set Dependencies = New Collection
    Set Nodes = New Collection
    Version = "1.0.0"
End Sub
