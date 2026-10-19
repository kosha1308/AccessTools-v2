VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtNode"
Option Explicit

Public Id As String
Public Kind As String
Public ResultType As String
Public Values As Object
Public Parameters As Object
Public Bindings As Collection
Public Children As Collection

Private Sub Class_Initialize()
    Set Values = AtDictionary(): Set Parameters = AtDictionary()
    Set Bindings = New Collection: Set Children = New Collection
End Sub
