VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtPackage"
Option Explicit

Public Id As String
Public Version As String
Public DisplayName As String
Public SchemaVersion As String
Public State As String
Public PublishedBy As String
Public PublishedAt As String
Public Definitions As Collection

Private Sub Class_Initialize()
    SchemaVersion = "1.0.0"
    Version = "1.0.0"
    State = "Draft"
    Set Definitions = New Collection
End Sub
