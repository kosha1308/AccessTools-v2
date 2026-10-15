VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATCompileResult"
Option Explicit

Public Root As ATAstNode
Public Diagnostics As Collection
Public Dependencies As ATDependencyGraph
Private Sub Class_Initialize()
    Set Diagnostics = New Collection
    Set Dependencies = New ATDependencyGraph
End Sub

Public Property Get HasErrors() As Boolean
    Dim diagnostic As ATDiagnostic
    For Each diagnostic In Diagnostics
        If diagnostic.Severity = "ERROR" Then
            HasErrors = True
            Exit Property
        End If
    Next diagnostic
End Property

Public Property Get DSLVersion() As String
    DSLVersion = AT_DSL_VERSION
End Property

Public Property Get SchemaVersion() As String
    SchemaVersion = AT_SCHEMA_VERSION
End Property

