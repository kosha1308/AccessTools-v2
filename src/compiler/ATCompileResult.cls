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

