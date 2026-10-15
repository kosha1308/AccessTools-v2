VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATParseResult"
Option Explicit

Public Root As ATAstNode
Public Diagnostics As Collection
Private Sub Class_Initialize()
    Set Diagnostics = New Collection
End Sub

