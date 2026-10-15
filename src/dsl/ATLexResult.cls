VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATLexResult"
Option Explicit

Public Tokens As Collection
Public Diagnostics As Collection

Private Sub Class_Initialize()
    Set Tokens = New Collection
    Set Diagnostics = New Collection
End Sub

