VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATAstNode"
Option Explicit

Public Kind As String
Public Text As String
Public DataType As String
Public Start As Long
Public Length As Long
Public Children As Collection
Public Expansion As ATAstNode
Public Depth As Long

Private Sub Class_Initialize()
    Set Children = New Collection
End Sub

