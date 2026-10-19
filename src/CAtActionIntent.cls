VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtActionIntent"
Option Explicit

Public ActionInstanceKey As String
Public Sheet As String
Public Address As String
Public Severity As String
Public Message As String
Public FillColor As Long
Public Highlight As Boolean
Public AddNote As Boolean

Private Sub Class_Initialize()
    Severity = "ERROR": FillColor = 255: Highlight = True: AddNote = True
End Sub
