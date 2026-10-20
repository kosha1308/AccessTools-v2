VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtTemplateField"
Option Explicit
Private identity As String
Public DisplayName As String, Sheet As String, Address As String
Friend Sub Initialize(ByVal id As String)
    If identity <> "" Then Err.Raise 5, , "Field identity is immutable"
    identity = id
End Sub
Public Property Get Id() As String
    Id = identity
End Property
