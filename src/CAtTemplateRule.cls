VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtTemplateRule"
Option Explicit
Private identity As String
Public DisplayName As String, FieldId As String, Condition As String, ExpectedText As String
Public Audience As String, Message As String, Severity As String, Color As Long, AddNote As Boolean, Enabled As Boolean
Friend Sub Initialize(ByVal id As String)
    If identity <> "" Then Err.Raise 5, , "Rule identity is immutable"
    identity = id
End Sub
Public Property Get Id() As String
    Id = identity
End Property
Private Sub Class_Initialize()
    Audience = "FirstLine": Condition = "EQUAL": Message = "Проверьте значение поля.": Severity = "ERROR": Color = 255: Enabled = True
End Sub
