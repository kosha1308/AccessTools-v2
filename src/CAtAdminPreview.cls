VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtAdminPreview"
Option Explicit
Public Outcome As CAtRuleOutcome, Actual As CAtEvalResult
Public Summary As String, Target As String, Remark As String, Actions As String, Diagnostics As String, Output As String

Public Property Get ActualText() As String
    If Actual Is Nothing Then ActualText = "Значение не прочитано": Exit Property
    Select Case Actual.State
        Case "OK": ActualText = Actual.Value
        Case "EMPTY": ActualText = "Пустая ячейка"
        Case "MISSING": ActualText = "Поле не найдено"
        Case "INVALID": ActualText = "Значение не соответствует типу поля"
        Case Else: ActualText = "Значение недоступно"
    End Select
End Property

Public Property Get InitiatorText() As String
    If Not Outcome Is Nothing Then
        If Outcome.Status = "FAIL" Then InitiatorText = Remark
    End If
End Property
