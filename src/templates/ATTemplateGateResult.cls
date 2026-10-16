VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATTemplateGateResult"
Option Explicit
Private mAllowed As Boolean, mCode As String, mDetail As String, mFrozen As Boolean
Friend Sub Initialize(ByVal allowed As Boolean, ByVal code As String, ByVal detail As String)
    If mFrozen Then Err.Raise vbObjectError + 2830, , "TEMPLATE_RESULT_FROZEN"
    mAllowed = allowed: mCode = code: mDetail = detail: mFrozen = True
End Sub
Public Property Get Allowed() As Boolean
    Allowed = mAllowed
End Property
Public Property Get Code() As String
    Code = mCode
End Property
Public Property Get Detail() As String
    Detail = mDetail
End Property
Public Property Get Message() As String
    If mCode = "TEMPLATE_MISMATCH" Then
        Message = "Используйте актуальный шаблон заявки."
    ElseIf Not mAllowed Then
        Message = "Не удалось проверить шаблон. Обратитесь к администратору."
    End If
End Property
