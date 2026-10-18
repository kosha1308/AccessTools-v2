Attribute VB_Name = "modATFeedbackMenu"
Option Explicit
Public Function AT_BuildFeedbackMenu(ByVal errorCount As Long, ByVal warning As Boolean) As Object
    Dim bar As Object, button As Object, fs As Object, name As String, qualifier As String
    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject"): name = "AccessToolsv2_" & fs.GetTempName
    Set bar = Application.CommandBars.Add(name, 5, False, True)
    qualifier = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"
    Set button = bar.Controls.Add(1): button.Caption = "Ошибок пользователя: " & CStr(errorCount): button.Enabled = False
    Set button = bar.Controls.Add(1): button.Caption = "Скопировать все ошибки пользователя": button.Tag = "COPY"
    button.OnAction = qualifier & "AT_ResultCopy": button.Enabled = errorCount > 0
    Set button = bar.Controls.Add(1): button.Caption = "Открыть отчёт проверок": button.OnAction = qualifier & "AT_ResultReport"
    Set button = bar.Controls.Add(1): button.Caption = "Открыть сводку координатора": button.OnAction = qualifier & "AT_ResultCoordinator"
    Set button = bar.Controls.Add(1): button.Caption = "Вернуться к заявке с примечаниями": button.OnAction = qualifier & "AT_ResultRequest"
    Set button = bar.Controls.Add(1): button.Caption = "Сохранить результаты": button.OnAction = qualifier & "AT_ResultSave"
    If warning Then
        Set button = bar.Controls.Add(1): button.Caption = "Есть ограничения: подробности показаны в сообщении": button.Enabled = False
    End If
    Set AT_BuildFeedbackMenu = bar: Exit Function
Failed:
    On Error Resume Next
    If Not bar Is Nothing Then bar.Delete
    On Error GoTo 0
    Err.Raise vbObjectError + 3230, , "FEEDBACK_MENU_FAILED"
End Function
