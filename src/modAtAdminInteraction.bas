Attribute VB_Name = "modAtAdminInteraction"
Option Explicit

' The UI command and its tests share this exact dispatch; no cached write plan.
Public Function AtUiRunMode(ByVal editor As Object, ByVal book As Object, ByVal mode As String) As CAtAdminPreview
    Select Case mode
        Case "TEST"
            Set AtUiRunMode = editor.Apply(book, "")
            If AtUiRunMode.Outcome.Status = "CANNOT_EVALUATE" And AtUiRunMode.Findings Is Nothing Then AtUiRunMode.Output = "Вычисление не завершено. Предыдущие отметки не изменялись."
        Case "PREVIEW": Set AtUiRunMode = editor.Preview(book)
        Case Else: Err.Raise 5, , "Неизвестная команда проверки."
    End Select
End Function

Public Function AtUiPickerBook(ByVal form As Object, ByVal mode As String) As Object
    Dim book As Object
    Set book = AtUiSelectedBook(form)
    If mode = "PICK" Then
        If form.cboSheet.ListIndex < 0 Then Err.Raise 5, , "Выберите лист."
        If book.Worksheets(form.cboSheet.Value).Visible <> -1 Then Err.Raise 5, , "Выберите видимый лист."
    ElseIf mode <> "PICK_NAME" Then
        Err.Raise 5, , "Неизвестный способ выбора ячейки."
    End If
    Set AtUiPickerBook = book
End Function

Public Function AtUiReadFieldName(ByVal book As Object, ByVal selected As Object) As String
    Dim area As Object, value As Variant
    If book Is Nothing Or selected Is Nothing Then Err.Raise 5, , "Выберите ячейку с названием поля."
    If TypeName(selected) <> "Range" Then Err.Raise 5, , "Требуется ячейка Excel."
    If Not selected.Parent.Parent Is book Then Err.Raise 5, , "Выберите название в указанной книге."
    If selected.Areas.Count <> 1 Then Err.Raise 5, , "Выберите одну ячейку или одно объединение."
    Set area = selected.Cells(1, 1).MergeArea
    If selected.Row < area.Row Or selected.Column < area.Column Or selected.Row + selected.Rows.Count > area.Row + area.Rows.Count Or selected.Column + selected.Columns.Count > area.Column + area.Columns.Count Then Err.Raise 5, , "Выберите одну ячейку или одно объединение."
    value = area.Cells(1, 1).Value2
    If IsError(value) Then Err.Raise 5, , "Ячейка с названием содержит ошибку Excel."
    If IsEmpty(value) Then Err.Raise 5, , "Ячейка с названием пуста."
    If Len(Trim$(CStr(value))) = 0 Then Err.Raise 5, , "Ячейка с названием пуста."
    AtUiReadFieldName = CStr(value)
End Function

Public Function AtUiAcceptFieldName(ByVal form As Object, ByVal book As Object, ByVal selected As Object) As Boolean
    If selected Is Nothing Then Exit Function
    ' Validate before replacing the existing text. The data locator is untouched.
    form.txtField.Text = AtUiReadFieldName(book, selected)
    AtUiChanged form
    AtUiAcceptFieldName = True
End Function
