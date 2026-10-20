Attribute VB_Name = "modAtAdminUi"
Option Explicit

Private editor As CAtAdminEditor, books As Collection, painting As Boolean

Public Function AtNewAdminEditor(Optional ByVal journalRoot As String = "") As CAtAdminEditor
    Dim instance As New CAtAdminEditor
    instance.ConfigureJournalRoot journalRoot: Set AtNewAdminEditor = instance
End Function

Public Function AtUiDiagnostics(ByVal issues As Collection) As String
    Dim issue As Variant
    For Each issue In issues
        AtUiDiagnostics = AtUiDiagnostics & issue("Code") & ": " & issue("Message") & vbLf
    Next issue
End Function

Public Function AtUiStatus(ByVal status As String) As String
    Select Case status
        Case "READY": AtUiStatus = "Готово к записи"
        Case "APPLIED": AtUiStatus = "Отметки нанесены"
        Case "ROLLED_BACK": AtUiStatus = "Собственные отметки сняты"
        Case "SKIPPED": AtUiStatus = "Без изменений"
        Case "CONFLICT": AtUiStatus = "Книга изменена: запись заблокирована"
        Case "RECOVERY_REQUIRED": AtUiStatus = "Требуется разбор состояния журнала"
        Case Else: AtUiStatus = "Запись не выполнена"
    End Select
End Function

Public Function AtUiNotesSupported(ByVal book As Object) As Boolean
#If Win64 Then
    AtUiNotesSupported = book.FileFormat = 56 And book.Application.Version = "16.0" And CStr(book.Application.Build) = "17932"
#End If
End Function

Public Sub AtAdminOpen()
    On Error GoTo Failed
    If editor Is Nothing Then Set editor = AtNewAdminEditor()
    Load AtAdminMain
    AtUiFill AtAdminMain
    AtUiRefreshBooks AtAdminMain
    AtUiFit AtAdminMain
    AtAdminMain.Show 0
    Exit Sub
Failed:
    MsgBox "Не удалось открыть конструктор: " & Err.Description, vbExclamation, "AccessTools"
End Sub

Public Sub AtUiFit(ByVal form As Object)
    AtUiLayout form, Application.UsableWidth * 0.94, Application.UsableHeight * 0.9
    form.StartUpPosition = 0
    form.Left = Application.Left + 12: form.Top = Application.Top + 12
End Sub

Public Sub AtUiFill(ByVal f As Object)
    painting = True
    f.cboSheet.Clear
    If editor Is Nothing Then Set editor = AtNewAdminEditor()
    If editor.SheetName <> "" Then f.cboSheet.AddItem editor.SheetName: f.cboSheet.ListIndex = 0
    f.txtTemplate.Text = editor.TemplateName: f.txtField.Text = editor.FieldName
    f.txtAddress.Text = editor.Address: f.txtExpected.Text = editor.ExpectedText: f.txtMessage.Text = editor.Message
    f.chkNote.Value = editor.AddNote
    f.cboCondition.Clear: f.cboCondition.AddItem "Текст равен": f.cboCondition.ListIndex = 0
    f.cboSeverity.Clear: f.cboSeverity.AddItem "Ошибка": f.cboSeverity.AddItem "Предупреждение": f.cboSeverity.AddItem "Информация"
    Select Case editor.Severity
        Case "ERROR": f.cboSeverity.ListIndex = 0
        Case "WARNING": f.cboSeverity.ListIndex = 1
        Case "INFO": f.cboSeverity.ListIndex = 2
    End Select
    f.cboColor.Clear
    Dim i As Long
    For i = 0 To 5: f.cboColor.AddItem AtUiColorName(i): Next i
    f.cboColor.ListIndex = -1
    For i = 0 To 5
        If AtUiColor(i) = editor.Color Then f.cboColor.ListIndex = i
    Next i
    If f.cboColor.ListIndex = -1 Then f.cboColor.AddItem "Цвет из файла": f.cboColor.ListIndex = 6
    f.cboMode.Clear: f.cboMode.AddItem "Простой режим": f.cboMode.AddItem "Расширенный режим": f.cboMode.ListIndex = 0
    f.lstNavigation.Clear: f.lstNavigation.AddItem "Правило: текст равен": f.lstNavigation.ListIndex = 0
    painting = False: AtUiChanged f
End Sub

Public Function AtUiColorName(ByVal index As Long) As String
    AtUiColorName = Choose(index + 1, "Красный", "Жёлтый", "Голубой", "Зелёный", "Сиреневый", "Серый")
End Function
Public Function AtUiColor(ByVal index As Long) As Long
    Select Case index
        Case 0: AtUiColor = 255
        Case 1: AtUiColor = 65535
        Case 2: AtUiColor = RGB(218, 232, 252)
        Case 3: AtUiColor = RGB(213, 232, 212)
        Case 4: AtUiColor = RGB(225, 213, 231)
        Case 5: AtUiColor = RGB(230, 230, 230)
        Case Else: Err.Raise 5, , "Выберите цвет из палитры."
    End Select
End Function

Public Sub AtUiRefreshBooks(ByVal f As Object)
    Dim book As Object
    painting = True: Set books = New Collection: f.cboBook.Clear
    For Each book In Application.Workbooks
        If Not book Is ThisWorkbook Then
            books.Add book: f.cboBook.AddItem book.Name
        End If
    Next book
    f.cboBook.ListIndex = -1: painting = False
    f.lblStatus.Caption = "Выберите открытую книгу для теста. Она не выбирается автоматически."
End Sub

Public Sub AtUiLoadFile(ByVal f As Object, ByVal path As String)
    editor.LoadFile path
    AtUiFill f: AtUiRefreshBooks f
    f.lblSaved.Caption = "Загружено. Выберите книгу; сохранение — в новый файл редакции."
End Sub

Public Function AtUiSelectedBook(ByVal f As Object) As Object
    If books Is Nothing Then Err.Raise 5, , "Обновите список открытых книг."
    If f.cboBook.ListIndex < 0 Then Err.Raise 5, , "Явно выберите открытую книгу."
    Dim selected As Object, book As Object
    Set selected = books(f.cboBook.ListIndex + 1)
    For Each book In Application.Workbooks
        If book Is selected Then Set AtUiSelectedBook = book: Exit Function
    Next book
    Err.Raise 5, , "Выбранная книга закрыта. Обновите список."
End Function

Public Sub AtUiBookChanged(ByVal f As Object)
    If painting Then Exit Sub
    On Error GoTo Failed
    Dim book As Object, sheet As Object, i As Long
    Set book = AtUiSelectedBook(f): painting = True: f.cboSheet.Clear
    For Each sheet In book.Worksheets
        f.cboSheet.AddItem sheet.Name
    Next sheet
    f.cboSheet.ListIndex = -1
    For i = 0 To f.cboSheet.ListCount - 1
        If f.cboSheet.List(i) = editor.SheetName Then f.cboSheet.ListIndex = i
    Next i
    painting = False: AtUiChanged f: Exit Sub
Failed:
    painting = False: f.lblStatus.Caption = Err.Description
End Sub

Public Sub AtUiChanged(ByVal f As Object)
    If painting Then Exit Sub
    If editor Is Nothing Then Exit Sub
    editor.TemplateName = f.txtTemplate.Text: editor.FieldName = f.txtField.Text
    editor.Address = f.txtAddress.Text: editor.ExpectedText = f.txtExpected.Text: editor.Message = f.txtMessage.Text
    If f.cboSheet.ListIndex >= 0 Then editor.SheetName = f.cboSheet.Value
    editor.AddNote = f.chkNote.Value
    f.lblNavTitle.Caption = editor.TemplateName
    f.lblHeading.Caption = "AccessTools — " & Left$(editor.TemplateName, 60)
    If f.cboSeverity.ListIndex >= 0 Then editor.Severity = Choose(f.cboSeverity.ListIndex + 1, "ERROR", "WARNING", "INFO")
    If f.cboColor.ListIndex >= 0 And f.cboColor.ListIndex < 6 Then editor.Color = AtUiColor(f.cboColor.ListIndex)
    f.lblColorSample.BackColor = editor.Color
    f.lblWhen.Caption = "По кнопке «Протестировать» для выбранной книги."
    f.lblWhat.Caption = editor.FieldName & " • " & editor.SheetName & "!" & editor.Address & vbLf & "Текст равен: «" & Left$(editor.ExpectedText, 100) & "»"
    f.lblExceptions.Caption = "Не настроены. Редактор исключений пока недоступен."
    f.lblAction.Caption = Left$(editor.Message, 150) & IIf(Len(editor.Message) > 150, "…", "") & vbLf & "Заливка" & IIf(editor.AddNote, " и внутреннее примечание", "; без примечания")
    f.lblDetails.Visible = f.cboMode.ListIndex = 1
    f.lblDetails.Caption = "Подробности: одно поле типа «Текст» → точное сравнение с учётом регистра → отметка при ошибке. Пустое/неверное/недоступное значение не считается успехом. Вложенные условия и исключения пока недоступны."
    f.lblSaved.Caption = "Изменения в памяти. Для новой редакции выберите новое имя файла."
End Sub

Public Sub AtUiCommand(ByVal f As Object, ByVal command As String)
    On Error GoTo Failed
    Dim path As Variant, book As Object, selected As Object, pickError As Long
    Dim candidate As CAtAdminPreview
    AtUiChanged f
    Select Case command
        Case "SAVE"
            path = Application.GetSaveAsFilename("Проверка.xml", "Конфигурация (*.xml),*.xml")
            If VarType(path) = vbBoolean Then Exit Sub
            editor.SaveFile CStr(path): f.lblSaved.Caption = "Сохранено: " & CStr(path)
        Case "LOAD"
            path = Application.GetOpenFilename("Конфигурация (*.xml),*.xml")
            If VarType(path) = vbBoolean Then Exit Sub
            AtUiLoadFile f, CStr(path)
        Case "PICK"
            Set book = AtUiSelectedBook(f)
            If f.cboSheet.ListIndex < 0 Then Err.Raise 5, , "Выберите лист."
            If book.Worksheets(f.cboSheet.Value).Visible <> -1 Then Err.Raise 5, , "Выберите видимый лист."
            f.Hide: book.Activate: book.Worksheets(f.cboSheet.Value).Activate
            On Error Resume Next
            Set selected = Application.InputBox("Выберите одну ячейку поля на выбранном листе", "Выбор поля", Type:=8)
            pickError = Err.Number: Err.Clear
            On Error GoTo Failed
            f.Show 0
            If selected Is Nothing Then
                If pickError <> 0 And pickError <> 424 Then Err.Raise pickError
                Exit Sub
            End If
            If Not selected.Parent.Parent Is book Then Err.Raise 5, , "Выберите ячейку в указанной книге."
            If selected.Parent.Name <> f.cboSheet.Value Or selected.Cells.CountLarge <> 1 Then Err.Raise 5, , "Выберите одну ячейку указанного листа."
            f.txtAddress.Text = selected.Address(False, False)
        Case "TEST", "APPLY"
            Set book = AtUiSelectedBook(f)
            If command = "TEST" Then
                Set candidate = editor.Preview(book)
            Else
                Set candidate = editor.Apply(book, "")
            End If
            AtUiShowResult candidate
        Case "CLEAR": f.lblStatus.Caption = editor.Clear(AtUiSelectedBook(f))
    End Select
    Exit Sub
Failed:
    If Not f.Visible Then f.Show 0
    f.lblStatus.Caption = Err.Description
End Sub

Public Sub AtUiShowResult(ByVal result As CAtAdminPreview)
    Load AtAdminResults
    AtUiFillResult AtAdminResults, result
    AtUiFit AtAdminResults
    AtAdminResults.Show 0
End Sub

Public Sub AtUiFillResult(ByVal form As Object, ByVal result As CAtAdminPreview)
    With form
        .lblResult.Caption = result.Summary
        .lblResult.ForeColor = AtUiOutcomeColor(result.Outcome.Status)
        .txtTarget.Text = result.Target: .txtActual.Text = result.ActualText
        .txtRemark.Text = result.InitiatorText
        .txtActions.Text = result.Actions: .txtOutput.Text = result.Output
        Select Case result.Outcome.Status
            Case "PASS": .txtDiagnostic.Text = "Условие выполнено; данные успешно прочитаны."
            Case "FAIL": .txtDiagnostic.Text = "Условие удалось вычислить, но фактический текст не совпадает с ожидаемым."
            Case Else: .txtDiagnostic.Text = "Результат не определён: проверьте доступность поля и корректность значения."
        End Select
        If result.Diagnostics <> "" Then .txtDiagnostic.Text = .txtDiagnostic.Text & vbLf & result.Diagnostics
        .txtDiagnostic.Visible = False: .lblDiagnostic.Visible = False
        .cmdDetails.Caption = "Технические подробности"
        .cmdCopy.Enabled = Len(result.InitiatorText) > 0
        .lblCopy.Caption = ""
    End With
End Sub

Public Function AtAdminUiSmoke() As String
    Load AtAdminMain
    Load AtAdminResults
    AtUiFill AtAdminMain
    AtAdminUiSmoke = "READY|" & AtAdminMain.Caption & "|" & AtAdminMain.Controls.Count & "|" & AtAdminResults.Controls.Count
End Function

Public Sub AtUiCopyRemark()
    On Error GoTo Failed
    If AtAdminResults.txtRemark.Text = "" Then Exit Sub
    Dim clipboard As New MSForms.DataObject
    ' Copy exactly the displayed initiator-only text, never a hidden candidate.
    clipboard.SetText AtAdminResults.txtRemark.Text: clipboard.PutInClipboard
    AtAdminResults.lblCopy.Caption = "Текст замечания скопирован без служебной информации."
    Exit Sub
Failed:
    AtAdminResults.lblCopy.Caption = "Не удалось скопировать. Выделите текст замечания и нажмите Ctrl+C."
End Sub
