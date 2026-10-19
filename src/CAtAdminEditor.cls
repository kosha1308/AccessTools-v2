VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtAdminEditor"
Option Explicit

Public TemplateName As String, FieldName As String, SheetName As String, Address As String
Public ExpectedText As String, Message As String, Severity As String, Color As Long, AddNote As Boolean
Private identity As CAtPackage, audience As String
Private journal As CAtOwnershipJournal, journalBook As Object, sequence As Long

Private Sub Class_Initialize()
    TemplateName = "Новый шаблон": FieldName = "Поле заявки": Address = "A1"
    Message = "Проверьте значение поля.": Severity = "ERROR": Color = 255: audience = "FirstLine"
End Sub

Public Function Model() As CAtPackage
    Dim p As CAtPackage, issues As Collection, registry As New CAtOperationRegistry
    Set p = AtAdminDefinition(SheetName, Address, FieldName, ExpectedText, Message, Severity, Color, AddNote)
    p.DisplayName = TemplateName
    If Not identity Is Nothing Then p.Id = identity.Id: p.Version = identity.Version
    AtDefinitionById(p, "Action", "ACTION_FAIL").Values("Audience") = audience
    Set issues = AtValidate(p, registry.Catalog)
    If AtHasErrors(issues) Then Err.Raise 5, , "Проверьте название шаблона и поля, лист, адрес, сообщение и настройки отметки. Сообщение должно быть коротким (до 490 символов)."
    Set Model = p
End Function

Public Sub LoadFile(ByVal path As String)
    Dim p As CAtPackage, candidate As New CAtAdminEditor, issues As Collection, registry As New CAtOperationRegistry
    Dim locator As CAtDefinition, action As CAtDefinition, expression As CAtDefinition, n As CAtNode
    Set p = AtLoad(path, registry.Catalog, issues)
    If p Is Nothing Then Err.Raise 5, , "Не удалось загрузить файл настроек проверки."
    If AtHasErrors(issues) Then Err.Raise 5, , "Настройки проверки содержат ошибки. Файл отклонён."
    If p.SchemaVersion <> "3.0.0" Or p.Definitions.Count <> 6 Then Err.Raise 5, , "Редактор поддерживает одно условие «Текст равен». Файл не изменён."
    Set locator = AtDefinitionById(p, "Locator", "LOCATION")
    Set action = AtDefinitionById(p, "Action", "ACTION_FAIL")
    Set expression = AtDefinitionById(p, "Expression", "EXPRESSION")
    candidate.TemplateName = p.DisplayName
    candidate.FieldName = AtDefinitionById(p, "Field", "FIELD_VALUE").DisplayName
    candidate.SheetName = AtValue(locator, "Sheet"): candidate.Address = AtValue(locator, "Address")
    For Each n In expression.Nodes
        If n.Id = "EXPECTED" And n.Values.Exists("Value") Then candidate.ExpectedText = CStr(n.Values("Value"))
    Next n
    candidate.Message = AtValue(action, "Message"): candidate.Severity = UCase$(AtValue(action, "Trigger"))
    candidate.Color = CLng(AtValue(action, "FillColor")): candidate.AddNote = AtValue(action, "AddNote") = "true"
    candidate.SetOrigin p, AtValue(action, "Audience")
    If AtToXml(candidate.Model) <> AtToXml(p) Then Err.Raise 5, , "В файле есть недоступные этому редактору настройки. Файл не изменён."
    TemplateName = candidate.TemplateName: FieldName = candidate.FieldName: SheetName = candidate.SheetName: Address = candidate.Address
    ExpectedText = candidate.ExpectedText: Message = candidate.Message: Severity = candidate.Severity
    Color = candidate.Color: AddNote = candidate.AddNote: audience = AtValue(action, "Audience"): Set identity = p
End Sub

Friend Sub SetOrigin(ByVal p As CAtPackage, ByVal role As String)
    Set identity = p: audience = role
End Sub

Public Sub SaveFile(ByVal path As String)
    Dim issues As Collection, registry As New CAtOperationRegistry
    If Not AtSaveNew(Model, path, registry.Catalog, issues) Then Err.Raise 5, , "Не удалось сохранить. Выберите доступную папку и новое имя файла редакции."
End Sub

Public Function Preview(ByVal book As Object) As CAtAdminPreview
    Dim p As CAtPackage, resolver As CAtExcelResolver, context As CAtEvalContext, issues As Collection
    Dim result As New CAtAdminPreview
    Set p = Model
    Set resolver = AtNewExcelResolver(p, "TEMPLATE", AtDictionary())
    Set context = resolver.Resolve(book, issues)
    If context Is Nothing Then Err.Raise 5, , "Не удалось прочитать выбранную книгу." & vbLf & AtUiDiagnostics(issues)
    Set result.Outcome = AtEvaluateRule(p, "RULE_EQUAL", context)
    Set result.Actual = context.ReadField("FIELD_VALUE")
    result.Target = SheetName & "!" & Address
    result.Remark = Message
    Select Case result.Outcome.Status
        Case "PASS": result.Summary = "Пройдено": result.Actions = "Снять предыдущие собственные отметки, если они есть в этой сессии."
        Case "FAIL": result.Summary = "Ошибка": result.Actions = "Заливка выбранным цветом."
            If AddNote Then result.Actions = result.Actions & " Внутреннее примечание для 1ЛП / координатора."
            If AddNote And Not AtUiNotesSupported(book) Then result.Actions = result.Actions & vbLf & "Примечания в этой книге или версии Excel не поддерживаются. Нанесение будет отклонено целиком."
        Case Else: result.Summary = "Не удалось проверить": result.Actions = "Не изменять отметки: вычисление не завершено."
    End Select
    result.Diagnostics = AtUiDiagnostics(result.Outcome.Result.Diagnostics)
    result.Output = "Только предварительное тестирование. Книга не изменялась."
    Set Preview = result
End Function

Public Property Get NeedsJournal() As Boolean
    NeedsJournal = journal Is Nothing
End Property

' Always reread/re-evaluate; a displayed preview is never a reusable write plan.
Public Function Apply(ByVal book As Object, ByVal newJournalDirectory As String) As CAtAdminPreview
    Dim result As CAtAdminPreview, plan As CAtActionPlan, receipts As Collection, receipt As CAtActionReceipt
    Set result = Preview(book)
    If result.Outcome.Status = "CANNOT_EVALUATE" Then Set Apply = result: Exit Function
    If journal Is Nothing Then
        If result.Outcome.Status = "PASS" Then result.Output = "Нанесённых отметок этой сессии нет.": Set Apply = result: Exit Function
        Set journal = AtNewOwnershipJournal(book, newJournalDirectory): Set journalBook = book
    End If
    If Not journalBook Is book Then Err.Raise 5, , "Журнал связан с другой книгой. Сначала очистите её отметки."
    sequence = sequence + 1
    Set plan = AtPlanActions(book, AtOutcomeIntents(Model, result.Outcome), "VISUAL_" & sequence, journal)
    result.Output = AtUiStatus(plan.Status)
    result.Diagnostics = result.Diagnostics & vbLf & plan.Status & " " & plan.Code & " " & plan.Message
    If plan.Status = "READY" Then
        Set receipts = AtApplyPlan(plan, journal)
        For Each receipt In receipts
            result.Output = result.Output & vbLf & AtUiStatus(receipt.Status)
            result.Diagnostics = result.Diagnostics & vbLf & receipt.Status & " " & receipt.Code & " " & receipt.Message
        Next receipt
    End If
    Set Apply = result
End Function

Public Function Clear(ByVal book As Object) As String
    Dim receipts As Collection, receipt As CAtActionReceipt, complete As Boolean
    If journal Is Nothing Then Clear = "Нанесённых отметок этой сессии нет.": Exit Function
    If Not journalBook Is book Then Err.Raise 5, , "Выберите книгу, в которой были нанесены отметки этой сессии."
    sequence = sequence + 1: complete = True
    Set receipts = AtClearOwned(journalBook, journal, "VISUAL_CLEAR_" & sequence)
    For Each receipt In receipts
        Clear = Clear & AtUiStatus(receipt.Status) & vbLf
        If receipt.Status <> "ROLLED_BACK" And receipt.Status <> "SKIPPED" Then complete = False
    Next receipt
    If complete And Not journal.RecoveryRequired Then Set journal = Nothing: Set journalBook = Nothing
End Function
