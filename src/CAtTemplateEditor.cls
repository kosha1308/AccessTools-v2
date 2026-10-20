VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtTemplateEditor"
Option Explicit
Private mFields As Collection, mRules As Collection, packageKey As String, fieldKey As String, ruleKey As String
Private journal As CAtOwnershipJournal, journalBook As Object, store As CAtJournalStore
Private directory As String, journalRoot As String, sequence As Long
Public TemplateName As String
Private Sub Class_Initialize()
    Set mFields = New Collection: Set mRules = New Collection
    packageKey = AtTemplateId("P"): TemplateName = "Новый шаблон"
End Sub
Friend Sub ConfigureJournalRoot(ByVal root As String)
    journalRoot = root
End Sub
Friend Sub SetPackageId(ByVal id As String)
    packageKey = id
End Sub
Public Property Get PackageId() As String
    PackageId = packageKey
End Property
Public Property Get Fields() As Collection
    Set Fields = mFields
End Property
Public Property Get Rules() As Collection
    Set Rules = mRules
End Property
Public Property Get SelectedFieldId() As String
    SelectedFieldId = fieldKey
End Property
Public Property Get SelectedRuleId() As String
    SelectedRuleId = ruleKey
End Property
Public Function GetField(ByVal id As String) As CAtTemplateField
    Dim field As CAtTemplateField
    For Each field In mFields
        If field.Id = id Then Set GetField = field: Exit Function
    Next field
    Err.Raise 5, , "Поле не найдено."
End Function
Public Function GetRule(ByVal id As String) As CAtTemplateRule
    Dim rule As CAtTemplateRule
    For Each rule In mRules
        If rule.Id = id Then Set GetRule = rule: Exit Function
    Next rule
    Err.Raise 5, , "Правило не найдено."
End Function
Friend Function ImportField(ByVal id As String, ByVal name As String, ByVal sheet As String, ByVal address As String) As CAtTemplateField
    Dim field As New CAtTemplateField
    If mFields.Count >= 32 Then Err.Raise 5, , "Редактор поддерживает не более32полей."
    field.Initialize id: field.DisplayName = name: field.Sheet = sheet: field.Address = address
    mFields.Add field: fieldKey = id: Set ImportField = field
End Function
Public Function AddField(ByVal name As String, ByVal sheet As String, ByVal address As String) As CAtTemplateField
    Set AddField = ImportField(AtTemplateId("F"), name, sheet, address): ruleKey = ""
End Function
Friend Function ImportRule(ByVal id As String, ByVal fieldId As String, ByVal name As String) As CAtTemplateRule
    Dim rule As New CAtTemplateRule, field As CAtTemplateField
    If mRules.Count >= 32 Then Err.Raise 5, , "Редактор поддерживает не более32правил; бюджет Writer не расширяется."
    Set field = GetField(fieldId)
    rule.Initialize id: rule.FieldId = field.Id: rule.DisplayName = name: mRules.Add rule
    SelectRule rule.Id: Set ImportRule = rule
End Function
Public Function AddRule(ByVal fieldId As String, ByVal name As String) As CAtTemplateRule
    Set AddRule = ImportRule(AtTemplateId("R"), fieldId, name)
End Function
Public Sub SelectRule(ByVal id As String)
    Dim rule As CAtTemplateRule
    Set rule = GetRule(id): ruleKey = rule.Id: fieldKey = rule.FieldId
End Sub
Public Sub SelectField(ByVal id As String)
    Dim field As CAtTemplateField, rule As CAtTemplateRule
    Set field = GetField(id): fieldKey = id: ruleKey = ""
    For Each rule In mRules
        If rule.FieldId = id Then ruleKey = rule.Id: Exit For
    Next rule
End Sub
Public Sub BindRuleField(ByVal id As String)
    Dim field As CAtTemplateField
    Set field = GetField(id)
    If ruleKey = "" Then Err.Raise 5, , "Сначала выберите правило."
    GetRule(ruleKey).FieldId = field.Id: fieldKey = field.Id
End Sub
Public Function CopyRule(ByVal id As String) As CAtTemplateRule
    Dim original As CAtTemplateRule, copy As CAtTemplateRule
    Set original = GetRule(id): Set copy = AddRule(original.FieldId, original.DisplayName & " — копия")
    copy.Condition = original.Condition: copy.ExpectedText = original.ExpectedText: copy.Enabled = original.Enabled
    copy.Message = original.Message: copy.Severity = original.Severity: copy.Color = original.Color: copy.AddNote = original.AddNote: copy.Audience = original.Audience
    Set CopyRule = copy
End Function
Public Sub DeleteRule(ByVal id As String)
    Dim i As Long, removed As Boolean
    For i = mRules.Count To 1 Step -1
        If mRules(i).Id = id Then mRules.Remove i: removed = True
    Next i
    If Not removed Then Err.Raise 5, , "Правило не найдено."
    ruleKey = ""
    If mRules.Count > 0 Then SelectRule mRules(1).Id
End Sub
Public Function Model() As CAtPackage
    Dim p As CAtPackage, registry As New CAtOperationRegistry, issues As Collection
    If mFields.Count > 32 Or mRules.Count > 32 Then Err.Raise 5, , "Превышен бюджет редактора."
    Set p = AtCompileTemplate(Me): Set issues = AtValidate(p, registry.Catalog)
    If AtHasErrors(issues) Then Err.Raise 5, , "Исправьте конфигурацию перед запуском/сохранением:" & vbLf & AtUiDiagnostics(issues)
    Set Model = p
End Function
Public Sub SaveFile(ByVal path As String)
    Dim registry As New CAtOperationRegistry, issues As Collection
    If Not AtSaveNew(Model, path, registry.Catalog, issues) Then Err.Raise 5, , "Не удалось сохранить. Используйте новое имя файла."
End Sub
Public Sub LoadFile(ByVal path As String)
    Dim p As CAtPackage, registry As New CAtOperationRegistry, issues As Collection, candidate As CAtTemplateEditor
    Dim old As New CAtAdminEditor, field As CAtTemplateField, rule As CAtTemplateRule
    Set p = AtLoad(path, registry.Catalog, issues)
    If p Is Nothing Then Err.Raise 5, , "Файл отклонён без изменения текущей конфигурации."
    If AtHasErrors(issues) Then Err.Raise 5, , AtUiDiagnostics(issues)
    If p.SchemaVersion = "3.0.0" Then
        old.LoadFile path
        Set candidate = New CAtTemplateEditor: candidate.SetPackageId p.Id: candidate.TemplateName = old.TemplateName
        Set field = candidate.ImportField("FIELD_VALUE", old.FieldName, old.SheetName, old.Address)
        Set rule = candidate.ImportRule("RULE_EQUAL", field.Id, AtDefinitionById(p, "Rule", "RULE_EQUAL").DisplayName)
        rule.ExpectedText = old.ExpectedText: rule.Message = old.Message: rule.Severity = old.Severity: rule.Color = old.Color: rule.AddNote = old.AddNote
        rule.Audience = AtValue(AtDefinitionById(p, "Action", "ACTION_FAIL"), "Audience")
    Else
        Set candidate = AtReadTemplate(p)
    End If
    Set mFields = candidate.Fields: Set mRules = candidate.Rules: packageKey = candidate.PackageId: TemplateName = candidate.TemplateName
    fieldKey = candidate.SelectedFieldId: ruleKey = candidate.SelectedRuleId
End Sub
Public Function Preview(ByVal book As Object) As CAtAdminPreview
    Set Preview = EvaluateSet(book, "")
End Function
Public Function PreviewSelected(ByVal book As Object) As CAtAdminPreview
    If ruleKey = "" Then Err.Raise 5, , "Выберите правило."
    Set PreviewSelected = EvaluateSet(book, ruleKey)
    PreviewSelected.Output = "Предпросмотр выбранного правила. Отметки всех правил оставлены без изменений."
End Function
Private Function EvaluateSet(ByVal book As Object, ByVal onlyRule As String) As CAtAdminPreview
    Dim p As CAtPackage, resolver As CAtExcelResolver, context As CAtEvalContext, issues As Collection
    Dim result As New CAtAdminPreview, rule As CAtTemplateRule, field As CAtTemplateField, finding As CAtRuleFinding
    Dim incomplete As Boolean, errors As Long, count As Long, aggregate As CAtEvalResult
    Set p = Model: Set resolver = AtNewExcelResolver(p, "TEMPLATE", AtDictionary()): Set context = resolver.Resolve(book, issues)
    If context Is Nothing Then Err.Raise 5, , "Не удалось получить цельный снимок: " & AtUiDiagnostics(issues)
    Set result.Findings = New Collection
    For Each rule In mRules
        If (rule.Enabled And onlyRule = "") Or onlyRule = rule.Id Then
            count = count + 1: Set field = GetField(rule.FieldId): Set finding = New CAtRuleFinding
            finding.RuleId = rule.Id: finding.RuleName = rule.DisplayName: finding.FieldId = field.Id: finding.FieldName = field.DisplayName
            finding.Target = field.Sheet & "!" & field.Address: finding.Message = rule.Message: finding.Severity = rule.Severity
            Set finding.Outcome = AtEvaluateRule(p, rule.Id, context): Set finding.Actual = context.ReadField(field.Id)
            result.Findings.Add finding
            If finding.Outcome.Status = "CANNOT_EVALUATE" Then incomplete = True
            If finding.Outcome.Status = "FAIL" Then errors = errors + 1
            result.Diagnostics = result.Diagnostics & rule.DisplayName & ": " & finding.Outcome.Status & vbLf & AtUiDiagnostics(finding.Outcome.Result.Diagnostics)
        End If
    Next rule
    If incomplete Then
        Set aggregate = AtEvalValue("Bool", "ERROR", ""): result.Summary = "Проверка не завершена"
    Else
        Set aggregate = AtEvalValue("Bool", "OK", IIf(errors = 0, "true", "false"))
        result.Summary = IIf(errors = 0, "Пройдено", "Ошибок: " & errors)
        If count = 0 Then result.Summary = "Нет активных правил"
    End If
    Set result.Outcome = New CAtRuleOutcome: result.Outcome.Initialize "TEMPLATE_RUN", aggregate
    result.Target = "Активных правил: " & count: result.Remark = result.InitiatorText
    result.Actions = "Единое согласование отметок всех активных правил."
    result.Output = "Предпросмотр без изменений. Книга и журналы не изменялись."
    Set EvaluateSet = result
End Function
Public Property Get NeedsJournal() As Boolean
    NeedsJournal = journal Is Nothing
End Property
Public Function Apply(ByVal book As Object, ByVal ignoredDirectory As String) As CAtAdminPreview
    Dim result As CAtAdminPreview, finding As CAtRuleFinding, intents As New Collection, intent As CAtActionIntent
    Dim plan As CAtActionPlan, receipts As Collection, receipt As CAtActionReceipt, p As CAtPackage
    Set result = Preview(book)
    If result.Outcome.Status = "CANNOT_EVALUATE" Then
        result.Output = "Запись заблокирована: не все активные правила удалось вычислить. Прежние отметки сохранены."
        Set Apply = result: Exit Function
    End If
    If ignoredDirectory <> "" Then Err.Raise 5, , "Для целого шаблона используется автоматический журнал."
    Set p = Model
    For Each finding In result.Findings
        For Each intent In AtOutcomeIntents(p, finding.Outcome): intents.Add intent: Next intent
    Next finding
    If journal Is Nothing Then
        If intents.Count = 0 Then result.Output = "Новых ошибок нет; собственных отметок этой сессии нет.": Set Apply = result: Exit Function
        If store Is Nothing Then Set store = AtNewJournalStore(journalRoot)
        directory = store.NewDirectory(): Set journal = AtNewOwnershipJournal(book, directory): Set journalBook = book
    End If
    If Not journalBook Is book Then Err.Raise 5, , "Сначала очистите отметки другой связанной книги."
    sequence = sequence + 1: Set plan = AtPlanActions(book, intents, "TEMPLATE_" & sequence, journal)
    result.Output = AtUiStatus(plan.Status): result.Diagnostics = result.Diagnostics & vbLf & plan.Code & " " & plan.Message
    If plan.Status <> "READY" Then
        result.Output = result.Output & vbLf & "Правило и запись — разные результаты. Причина отказа: " & plan.Code & ". " & plan.Message
    End If
    If plan.Status = "READY" Then
        Set receipts = AtApplyPlan(plan, journal)
        For Each receipt In receipts
            result.Output = result.Output & vbLf & AtUiStatus(receipt.Status)
            result.Diagnostics = result.Diagnostics & receipt.Code & " " & receipt.Message & vbLf
        Next receipt
        If receipts.Count = 0 Then result.Output = "Актуальные отметки уже согласованы."
    End If
    result.Diagnostics = result.Diagnostics & ReleaseCompleted()
    Set Apply = result
End Function
Public Function Clear(ByVal book As Object) As String
    Dim receipt As CAtActionReceipt, receipts As Collection
    If journal Is Nothing Then Clear = "Отметок этой сессии нет.": Exit Function
    If Not journalBook Is book Then Err.Raise 5, , "Выберите связанную с отметками книгу."
    sequence = sequence + 1: Set receipts = AtClearOwned(book, journal, "TEMPLATE_CLEAR_" & sequence)
    For Each receipt In receipts: Clear = Clear & AtUiStatus(receipt.Status) & vbLf: Next receipt
    Clear = Clear & ReleaseCompleted()
End Function
Private Function ReleaseCompleted() As String
    Dim manifest As Object
    If journal Is Nothing Then Exit Function
    Set manifest = journal.CompletedManifest(): If manifest Is Nothing Then Exit Function
    If Not store.PruneCompleted(journal, directory) Then ReleaseCompleted = vbLf & "Завершённый журнал оставлен: уборка недоступна."
    Set journal = Nothing: Set journalBook = Nothing: directory = ""
End Function

Public Property Get FieldName() As String
    If fieldKey <> "" Then FieldName = GetField(fieldKey).DisplayName
End Property
Public Property Let FieldName(ByVal value As String)
    If fieldKey <> "" Then GetField(fieldKey).DisplayName = value
End Property

Public Property Get SheetName() As String
    If fieldKey <> "" Then SheetName = GetField(fieldKey).Sheet
End Property
Public Property Let SheetName(ByVal value As String)
    If fieldKey <> "" Then GetField(fieldKey).Sheet = value
End Property

Public Property Get Address() As String
    If fieldKey <> "" Then Address = GetField(fieldKey).Address
End Property
Public Property Let Address(ByVal value As String)
    If fieldKey <> "" Then GetField(fieldKey).Address = value
End Property

Public Property Get ExpectedText() As String
    If ruleKey <> "" Then ExpectedText = GetRule(ruleKey).ExpectedText
End Property
Public Property Let ExpectedText(ByVal value As String)
    If ruleKey <> "" Then GetRule(ruleKey).ExpectedText = value
End Property

Public Property Get Message() As String
    If ruleKey <> "" Then Message = GetRule(ruleKey).Message
End Property
Public Property Let Message(ByVal value As String)
    If ruleKey <> "" Then GetRule(ruleKey).Message = value
End Property

Public Property Get Severity() As String
    If ruleKey <> "" Then Severity = GetRule(ruleKey).Severity
End Property
Public Property Let Severity(ByVal value As String)
    If ruleKey <> "" Then GetRule(ruleKey).Severity = value
End Property

Public Property Get Color() As Long
    If ruleKey <> "" Then Color = GetRule(ruleKey).Color
End Property
Public Property Let Color(ByVal value As Long)
    If ruleKey <> "" Then GetRule(ruleKey).Color = value
End Property

Public Property Get AddNote() As Boolean
    If ruleKey <> "" Then AddNote = GetRule(ruleKey).AddNote
End Property
Public Property Let AddNote(ByVal value As Boolean)
    If ruleKey <> "" Then GetRule(ruleKey).AddNote = value
End Property

Public Property Get Condition() As String
    If ruleKey <> "" Then Condition = GetRule(ruleKey).Condition
End Property
Public Property Let Condition(ByVal value As String)
    If ruleKey <> "" Then GetRule(ruleKey).Condition = value
End Property

Public Property Get RuleName() As String
    If ruleKey <> "" Then RuleName = GetRule(ruleKey).DisplayName
End Property
Public Property Let RuleName(ByVal value As String)
    If ruleKey <> "" Then GetRule(ruleKey).DisplayName = value
End Property

Public Property Get RuleEnabled() As Boolean
    If ruleKey <> "" Then RuleEnabled = GetRule(ruleKey).Enabled
End Property
Public Property Let RuleEnabled(ByVal value As Boolean)
    If ruleKey <> "" Then GetRule(ruleKey).Enabled = value
End Property
