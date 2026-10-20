Attribute VB_Name = "modAtMultiUi"
Option Explicit
Private syncing As Boolean, fieldKeys As Collection, ruleKeys As Collection, filterKey As String
Public Sub AtUiMultiFill(ByVal form As Object, ByVal editor As Object)
    Dim multi As Boolean, name As Variant, field As CAtTemplateField, rule As CAtTemplateRule, i As Long, selected As Long
    multi = TypeName(editor) = "CAtTemplateEditor": syncing = True
    For Each name In Array("cmdAddField", "cmdAddRule", "cmdCopyRule", "cmdDeleteRule", "cmdSelectedPreview", "cboFieldFilter", "cboRuleField", "txtRuleName", "lblRuleName", "lblRuleField", "chkRuleEnabled")
        form.Controls(CStr(name)).Visible = multi
    Next name
    If Not multi Then syncing = False: Exit Sub
    Set fieldKeys = New Collection: Set ruleKeys = New Collection
    form.cboFieldFilter.Clear: form.cboFieldFilter.AddItem "Все правила шаблона": form.cboRuleField.Clear
    selected = 0
    For Each field In editor.Fields
        fieldKeys.Add field.Id
        form.cboFieldFilter.AddItem field.DisplayName & " · " & field.Sheet & "!" & field.Address
        form.cboRuleField.AddItem field.DisplayName & " · " & field.Sheet & "!" & field.Address
        If field.Id = filterKey Then selected = fieldKeys.Count
        If field.Id = editor.SelectedFieldId Then form.cboRuleField.ListIndex = fieldKeys.Count - 1
    Next field
    If selected = 0 Then filterKey = ""
    form.cboFieldFilter.ListIndex = selected
    form.lstNavigation.Clear: selected = -1
    For Each rule In editor.Rules
        If filterKey = "" Or rule.FieldId = filterKey Then
            ruleKeys.Add rule.Id: form.lstNavigation.AddItem rule.DisplayName & IIf(rule.Enabled, "", " (выключено)")
            If rule.Id = editor.SelectedRuleId Then selected = ruleKeys.Count - 1
        End If
    Next rule
    form.lstNavigation.ListIndex = selected
    form.txtRuleName.Text = editor.RuleName: form.chkRuleEnabled.Value = editor.RuleEnabled
    form.txtRuleName.Enabled = editor.SelectedRuleId <> "": form.chkRuleEnabled.Enabled = form.txtRuleName.Enabled
    form.cboRuleField.Enabled = form.txtRuleName.Enabled: form.cmdCopyRule.Enabled = form.txtRuleName.Enabled
    form.cmdDeleteRule.Enabled = form.txtRuleName.Enabled: form.cmdSelectedPreview.Enabled = form.txtRuleName.Enabled
    For Each name In Array("cboCondition", "txtExpected", "txtMessage", "cboSeverity", "cboColor", "chkNote")
        form.Controls(CStr(name)).Enabled = form.txtRuleName.Enabled
    Next name
    For Each name In Array("txtField", "txtAddress", "cmdPick", "cmdPickName", "cboSheet")
        form.Controls(CStr(name)).Enabled = editor.SelectedFieldId <> ""
    Next name
    form.cmdAddRule.Enabled = editor.SelectedFieldId <> ""
    form.lblNavHelp.Caption = "Поля общие для правил. Выберите поле для фильтра или откройте все правила."
    syncing = False
End Sub
Public Sub AtUiMultiFilter(ByVal form As Object)
    If syncing Or AtUiIsPainting() Then Exit Sub
    Dim editor As Object
    Set editor = AtUiMultiEditor(): If editor Is Nothing Then Exit Sub
    On Error GoTo Failed
    AtUiChanged form
    filterKey = ""
    If form.cboFieldFilter.ListIndex > 0 Then
        filterKey = fieldKeys(form.cboFieldFilter.ListIndex): editor.SelectField filterKey
    End If
    AtUiFill form: Exit Sub
Failed:
    syncing = False: form.lblStatus.Caption = Err.Description
End Sub
Public Sub AtUiMultiSelectRule(ByVal form As Object)
    If syncing Or AtUiIsPainting() Then Exit Sub
    Dim editor As Object, id As String
    Set editor = AtUiMultiEditor(): If editor Is Nothing Then Exit Sub
    If form.lstNavigation.ListIndex < 0 Then Exit Sub
    On Error GoTo Failed
    id = ruleKeys(form.lstNavigation.ListIndex + 1): AtUiChanged form: editor.SelectRule id
    AtUiFill form: Exit Sub
Failed:
    form.lblStatus.Caption = Err.Description
End Sub
Public Sub AtUiMultiBindField(ByVal form As Object)
    If syncing Or AtUiIsPainting() Then Exit Sub
    Dim editor As Object, id As String
    Set editor = AtUiMultiEditor(): If editor Is Nothing Then Exit Sub
    If form.cboRuleField.ListIndex < 0 Then Exit Sub
    On Error GoTo Failed
    id = fieldKeys(form.cboRuleField.ListIndex + 1): AtUiChanged form: editor.BindRuleField id
    filterKey = "": AtUiFill form: Exit Sub
Failed:
    form.lblStatus.Caption = Err.Description
End Sub
Public Sub AtUiMultiCommand(ByVal form As Object, ByVal command As String)
    If syncing Then Exit Sub
    Dim editor As Object, field As CAtTemplateField, rule As CAtTemplateRule, preview As CAtAdminPreview
    Set editor = AtUiMultiEditor(): If editor Is Nothing Then Exit Sub
    On Error GoTo Failed
    AtUiChanged form
    Select Case command
        Case "ADD_FIELD"
            Set field = editor.AddField("Новое поле " & editor.Fields.Count + 1, editor.SheetName, "A1"): filterKey = field.Id
        Case "ADD_RULE"
            If editor.SelectedFieldId = "" Then Err.Raise 5, , "Сначала добавьте или выберите поле."
            Set rule = editor.AddRule(editor.SelectedFieldId, "Новое правило " & editor.Rules.Count + 1)
        Case "COPY_RULE": Set rule = editor.CopyRule(editor.SelectedRuleId)
        Case "DELETE_RULE": editor.DeleteRule editor.SelectedRuleId: filterKey = ""
        Case "SELECTED"
            Set preview = editor.PreviewSelected(AtUiSelectedBook(form)): AtUiShowResult preview: Exit Sub
        Case Else: Err.Raise 5, , "Неизвестная команда редактора."
    End Select
    AtUiFill form: Exit Sub
Failed:
    form.lblStatus.Caption = Err.Description
End Sub
Public Sub AtUiMultiRefreshLabel(ByVal form As Object, ByVal editor As Object)
    If syncing Or ruleKeys Is Nothing Then Exit Sub
    Dim i As Long
    syncing = True
    For i = 1 To ruleKeys.Count
        If ruleKeys(i) = editor.SelectedRuleId Then form.lstNavigation.List(i - 1) = editor.RuleName & IIf(editor.RuleEnabled, "", " (выключено)")
    Next i
    syncing = False
End Sub
