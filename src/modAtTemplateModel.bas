Attribute VB_Name = "modAtTemplateModel"
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef value As Any) As Long
#End If
Public Function AtTemplateId(ByVal prefix As String) As String
    Dim bytes(0 To 15) As Byte, i As Long
    If CoCreateGuid(bytes(0)) <> 0 Then Err.Raise 5, , "Cannot generate identity"
    AtTemplateId = prefix & "_"
    For i = 0 To 15: AtTemplateId = AtTemplateId & Right$("0" & Hex$(bytes(i)), 2): Next i
End Function
Public Function AtNewTemplateEditor(Optional ByVal journalRoot As String = "", Optional ByVal seed As Boolean = True) As CAtTemplateEditor
    Dim editor As New CAtTemplateEditor, field As CAtTemplateField, rule As CAtTemplateRule
    editor.ConfigureJournalRoot journalRoot
    If seed Then
        Set field = editor.AddField("Поле заявки", "", "A1")
        Set rule = editor.AddRule(field.Id, "Текст равен")
    End If
    Set AtNewTemplateEditor = editor
End Function
Public Function AtCompileTemplate(ByVal editor As CAtTemplateEditor) As CAtPackage
    Dim p As CAtPackage, d As CAtDefinition, field As CAtTemplateField, rule As CAtTemplateRule
    Set p = AtNewPackage(editor.PackageId, editor.TemplateName): p.SchemaVersion = "4.0.0"
    Set d = AtNewDefinition("Template", "TEMPLATE", editor.TemplateName): p.Definitions.Add d
    For Each field In editor.Fields
        Set d = AtNewDefinition("Locator", "L_" & field.Id, "Расположение поля")
        d.Values.Add "Strategy", "A1": d.Values.Add "Sheet", field.Sheet: d.Values.Add "Address", field.Address
        AtAddReference d, "OwnerTemplate", "Template", "TEMPLATE", "1.0.0": p.Definitions.Add d
        Set d = AtNewDefinition("Field", field.Id, field.DisplayName): d.Values.Add "DataType", "Text"
        AtAddReference d, "OwnerTemplate", "Template", "TEMPLATE", "1.0.0"
        AtAddReference d, "Locator", "Locator", "L_" & field.Id, "1.0.0": p.Definitions.Add d
    Next field
    For Each rule In editor.Rules: CompileRule p, rule: Next rule
    Set AtCompileTemplate = p
End Function
Private Sub CompileRule(ByVal p As CAtPackage, ByVal rule As CAtTemplateRule)
    Dim d As CAtDefinition, n As CAtNode, trigger As String
    Select Case rule.Severity
        Case "ERROR": trigger = "Error"
        Case "WARNING": trigger = "Warning"
        Case "INFO": trigger = "Info"
        Case Else: Err.Raise 5, , "Выберите уровень замечания."
    End Select
    Set d = AtNewDefinition("Expression", "E_" & rule.Id, rule.DisplayName)
    d.Values.Add "ReturnType", "Bool": d.Values.Add "Deferred", "false": d.ScopeTemplateId = "TEMPLATE": d.RootNodeId = "CHECK"
    Set n = AtNewNode("FIELD", "FieldRef", "Text"): n.Values.Add "FieldId", rule.FieldId: n.Values.Add "FieldVersion", "1.0.0": d.Nodes.Add n
    If rule.Condition = "REQUIRED" Then
        Set n = AtNewNode("CHECK", "Call", "Bool"): n.Values.Add "OperationId", "OP_TEXT_REQUIRED": n.Values.Add "OperationVersion", "1.0.0"
        n.Children.Add "FIELD": d.Nodes.Add n
    ElseIf rule.Condition = "EQUAL" Or rule.Condition = "NOT_EQUAL" Then
        Set n = AtNewNode("EXPECTED", "TypedLiteral", "Text"): n.Values.Add "State", "OK": n.Values.Add "Value", rule.ExpectedText: d.Nodes.Add n
        Set n = AtNewNode("CHECK", "Call", "Bool"): n.Values.Add "OperationId", "OP_TEXT_EQUAL": n.Values.Add "OperationVersion", "1.0.0"
        n.Parameters.Add "MODE", AtTypedParameter("Text", "binary"): n.Children.Add "FIELD": n.Children.Add "EXPECTED": d.Nodes.Add n
        If rule.Condition = "NOT_EQUAL" Then
            Set n = AtNewNode("NOT", "Not", "Bool"): n.Children.Add "CHECK": d.Nodes.Add n: d.RootNodeId = "NOT"
        End If
    Else
        Err.Raise 5, , "Условие не поддерживается редактором."
    End If
    p.Definitions.Add d
    Set d = AtNewDefinition("Rule", rule.Id, rule.DisplayName): d.Values.Add "Enabled", IIf(rule.Enabled, "true", "false")
    AtAddReference d, "OwnerTemplate", "Template", "TEMPLATE", "1.0.0"
    AtAddReference d, "Expression", "Expression", "E_" & rule.Id, "1.0.0", "Bool"
    AtAddReference d, "Field", "Field", rule.FieldId, "1.0.0", "Text": p.Definitions.Add d
    Set d = AtNewDefinition("Action", "A_" & rule.Id, rule.DisplayName)
    d.Values.Add "Trigger", trigger: d.Values.Add "Audience", rule.Audience
    d.Values.Add "Channel", IIf(rule.AddNote, "Comment", "Results"): d.Values.Add "OnOutcome", "FAIL": d.Values.Add "Message", rule.Message
    d.Values.Add "FillColor", CStr(rule.Color): d.Values.Add "AddNote", IIf(rule.AddNote, "true", "false")
    AtAddReference d, "OwnerTemplate", "Template", "TEMPLATE", "1.0.0"
    AtAddReference d, "Rule", "Rule", rule.Id, "1.0.0": AtAddReference d, "Field", "Field", rule.FieldId, "1.0.0", "Text": p.Definitions.Add d
End Sub

Public Function AtReadTemplate(ByVal p As CAtPackage) As CAtTemplateEditor
    Dim editor As New CAtTemplateEditor, d As CAtDefinition, locator As CAtDefinition, action As CAtDefinition, expression As CAtDefinition
    Dim field As CAtTemplateField, rule As CAtTemplateRule, n As CAtNode, condition As String, expected As String
    If p.SchemaVersion <> "4.0.0" Or p.State <> "Draft" Then Err.Raise 5, , "Требуется черновик схемы4. Старый файл не изменён."
    editor.SetPackageId p.Id: editor.TemplateName = p.DisplayName
    For Each d In p.Definitions
        If d.Kind = "Field" Then
            Set locator = AtDefinitionById(p, "Locator", AtReferenceId(d, "Locator"))
            Set field = editor.ImportField(d.Id, d.DisplayName, AtValue(locator, "Sheet"), AtValue(locator, "Address"))
        End If
    Next d
    For Each d In p.Definitions
        If d.Kind = "Rule" Then
            Set expression = AtDefinitionById(p, "Expression", AtReferenceId(d, "Expression"))
            condition = "EQUAL": expected = ""
            If expression.RootNodeId = "NOT" Then condition = "NOT_EQUAL"
            For Each n In expression.Nodes
                If n.Kind = "Call" Then
                    If n.Values("OperationId") = "OP_TEXT_REQUIRED" Then condition = "REQUIRED"
                End If
                If n.Id = "EXPECTED" Then expected = CStr(n.Values("Value"))
            Next n
            Set rule = editor.ImportRule(d.Id, AtReferenceId(d, "Field"), d.DisplayName)
            rule.Condition = condition: rule.ExpectedText = expected: rule.Enabled = AtValue(d, "Enabled") = "true"
            Set action = AtDefinitionById(p, "Action", "A_" & rule.Id)
            rule.Message = AtValue(action, "Message"): rule.Severity = UCase$(AtValue(action, "Trigger"))
            rule.Color = CLng(AtValue(action, "FillColor")): rule.AddNote = AtValue(action, "AddNote") = "true": rule.Audience = AtValue(action, "Audience")
        End If
    Next d
    If AtToXml(editor.Model) <> AtToXml(p) Then Err.Raise 5, , "В пакете есть настройки вне трёх поддержанных условий. Загрузка отклонена без изменений."
    Set AtReadTemplate = editor
End Function
