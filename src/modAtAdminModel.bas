Attribute VB_Name = "modAtAdminModel"
Option Explicit

Public Sub AtValidateActionPayload(ByVal d As CAtDefinition, ByVal issues As Collection)
    Dim color As String, message As String, r As Variant, targets As Long
    If AtValue(d, "OnOutcome") <> "FAIL" Then AtProblem issues, "ACTION_OUTCOME", d.Id, "OnOutcome", "", "Schema 3 requires explicit FAIL routing."
    message = Replace(Replace(AtValue(d, "Message"), vbCrLf, vbLf), vbCr, vbLf)
    If Len(message) = 0 Or Len("AccessTools" & vbLf & "[" & UCase$(AtValue(d, "Trigger")) & "] " & message) > 512 Then AtProblem issues, "ACTION_MESSAGE", d.Id, "Message", "", "Message plus Writer heading/severity must fit 512 characters."
    For Each r In d.References
        If r("Role") = "Field" Then targets = targets + 1
    Next r
    If targets <> 1 Then AtProblem issues, "ACTION_TARGET", d.Id, "Field", "", "Schema 3 scalar output requires exactly one Field target."
    If AtValue(d, "AddNote") <> "true" And AtValue(d, "AddNote") <> "false" Then AtProblem issues, "ACTION_NOTE", d.Id, "AddNote", "", "Use true or false."
    color = AtValue(d, "FillColor")
    If Not AtLiteralValid("Int64", color) Then
        AtProblem issues, "ACTION_COLOR", d.Id, "FillColor", color, "Use a canonical RGB integer 0-16777215."
    ElseIf CDec(color) < 0 Or CDec(color) > 16777215 Then
        AtProblem issues, "ACTION_COLOR", d.Id, "FillColor", color, "RGB integer outside range."
    End If
    If AtValue(d, "Trigger") = "Approval" Or AtValue(d, "Channel") <> IIf(AtValue(d, "AddNote") = "true", "Comment", "Results") Then AtProblem issues, "ACTION_PROFILE", d.Id, "Channel", "", "Scalar FAIL actions support Error/Warning/Info and explicit fill/Note channels."
End Sub

Public Function AtDefinitionById(ByVal p As CAtPackage, ByVal kind As String, ByVal id As String) As CAtDefinition
    Dim d As CAtDefinition
    For Each d In p.Definitions
        If d.Kind = kind And d.Id = id Then Set AtDefinitionById = d: Exit Function
    Next d
    Err.Raise 5, , "Definition not found: " & kind & "/" & id
End Function

Public Function AtReferenceId(ByVal d As CAtDefinition, ByVal role As String) As String
    Dim r As Variant, count As Long
    For Each r In d.References
        If r("Role") = role Then count = count + 1: AtReferenceId = r("Id")
    Next r
    If count <> 1 Then Err.Raise 5, , "Exactly one " & role & " reference required."
End Function

' This small editor constructs a standard AST; the runtime never knows its IDs.
Public Function AtAdminDefinition(ByVal sheet As String, ByVal address As String, ByVal fieldName As String, ByVal expected As String, ByVal message As String, ByVal severity As String, ByVal color As Long, ByVal addNote As Boolean) As CAtPackage
    Dim p As CAtPackage, d As CAtDefinition, n As CAtNode, trigger As String
    Select Case severity
        Case "ERROR": trigger = "Error"
        Case "WARNING": trigger = "Warning"
        Case "INFO": trigger = "Info"
        Case Else: Err.Raise 5, , "Severity must be ERROR, WARNING or INFO."
    End Select
    Set p = AtNewPackage("ADMIN_PACKAGE", "Text equality configuration"): p.SchemaVersion = "3.0.0"
    Set d = AtNewDefinition("Template", "TEMPLATE", "Template"): p.Definitions.Add d
    Set d = AtNewDefinition("Locator", "LOCATION", "Field location")
    d.Values.Add "Strategy", "A1": d.Values.Add "Sheet", sheet: d.Values.Add "Address", address
    AtAddReference d, "OwnerTemplate", "Template", "TEMPLATE", "1.0.0": p.Definitions.Add d
    Set d = AtNewDefinition("Field", "FIELD_VALUE", fieldName): d.Values.Add "DataType", "Text"
    AtAddReference d, "OwnerTemplate", "Template", "TEMPLATE", "1.0.0"
    AtAddReference d, "Locator", "Locator", "LOCATION", "1.0.0": p.Definitions.Add d
    Set d = AtNewDefinition("Expression", "EXPRESSION", "Text equals")
    d.Values.Add "ReturnType", "Bool": d.Values.Add "Deferred", "false"
    d.RootNodeId = "EQUAL": d.ScopeTemplateId = "TEMPLATE"
    Set n = AtNewNode("FIELD", "FieldRef", "Text")
    n.Values.Add "FieldId", "FIELD_VALUE": n.Values.Add "FieldVersion", "1.0.0": d.Nodes.Add n
    Set n = AtNewNode("EXPECTED", "TypedLiteral", "Text")
    n.Values.Add "State", "OK": n.Values.Add "Value", expected: d.Nodes.Add n
    Set n = AtNewNode("EQUAL", "Call", "Bool")
    n.Values.Add "OperationId", "OP_TEXT_EQUAL": n.Values.Add "OperationVersion", "1.0.0"
    n.Parameters.Add "MODE", AtTypedParameter("Text", "binary")
    n.Children.Add "FIELD": n.Children.Add "EXPECTED": d.Nodes.Add n: p.Definitions.Add d
    Set d = AtNewDefinition("Rule", "RULE_EQUAL", "Text equality")
    AtAddReference d, "OwnerTemplate", "Template", "TEMPLATE", "1.0.0"
    AtAddReference d, "Expression", "Expression", "EXPRESSION", "1.0.0", "Bool": p.Definitions.Add d
    Set d = AtNewDefinition("Action", "ACTION_FAIL", "Failure mark")
    d.Values.Add "Trigger", trigger: d.Values.Add "Audience", "Initiator"
    d.Values.Add "Channel", IIf(addNote, "Comment", "Results")
    d.Values.Add "OnOutcome", "FAIL": d.Values.Add "Message", message
    d.Values.Add "FillColor", CStr(color): d.Values.Add "AddNote", IIf(addNote, "true", "false")
    AtAddReference d, "OwnerTemplate", "Template", "TEMPLATE", "1.0.0"
    AtAddReference d, "Rule", "Rule", "RULE_EQUAL", "1.0.0"
    AtAddReference d, "Field", "Field", "FIELD_VALUE", "1.0.0", "Text": p.Definitions.Add d
    Set AtAdminDefinition = p
End Function

Public Function AtEvaluateRule(ByVal p As CAtPackage, ByVal ruleId As String, ByVal context As CAtEvalContext) As CAtRuleOutcome
    Dim rule As CAtDefinition, engine As New CAtEvaluator, outcome As New CAtRuleOutcome
    Set rule = AtDefinitionById(p, "Rule", ruleId)
    outcome.Initialize ruleId, engine.Evaluate(p, AtReferenceId(rule, "Expression"), context)
    Set AtEvaluateRule = outcome
End Function

Public Function AtOutcomeIntents(ByVal p As CAtPackage, ByVal outcome As CAtRuleOutcome) As Collection
    Dim intents As New Collection, d As CAtDefinition, field As CAtDefinition, locator As CAtDefinition, intent As CAtActionIntent
    If p.SchemaVersion <> "3.0.0" Then Err.Raise 5, , "Explicit schema 3 action payload required."
    If outcome.Status = "CANNOT_EVALUATE" Then Err.Raise 5, , "Incomplete evaluation cannot write or clear marks."
    If outcome.Status = "FAIL" Then
        For Each d In p.Definitions
            If d.Kind = "Action" Then
                If AtReferenceId(d, "Rule") = outcome.RuleId Then
                    Set field = AtDefinitionById(p, "Field", AtReferenceId(d, "Field"))
                    Set locator = AtDefinitionById(p, "Locator", AtReferenceId(field, "Locator"))
                    If AtValue(locator, "Strategy") <> "A1" Then Err.Raise 5, , "Action targeting currently supports A1 only."
                    Set intent = New CAtActionIntent
                    intent.ActionInstanceKey = d.Id
                    intent.Sheet = AtValue(locator, "Sheet"): intent.Address = AtValue(locator, "Address")
                    intent.Severity = UCase$(AtValue(d, "Trigger"))
                    intent.Message = AtValue(d, "Message"): intent.FillColor = CLng(AtValue(d, "FillColor"))
                    intent.AddNote = AtValue(d, "AddNote") = "true": intents.Add intent
                End If
            End If
        Next d
    End If
    Set AtOutcomeIntents = intents
End Function
