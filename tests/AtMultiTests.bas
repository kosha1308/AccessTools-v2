Attribute VB_Name = "AtMultiTests"
Option Explicit
Private passed As Long, failed As Long, log As String, book As Object, root As String
Public Function AtRunMultiTests(ByVal repo As String, ByVal scratch As String) As String
    On Error GoTo Broken
    Dim editor As Object, file As Object
    Set editor = Application.Run("'" & ThisWorkbook.Name & "'!AtNewTemplateEditor", scratch, False)
    passed = 0: failed = 0: log = "": root = scratch
    Set book = Application.Workbooks.Add: book.Worksheets(1).Name = "Input"
    TestRequired
    TestModel
    TestRun
    TestUi
    book.Close False: Set book = Nothing
    Set file = CreateObject("Scripting.FileSystemObject").CreateTextFile(scratch & "\multi-assertions.txt", True, True): file.Write log: file.Close
    AtRunMultiTests = IIf(failed = 0, "PASS|MULTI|", "FAIL|MULTI|") & passed & "|" & failed & IIf(failed = 0, "", "|" & log): Exit Function
Broken:
    AtRunMultiTests = "FAIL|MULTI|" & Err.Number & "|" & Err.Description & "|" & log
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
End Function
Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    If ok Then passed = passed + 1 Else failed = failed + 1
    log = log & IIf(ok, "PASS ", "FAIL ") & label & vbCrLf
End Sub
Private Function Fixture() As CAtTemplateEditor
    Dim editor As CAtTemplateEditor, field As CAtTemplateField, rule As CAtTemplateRule
    Set editor = AtNewTemplateEditor(root, False): editor.TemplateName = "Синтетический шаблон"
    Set field = editor.AddField("Статус", "Input", "B2")
    Set rule = editor.AddRule(field.Id, "Ожидается good"): rule.ExpectedText = "good": rule.Message = "Укажите good.": rule.Severity = "WARNING": rule.Color = 65535
    Set rule = editor.AddRule(field.Id, "Запрещено bad"): rule.Condition = "NOT_EQUAL": rule.ExpectedText = "bad": rule.Message = "Значение bad запрещено.": rule.Color = 255
    Set field = editor.AddField("Оборудование", "Input", "C2")
    Set rule = editor.AddRule(field.Id, "Нужно оборудование"): rule.Condition = "REQUIRED": rule.Message = "Заполните оборудование."
    Set Fixture = editor
End Function
Private Sub TestRequired()
    Dim registry As New CAtOperationRegistry, args As Collection, result As CAtEvalResult, value As CAtEvalResult, state As Variant
    For Each state In Array("OK", "EMPTY", "MISSING", "INVALID", "UNAVAILABLE", "ERROR", "CONFLICT", "LIMIT", "NOT_APPLICABLE")
        Set args = New Collection: Set value = AtEvalValue("Text", CStr(state), IIf(state = "OK", "x", ""))
        args.Add value: Set result = registry.ExecuteOperation("OP_TEXT_REQUIRED", "1.0.0", args, AtDictionary())
        If state = "OK" Or state = "EMPTY" Then
            Check result.IsComplete And result.Value = IIf(state = "EMPTY", "false", "true"), "required explicitly handles " & state
        Else
            Check result.State = state And Not result.IsComplete, "required preserves " & state
        End If
    Next state
    Set args = New Collection: args.Add AtEvalValue("Text", "OK", "   ")
    Set result = registry.ExecuteOperation("OP_TEXT_REQUIRED", "1.0.0", args, AtDictionary())
    Check result.IsComplete And result.Value = "false", "required whitespace uses explicit ASCII Trim"
    Set value = AtEvalValue("Text", "EMPTY", ""): value.AddRequirement "OBJ", "NODE", "Field", "UNAVAILABLE": value.AddDiagnostic "E_INTERNAL", "OBJ", "NODE", "technical"
    Set args = New Collection: args.Add value
    Set result = registry.ExecuteOperation("OP_TEXT_REQUIRED", "1.0.0", args, AtDictionary())
    Check Not result.IsComplete And result.Requirements.Count = 1 And result.Diagnostics.Count = 1, "EMPTY handling never hides technical obligations"
    Set args = New Collection: args.Add AtEvalValue("Bool", "OK", "true")
    Set result = registry.ExecuteOperation("OP_TEXT_REQUIRED", "1.0.0", args, AtDictionary())
    Check result.State = "ERROR", "required refuses wrong argument type"
    Check Not registry.HandlesEmpty("OP_TEXT_EQUAL", "1.0.0") And Not registry.HandlesEmpty("OP_TEXT_REQUIRED", "2.0.0"), "only exact trusted required operation can consume EMPTY"
End Sub
Private Sub TestModel()
    Dim editor As CAtTemplateEditor, loaded As CAtTemplateEditor, field As CAtTemplateField, rule As CAtTemplateRule, copy As CAtTemplateRule
    Dim p As CAtPackage, q As CAtPackage, issues As Collection, registry As New CAtOperationRegistry, xml As String, id As String, blocked As Boolean, old As New CAtAdminEditor
    Set editor = Fixture(): Set p = editor.Model
    Check p.SchemaVersion = "4.0.0" And editor.Fields.Count = 2 And editor.Rules.Count = 3, "separate fields and rules compile schema4"
    Check editor.Rules(1).FieldId = editor.Rules(2).FieldId, "several rules share exactly one field"
    Set q = AtFromXml(AtToXml(p), registry.Catalog, issues)
    Check Not q Is Nothing And Not AtHasErrors(issues) And AtToXml(q) = AtToXml(p), "full schema4 XML roundtrip"
    editor.SaveFile root & "\multi.xml": Set loaded = AtNewTemplateEditor(root, False): loaded.LoadFile root & "\multi.xml"
    Check AtToXml(loaded.Model) = AtToXml(editor.Model), "all fields rules links and settings reload"
    Set field = editor.Fields(1): id = field.Id: field.DisplayName = "Новое имя": field.Address = "D4"
    Check field.Id = id And editor.Rules(1).FieldId = id And editor.Rules(2).FieldId = id, "renaming and moving shared field preserves references"
    Set rule = editor.Rules(1): id = rule.Id: rule.DisplayName = "Переименовано": rule.Enabled = False
    Set copy = editor.CopyRule(id)
    Check copy.Id <> id And copy.FieldId = rule.FieldId And editor.Fields.Count = 2, "rule copy gets independent identity and shared field"
    Check Not copy.Enabled And copy.Message = rule.Message And copy.Color = rule.Color, "copy preserves settings and disabled state"
    editor.DeleteRule copy.Id
    Check editor.Rules.Count = 3 And rule.Id = id, "delete removes only selected rule and preserves other identities"
    editor.SelectRule id: editor.BindRuleField editor.Fields(2).Id
    Check rule.Id = id And rule.FieldId = editor.Fields(2).Id, "rule rebinds existing field without cloning"
    editor.SaveFile root & "\multi-edited.xml": loaded.LoadFile root & "\multi-edited.xml"
    Check Not loaded.Rules(1).Enabled And loaded.Rules(1).Id = id And loaded.Rules(1).FieldId = editor.Fields(2).Id, "disabled state rename and rebind survive save load"
    Set p = loaded.Model: p.Definitions.Add AtNewDefinition("Template", "EXTRA", "unsupported")
    xml = AtToXml(p): AtSaveText root & "\multi-unsupported.xml", xml
    xml = AtToXml(loaded.Model)
    On Error Resume Next
    loaded.LoadFile root & "\multi-unsupported.xml": blocked = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    Check blocked And AtToXml(loaded.Model) = xml, "unsupported configuration rejected atomically without data loss"
    Set p = editor.Model: AtDefinitionById(p, "Rule", id).Values.Remove "Enabled": Set issues = AtValidate(p, registry.Catalog)
    Check AtHasErrors(issues), "schema4 requires explicit enabled flag"
    old.SheetName = "Input": old.Address = "B2": old.ExpectedText = "yes": old.SaveFile root & "\multi-v3.xml"
    xml = AtToXml(old.Model): loaded.LoadFile root & "\multi-v3.xml"
    Check loaded.Fields(1).Id = "FIELD_VALUE" And loaded.Rules(1).Id = "RULE_EQUAL" And loaded.Model.SchemaVersion = "4.0.0", "known r5 profile explicitly imports preserving FieldId and RuleId"
    Set q = AtLoad(root & "\multi-v3.xml", registry.Catalog, issues)
    Check AtToXml(q) = xml And q.SchemaVersion = "3.0.0", "import never rewrites or reinterprets original schema3 file"
    Set editor = AtNewTemplateEditor(root, False): Set field = editor.AddField("F", "Input", "A1")
    Dim i As Long
    For i = 1 To 32: Set rule = editor.AddRule(field.Id, "R" & i): Next i
    blocked = False
    On Error Resume Next
    Set rule = editor.AddRule(field.Id, "overflow"): blocked = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    Check blocked And editor.Rules.Count = 32, "editor never expands Writer 32-intent budget"
End Sub
Private Sub AtSaveText(ByVal path As String, ByVal text As String)
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.WriteText text: stream.SaveToFile path, 1: stream.Close
End Sub
Private Sub TestRun()
    Dim editor As CAtTemplateEditor, result As CAtAdminPreview, field As CAtTemplateField, rule As CAtTemplateRule, aliasField As CAtTemplateField
    Dim cell As Object, initial As Long, id As String, before As Boolean
    Set editor = Fixture(): Set cell = book.Worksheets(1).Range("B2"): initial = RGB(20, 30, 40)
    cell.Value2 = "bad": cell.Interior.Color = initial: book.Worksheets(1).Range("C2").Value2 = "asset"
    book.Worksheets(1).Range("Z1").Formula = "=1+2": book.Worksheets(1).Range("Z2").AddComment "User note"
    book.SaveAs root & "\multi.xlsx", 51: before = book.Saved
    Set result = editor.Preview(book)
    Check result.Findings.Count = 3 And result.Findings(1).Outcome.Status = "FAIL" And result.Findings(2).Outcome.Status = "FAIL" And result.Findings(3).Outcome.Status = "PASS", "two FAIL plus PASS retained in whole-run findings"
    Check cell.Interior.Color = initial And book.Saved = before, "whole preview is read only including Saved"
    Check InStr(result.InitiatorText, "Укажите good") > 0 And InStr(result.InitiatorText, "bad запрещено") > 0 And InStr(result.InitiatorText, "PASS") = 0, "initiator copy retains both errors without service status"
    Set result = editor.Apply(book, "")
    Check cell.Interior.Color = 255 And InStr(result.Output, AtUiStatus("APPLIED")) > 0, "single grouped target uses highest severity ERROR red"
    editor.SelectRule editor.Rules(3).Id: Set result = editor.PreviewSelected(book)
    Check result.Findings.Count = 1 And cell.Interior.Color = 255, "selected-rule testing cannot clear other rules marks"
    cell.Value2 = "other": Set result = editor.Apply(book, "")
    Check result.Findings(1).Outcome.Status = "FAIL" And result.Findings(2).Outcome.Status = "PASS" And cell.Interior.Color = 65535, "partial correction retains one warning and changes severity color"
    cell.Formula = "=""good""": Set result = editor.Apply(book, "")
    Check result.Outcome.Status = "PASS" And cell.Interior.Color = initial And cell.Formula = "=""good""", "all PASS restores original fill preserving corrected formula"
    Check editor.NeedsJournal, "all PASS releases completed journal"
    cell.Value2 = "bad": Set result = editor.Apply(book, "")
    editor.Fields(2).Sheet = "Missing": cell.Value2 = "good": Set result = editor.Apply(book, "")
    Check result.Outcome.Status = "CANNOT_EVALUATE" And cell.Interior.Color = 255 And InStr(result.Output, "заблокирована") > 0, "one missing required field freezes whole annotation set"
    Check result.Findings(3).Outcome.Result.Requirements.Count > 0, "missing required field technical requirement retained"
    editor.Fields(2).Sheet = "Input": book.Worksheets(1).Range("C2").ClearContents: Set result = editor.Preview(book)
    Check result.Findings(3).Outcome.Status = "FAIL" And result.Findings(3).Outcome.Result.IsComplete, "actual EMPTY required field is ordinary FAIL not technical CE"
    Check cell.Interior.Color = 255, "preview never clears old marks even after other failures change"
    book.Worksheets(1).Range("C2").Value2 = "asset": cell.Value2 = "bad": Set result = editor.Apply(book, "")
    Set field = editor.Fields(1): id = field.Id: field.Address = "D4": book.Worksheets(1).Range("D4").Value2 = "bad"
    Set result = editor.Apply(book, "")
    Check cell.Interior.Color = initial And book.Worksheets(1).Range("D4").Interior.Color = 255 And editor.Rules(1).FieldId = id And editor.Rules(2).FieldId = id, "moving shared field reconciles old physical mark once"
    Set aliasField = editor.AddField("Alias", "Input", "D4"): Set rule = editor.AddRule(aliasField.Id, "Alias check")
    rule.ExpectedText = "good": rule.Message = "Alias error": rule.Severity = "INFO": rule.Color = RGB(213, 232, 212)
    Set result = editor.Apply(book, "")
    Check result.Findings.Count = 4 And book.Worksheets(1).Range("D4").Interior.Color = 255, "distinct FieldIds with same address share one highest-priority target"
    Dim clearText As String: clearText = editor.Clear(book)
    book.SaveAs root & "\multi.xls", 56
    For Each rule In editor.Rules: rule.AddNote = True: Next rule
    Set result = editor.Apply(book, "")
    Set cell = book.Worksheets(1).Range("D4")
    Check Not cell.Comment Is Nothing And InStr(cell.Comment.Text, "Укажите good") > 0 And InStr(cell.Comment.Text, "bad запрещено") > 0 And InStr(cell.Comment.Text, "Alias error") > 0, "one supported xls Note contains every failing rule message"
    Check book.Worksheets(1).Comments.Count = 2, "one own Note plus untouched user Note"
    editor.Rules(1).Enabled = False: Set result = editor.Apply(book, "")
    Check InStr(cell.Comment.Text, "Укажите good") = 0 And InStr(cell.Comment.Text, "bad запрещено") > 0, "disabled rule removed without dropping remaining errors"
    cell.Value2 = "good": Set result = editor.Apply(book, "")
    Check result.Outcome.Status = "PASS" And cell.Comment Is Nothing And cell.Interior.Pattern = -4142, "all remaining PASS clean single merged-message annotation"
    Check book.Worksheets(1).Range("Z1").Formula = "=1+2" And book.Worksheets(1).Range("Z2").Comment.Text = "User note", "user data formulas and Notes preserved through multi cycles"
    cell.Value2 = "bad": Set result = editor.Apply(book, ""): cell.Interior.Color = RGB(1, 2, 3)
    Set result = editor.Apply(book, "")
    Check InStr(result.Output, "CONFLICT") > 0 Or InStr(result.Diagnostics, "TARGET_CHANGED") > 0, "formatting conflict visibly separated from rule FAIL"
    Check cell.Interior.Color = RGB(1, 2, 3) And Not cell.Comment Is Nothing, "conflict preserves current formatting and evidence"
End Sub
Private Sub TestUi()
    Dim editor As CAtTemplateEditor, result As CAtAdminPreview, id As String, count As Long, loaded As CAtTemplateEditor
    Set editor = Fixture(): AtUiOpenTemplate editor
    Check AtAdminMain.cboCondition.ListCount = 3 And AtAdminMain.lstNavigation.ListCount = 3, "current UI offers three working conditions and all rules"
    Check AtAdminMain.cmdAddField.Visible And AtAdminMain.cmdAddRule.Visible And AtAdminMain.cmdSelectedPreview.Visible, "real field/rule commands visible"
    Dim second As Object
    Set second = book.Worksheets.Add: second.Name = "OtherInput"
    AtAdminMain.lstNavigation.ListIndex = 0: AtUiMultiSelectRule AtAdminMain
    editor.Fields(2).Sheet = "OtherInput"
    AtAdminMain.lstNavigation.ListIndex = 2: AtUiMultiSelectRule AtAdminMain
    Check editor.Fields(2).Sheet = "OtherInput" And AtAdminMain.cboSheet.Value = "OtherInput", "rule switch preserves different worksheet locator"
    AtAdminMain.lstNavigation.ListIndex = 0: AtUiMultiSelectRule AtAdminMain
    Check editor.Fields(1).Sheet = "Input" And AtAdminMain.cboSheet.Value = "Input", "switch back preserves shared first worksheet locator"
    editor.Fields(2).Sheet = "Input"
    id = editor.Rules(1).Id: AtAdminMain.lstNavigation.ListIndex = 0: AtUiMultiSelectRule AtAdminMain
    AtAdminMain.txtRuleName.Text = "Renamed in UI": AtAdminMain.chkRuleEnabled.Value = False: AtUiChanged AtAdminMain
    Check editor.GetRule(id).DisplayName = "Renamed in UI" And Not editor.GetRule(id).Enabled, "UI rename and enable flag edit stable selected rule"
    count = editor.Fields.Count: AtUiMultiCommand AtAdminMain, "COPY_RULE"
    Check editor.Rules.Count = 4 And editor.Fields.Count = count And editor.SelectedRuleId <> id, "UI copies rule without field duplication"
    AtUiMultiCommand AtAdminMain, "DELETE_RULE"
    Check editor.Rules.Count = 3, "UI deletes only copied rule"
    AtAdminMain.cboFieldFilter.ListIndex = 1: AtUiMultiFilter AtAdminMain
    Check AtAdminMain.lstNavigation.ListCount = 2, "field filter displays both shared rules"
    AtUiMultiCommand AtAdminMain, "ADD_FIELD"
    Check editor.Fields.Count = 3 And AtAdminMain.lstNavigation.ListCount = 0, "new field created without artificial rule copy"
    AtUiMultiCommand AtAdminMain, "ADD_RULE"
    Check editor.Rules.Count = 4 And editor.GetRule(editor.SelectedRuleId).FieldId = editor.SelectedFieldId, "new rule binds selected field"
    editor.SaveFile root & "\multi-ui.xml": Set loaded = AtNewTemplateEditor(root, False): loaded.LoadFile root & "\multi-ui.xml"
    Check loaded.Fields.Count = 3 And loaded.Rules.Count = 4, "UI settings save entire template not selected rule"
    AtUiLayout AtAdminMain, 900, 650
    Check AtAdminMain.Width <= 900 And AtAdminMain.Zoom = 100 And AtAdminMain.pnlBody.ScrollBars <> 3, "new editor more compact without horizontal scrollbar"
    AtUiLayout AtAdminMain, 480, 300
    Check AtAdminMain.cmdSelectedPreview.Top >= AtAdminMain.pnlBody.Top + AtAdminMain.pnlBody.Height, "selected and full test commands accessible in compact viewport"
    Set editor = Fixture(): Set result = editor.Preview(book): Load AtAdminResults: AtUiFillResult AtAdminResults, result
    Check AtAdminResults.lstResults.ListCount = 3 And AtAdminResults.lstResults.Visible, "results list retains independent rule findings"
    Check AtAdminResults.txtRemark.Text = result.InitiatorText, "copy uses displayed business messages only"
    AtUiHideMain: Unload AtAdminMain: Unload AtAdminResults
End Sub
