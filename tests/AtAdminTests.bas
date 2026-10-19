Attribute VB_Name = "AtAdminTests"
Option Explicit
Private passed As Long, failed As Long, transcript As String
Private book As Object

Public Function AtRunAdminTests(ByVal repo As String, ByVal scratch As String) As String
    On Error GoTo Broken
    passed = 0: failed = 0: transcript = ""
    Dim p As CAtPackage, q As CAtPackage, registry As New CAtOperationRegistry, issues As Collection
    Dim outcome As CAtRuleOutcome, intents As Collection, journal As CAtOwnershipJournal, plan As CAtActionPlan
    Dim receipts As Collection, cell As Object, fso As Object, file As Object, before As String, count As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set p = Application.Run("'" & ThisWorkbook.Name & "'!AtAdminDefinition", "Request", "B2", "Name", "expected", "Fix value", "ERROR", 255, False)
    Check p.SchemaVersion = "3.0.0", "explicit action schema"
    Set issues = AtValidate(p, registry.Catalog): Check Not AtHasErrors(issues), "standard typed AST validates"
    Check AtSaveNew(p, scratch & "\admin-rule.xml", registry.Catalog, issues), "save configuration"
    Set q = AtLoad(scratch & "\admin-rule.xml", registry.Catalog, issues)
    Check Not q Is Nothing, "load configuration"
    Check AtToXml(p) = AtToXml(q), "complete AST/action roundtrip"
    Set book = Application.Workbooks.Add(): book.Worksheets(1).Name = "Request"
    Set cell = book.Worksheets(1).Range("B2"): cell.Value2 = "expected"
    book.Worksheets(1).Range("C2").Formula = "=1+2": before = book.Worksheets(1).Range("C2").Formula
    Set outcome = Evaluate(q): Check outcome.Status = "PASS", "loaded configuration executes PASS"
    cell.Value2 = "other": Set outcome = Evaluate(q): Check outcome.Status = "FAIL", "Excel edit switches to FAIL"
    Set intents = AtOutcomeIntents(q, outcome)
    Check intents.Count = 1 And intents(1).Message = "Fix value", "FAIL creates structured action"
    fso.CreateFolder scratch & "\admin-journal"
    Set journal = AtNewOwnershipJournal(book, scratch & "\admin-journal")
    count = Application.Workbooks.Count
    Set plan = AtPlanActions(book, intents, "RUN_FIRST", journal)
    Check plan.Status = "READY", "fill planning uses unchanged capability gates"
    Set receipts = AtApplyPlan(plan, journal)
    Check receipts(1).Status = "APPLIED" And cell.Interior.Color = 255, "FAIL fill applied"
    Set plan = AtPlanActions(book, intents, "RUN_REPEAT", journal): Set receipts = AtApplyPlan(plan, journal)
    Check receipts(1).Status = "SKIPPED", "repeat has no duplicate output"
    cell.Value2 = "expected": Set outcome = Evaluate(q)
    Check outcome.Status = "PASS", "Excel edit switches back to PASS"
    Set plan = AtPlanActions(book, AtOutcomeIntents(q, outcome), "RUN_PASS", journal)
    Check plan.Status = "READY", "PASS after input correction permits owned cleanup"
    Set receipts = AtApplyPlan(plan, journal)
    Check cell.Value2 = "expected" And receipts(1).Status = "ROLLED_BACK", "cleanup preserves corrected input"
    Check cell.Interior.Pattern = -4142, "PASS restores original fill"
    Check cell.Value2 = "expected" And book.Worksheets(1).Range("C2").Formula = before, "values and formulas preserved"
    Check Application.Workbooks.Count = count, "execution opens no other workbooks"
    AtDefinitionById(q, "Expression", "EXPRESSION").Nodes(2).Values("Value") = "changed"
    Set outcome = Evaluate(q): Check outcome.Status = "FAIL", "condition changed without VBA"
    cell.Value2 = "changed": Set outcome = Evaluate(q): Check outcome.Status = "PASS", "new condition executes"
    AtDefinitionById(q, "Locator", "LOCATION").Values("Address") = "D5"
    book.Worksheets(1).Range("D5").Value2 = "changed"
    Set outcome = Evaluate(q): Check outcome.Status = "PASS", "coordinate changed without code"
    Check AtDefinitionById(q, "Field", "FIELD_VALUE").Id = "FIELD_VALUE" And outcome.RuleId = "RULE_EQUAL", "FieldId and RuleId stable"
    book.Worksheets(1).Range("D5").ClearContents: Set outcome = Evaluate(q)
    Check outcome.Status = "CANNOT_EVALUATE", "EMPTY is incomplete"
    book.Worksheets(1).Range("D5").Value2 = 42: Set outcome = Evaluate(q)
    Check outcome.Status = "CANNOT_EVALUATE", "INVALID numeric Text is incomplete"
    AtDefinitionById(q, "Locator", "LOCATION").Values("Sheet") = "Absent"
    Set outcome = Evaluate(q): Check outcome.Status = "CANNOT_EVALUATE", "MISSING is incomplete"
    Dim rejected As Boolean
    On Error Resume Next
    Set intents = AtOutcomeIntents(q, outcome): rejected = Err.Number <> 0: Err.Clear
    On Error GoTo Broken
    Check rejected, "incomplete outcome cannot clear old marks"
    AtDefinitionById(q, "Locator", "LOCATION").Values("Sheet") = "Request"
    book.Worksheets(1).Range("D5").Value2 = "wrong"
    AtDefinitionById(q, "Action", "ACTION_FAIL").Values("Message") = ChrW(&H416) & " changed"
    AtDefinitionById(q, "Action", "ACTION_FAIL").Values("FillColor") = "65535"
    AtDefinitionById(q, "Action", "ACTION_FAIL").Values("Message") = String$(492, "x")
    Set issues = AtValidate(q, registry.Catalog): Check Not AtHasErrors(issues), "ERROR message exact Writer envelope fits"
    AtDefinitionById(q, "Action", "ACTION_FAIL").Values("Message") = String$(493, "x")
    Set issues = AtValidate(q, registry.Catalog): Check AtHasErrors(issues), "message exceeding Writer envelope rejected"
    AtDefinitionById(q, "Action", "ACTION_FAIL").Values("Message") = ChrW(&H416) & " changed"
    AtDefinitionById(q, "Action", "ACTION_FAIL").References.Remove 3
    Set issues = AtValidate(q, registry.Catalog): Check AtHasErrors(issues), "missing Action Field target rejected"
    AtAddReference AtDefinitionById(q, "Action", "ACTION_FAIL"), "Field", "Field", "FIELD_VALUE", "1.0.0", "Text"
    Set outcome = Evaluate(q): Set intents = AtOutcomeIntents(q, outcome)
    Check intents(1).Message = ChrW(&H416) & " changed" And intents(1).FillColor = 65535, "message Unicode and color edited without VBA"
    Check AtSaveNew(q, scratch & "\admin-rule-edited.xml", registry.Catalog, issues), "edited configuration saved"
    Set p = AtLoad(scratch & "\admin-rule-edited.xml", registry.Catalog, issues)
    Check AtToXml(p) = AtToXml(q), "edited configuration full roundtrip"
    q.SchemaVersion = "2.0.0": Set issues = AtValidate(q, registry.Catalog)
    Check AtHasErrors(issues), "new properties not reinterpreted as schema 2"
    q.SchemaVersion = "3.0.0"
    AtDefinitionById(q, "Action", "ACTION_FAIL").Values("FillColor") = "16777216"
    Set issues = AtValidate(q, registry.Catalog): Check AtHasErrors(issues), "invalid action RGB rejected"
    AtDefinitionById(q, "Action", "ACTION_FAIL").Values("FillColor") = "65535"
    AtDefinitionById(q, "Action", "ACTION_FAIL").Values("AddNote") = "true"
    AtDefinitionById(q, "Action", "ACTION_FAIL").Values("Channel") = "Comment"
    book.SaveAs scratch & "\admin-synthetic.xlsx", 51
    Set outcome = Evaluate(q): Set intents = AtOutcomeIntents(q, outcome)
    Set plan = AtPlanActions(book, intents, "RUN_NOTE_XLSX", journal)
    Check plan.Code = "CAPABILITY_UNSUPPORTED", "xlsx Note explicitly refused"
    Check book.Worksheets(1).Range("D5").Interior.Pattern = -4142, "unsupported Notes do not downgrade to fill"
    book.Close False: Set book = Nothing
    TestLegacy q, scratch
    TestReconciliation scratch, False
    TestReconciliation scratch, True
    Set file = fso.CreateTextFile(scratch & "\admin-assertions.txt", True, True): file.Write transcript: file.Close
    AtRunAdminTests = IIf(failed = 0, "PASS|ADMIN|", "FAIL|ADMIN|") & passed & "|" & failed & "|" & IIf(failed = 0, "", transcript)
    Exit Function
Broken:
    AtRunAdminTests = "FAIL|ADMIN|" & Err.Number & "|" & Err.Description & "|" & transcript
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
End Function

Private Sub Check(ByVal condition As Boolean, ByVal label As String)
    If condition Then passed = passed + 1 Else failed = failed + 1
    transcript = transcript & IIf(condition, "PASS ", "FAIL ") & label & vbCrLf
End Sub

Private Function Evaluate(ByVal p As CAtPackage) As CAtRuleOutcome
    Dim resolver As CAtExcelResolver, context As CAtEvalContext, issues As Collection
    Set resolver = AtNewExcelResolver(p, "TEMPLATE", AtDictionary())
    Set context = resolver.Resolve(book, issues)
    Set Evaluate = AtEvaluateRule(p, "RULE_EQUAL", context)
End Function

Private Sub TestLegacy(ByVal p As CAtPackage, ByVal scratch As String)
    Dim fso As Object, journal As CAtOwnershipJournal, outcome As CAtRuleOutcome, plan As CAtActionPlan, receipts As Collection, cell As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set book = Application.Workbooks.Add(): book.Worksheets(1).Name = "Request"
    Set cell = book.Worksheets(1).Range("D5"): cell.Value2 = "wrong"
    book.SaveAs scratch & "\admin-legacy.xls", 56
    fso.CreateFolder scratch & "\admin-legacy-journal"
    Set journal = AtNewOwnershipJournal(book, scratch & "\admin-legacy-journal")
    Set outcome = Evaluate(p): Set plan = AtPlanActions(book, AtOutcomeIntents(p, outcome), "LEGACY_FAIL", journal)
    Check plan.Status = "READY", "xls fill plus Note passes capability profile"
    Set receipts = AtApplyPlan(plan, journal)
    Check receipts(1).Status = "APPLIED", "xls action applied"
    Check Not cell.Comment Is Nothing, "xls Note exists"
    Check cell.Comment.Text = "AccessTools" & vbLf & "[ERROR] " & ChrW(&H416) & " changed", "xls Note retains Unicode message and heading"
    Check cell.Interior.Color = 65535 And cell.Value2 = "wrong", "xls fill preserves value"
    cell.Value2 = "changed": Set outcome = Evaluate(p)
    Set plan = AtPlanActions(book, AtOutcomeIntents(p, outcome), "LEGACY_PASS", journal)
    Check outcome.Status = "PASS" And plan.Status = "READY", "xls corrected input permits cleanup"
    Set receipts = AtApplyPlan(plan, journal)
    Check receipts(1).Status = "ROLLED_BACK" And cell.Value2 = "changed", "xls cleanup preserves corrected input"
    Check cell.Comment Is Nothing And cell.Interior.Pattern = -4142, "xls Note and fill restored"
    book.Close False: Set book = Nothing
End Sub

Private Sub TestReconciliation(ByVal scratch As String, ByVal legacy As Boolean)
    Dim p As CAtPackage, journal As CAtOwnershipJournal, plan As CAtActionPlan, stale As CAtActionPlan
    Dim receipts As Collection, outcome As CAtRuleOutcome, cell As Object, fso As Object, prefix As String, prior As Long
    Dim entry As Object, baseline As Object, stepBefore As Object, snapshot As Object, altered As Object, code As String
    prefix = IIf(legacy, "xls", "xlsx")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set book = Application.Workbooks.Add(): book.Worksheets(1).Name = "Request"
    Set cell = book.Worksheets(1).Range("B2")
    cell.Interior.Color = RGB(70, 80, 90): cell.Value2 = "wrong"
    book.Worksheets(1).Range("F1").AddComment "User note"
    book.SaveAs scratch & "\reconcile." & prefix, IIf(legacy, 56, 51)
    fso.CreateFolder scratch & "\reconcile-" & prefix
    Set journal = AtNewOwnershipJournal(book, scratch & "\reconcile-" & prefix)
    Set p = AtAdminDefinition("Request", "B2", "Name", "expected", "Fix", "ERROR", 255, legacy)
    Set outcome = Evaluate(p): Set plan = AtPlanActions(book, AtOutcomeIntents(p, outcome), "MARK", journal)
    Set receipts = AtApplyPlan(plan, journal)
    Check receipts(1).Status = "APPLIED", prefix & " FAIL mark"
    cell.Value2 = "still wrong"
    AtDefinitionById(p, "Action", "ACTION_FAIL").Values("Message") = "Updated"
    AtDefinitionById(p, "Action", "ACTION_FAIL").Values("FillColor") = "65535"
    Set outcome = Evaluate(p): Set plan = AtPlanActions(book, AtOutcomeIntents(p, outcome), "UPDATE", journal)
    Set receipts = AtApplyPlan(plan, journal)
    Check receipts(1).Status = "APPLIED" And cell.Value2 = "still wrong" And cell.Interior.Color = 65535, prefix & " FAIL edit updates own annotation without restoring old data"
    Set entry = journal.Entry(book.Worksheets(1).CodeName & "!$B$2")
    Set baseline = AtWUnpack(entry("Before")): Set stepBefore = AtWUnpack(entry("StepBefore"))
    Check Not baseline.Exists("Data1") And Not baseline.Exists("Formula1"), prefix & " restoration baseline excludes historical inputs"
    Check stepBefore("Data1") = AtWScalar("still wrong"), prefix & " write-ahead step records fresh input"
    Set snapshot = AtWCapture(cell, code): Set altered = AtRecordCopy(snapshot)
    altered("Data1") = "different": altered("Formula1") = "different"
    Check AtWOwnedEqual(snapshot, altered) And Not AtWEqual(snapshot, altered), prefix & " ownership and full precondition are distinct comparisons"
    altered.Remove "Merge1"
    Check Not AtWOwnedEqual(snapshot, altered), prefix & " missing structural evidence cannot confer ownership"
    If legacy Then Check InStr(cell.Comment.Text, "Updated") > 0, prefix & " own Note updated"
    cell.Formula = "=""expected""": cell.Calculate
    Set outcome = Evaluate(p): Set plan = AtPlanActions(book, AtOutcomeIntents(p, outcome), "FORMULA_PASS", journal)
    Check outcome.Status = "PASS" And plan.Status = "READY", prefix & " corrected formula produces PASS cleanup plan"
    Set stale = plan: prior = journal.RevisionCount
    cell.Formula = "=""later""": cell.Calculate
    Set receipts = AtApplyPlan(stale, journal)
    Check receipts(1).Status = "CONFLICT" And journal.RevisionCount = prior And cell.Formula = "=""later""", prefix & " stale cleanup rejects formula edit after planning without write"
    cell.Formula = "=""expected""": cell.Calculate
    Set outcome = Evaluate(p): Set plan = AtPlanActions(book, AtOutcomeIntents(p, outcome), "FRESH_PASS", journal)
    Set receipts = AtApplyPlan(plan, journal)
    Check receipts(1).Status = "ROLLED_BACK" And cell.Formula = "=""expected""" And cell.Interior.Color = RGB(70, 80, 90), prefix & " fresh cleanup preserves new formula and original user color"
    Check cell.Comment Is Nothing, prefix & " own Note removed on PASS"
    Set receipts = AtApplyPlan(AtPlanActions(book, AtOutcomeIntents(p, outcome), "REPEAT_PASS", journal), journal)
    Check receipts.Count = 0, prefix & " repeat PASS has no remaining actions"
    cell.Value2 = "wrong": Set outcome = Evaluate(p)
    Set plan = AtPlanActions(book, AtOutcomeIntents(p, outcome), "STALE_MARK", journal)
    cell.Value2 = "expected": prior = journal.RevisionCount
    Set receipts = AtApplyPlan(plan, journal)
    Check receipts(1).Status = "CONFLICT" And journal.RevisionCount = prior And cell.Interior.Color = RGB(70, 80, 90), prefix & " stale mark rejects changed input before any mutation"
    cell.Value2 = "wrong": Set outcome = Evaluate(p)
    Set receipts = AtApplyPlan(AtPlanActions(book, AtOutcomeIntents(p, outcome), "MARK_AGAIN", journal), journal)
    cell.Value2 = "expected": cell.Interior.Color = RGB(20, 30, 40)
    Set receipts = AtClearOwned(book, journal, "USER_FILL")
    Check receipts(1).Status = "CONFLICT" And cell.Value2 = "expected" And cell.Interior.Color = RGB(20, 30, 40), prefix & " input correction cannot hide foreign fill change"
    cell.Interior.Color = 65535
    If legacy Then
        cell.Comment.Text "User changed own note"
        Set receipts = AtClearOwned(book, journal, "USER_NOTE")
        Check receipts(1).Status = "CONFLICT" And cell.Comment.Text = "User changed own note", prefix & " changed Note protected despite corrected input"
        cell.Comment.Text "AccessTools" & vbLf & "[ERROR] Updated"
    Else
        cell.AddComment "User note on annotated cell"
        Set receipts = AtClearOwned(book, journal, "USER_NOTE")
        Check receipts(1).Status = "CONFLICT" And cell.Comment.Text = "User note on annotated cell", prefix & " newly added user Note protected"
        cell.Comment.Delete
    End If
    cell.Resize(1, 2).Merge
    Set receipts = AtClearOwned(book, journal, "USER_MERGE")
    Check receipts(1).Status = "CONFLICT" And cell.MergeArea.Columns.Count = 2, prefix & " merge structure change protected"
    cell.MergeArea.UnMerge
    Set receipts = AtClearOwned(book, journal, "CLEAN_AFTER_STRUCTURE")
    Check receipts(1).Status = "ROLLED_BACK" And cell.Value2 = "expected", prefix & " cleanup after restored annotation guards preserves corrected value"
    Check book.Worksheets(1).Range("F1").Comment.Text = "User note", prefix & " unrelated user Note preserved"
    cell.Value2 = "wrong": Set outcome = Evaluate(p)
    Set receipts = AtApplyPlan(AtPlanActions(book, AtOutcomeIntents(p, outcome), "MARK_BEFORE_MOVE", journal), journal)
    AtDefinitionById(p, "Locator", "LOCATION").Values("Address") = "D5"
    cell.Value2 = "edited old location": book.Worksheets(1).Range("D5").Value2 = "expected"
    Set outcome = Evaluate(p): Set plan = AtPlanActions(book, AtOutcomeIntents(p, outcome), "MOVE_PASS", journal)
    Check outcome.Status = "PASS" And plan.Status = "READY", prefix & " moved field reconciles previous physical target"
    Set receipts = AtApplyPlan(plan, journal)
    Check receipts(1).Status = "ROLLED_BACK" And cell.Value2 = "edited old location" And cell.Interior.Color = RGB(70, 80, 90), prefix & " moving field clears old mark without restoring old input"
    Check cell.Comment Is Nothing And book.Worksheets(1).Range("D5").Value2 = "expected", prefix & " move preserves both inputs and removes old Note"
    Set cell = book.Worksheets(1).Range("D5"): cell.Value2 = "wrong": Set outcome = Evaluate(p)
    Set receipts = AtApplyPlan(AtPlanActions(book, AtOutcomeIntents(p, outcome), "MARK_NEW_LOCATION", journal), journal)
    Check receipts(1).Status = "APPLIED", prefix & " next FAIL marks moved field"
    cell.Value2 = "expected": Set outcome = Evaluate(p)
    Set receipts = AtApplyPlan(AtPlanActions(book, AtOutcomeIntents(p, outcome), "PASS_NEW_LOCATION", journal), journal)
    Check receipts(1).Status = "ROLLED_BACK" And cell.Value2 = "expected" And cell.Interior.Pattern = -4142, prefix & " moved field PASS clears mark and preserves input"
    book.Close False: Set book = Nothing
End Sub
