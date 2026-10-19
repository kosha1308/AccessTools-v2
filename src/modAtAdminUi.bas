Attribute VB_Name = "modAtAdminUi"
Option Explicit

' Native Excel dialogs, deliberately limited to one editable scalar condition.
' State survives closing dialogs, but never survives loss of the VBA session.
Private configuration As CAtPackage, targetBook As Object
Private journal As CAtOwnershipJournal, journalBook As Object
Private runNumber As Long

Public Sub AtAdminOpen()
    On Error GoTo Failed
    Dim command As String
    Do
        If Not Ask("AccessTools editor: NEW / EDIT / LOAD / SAVE / RUN / CLEAR / CLOSE" & vbLf & _
            "Text equality; Notes require the tested .xls profile. .xlsx: fill only." & vbLf & _
            "Closing retains ownership only in this live VBA session.", "CLOSE", command) Then Exit Sub
        Select Case UCase$(command)
            Case "NEW": EditDefinition True
            Case "EDIT": EditDefinition False
            Case "LOAD": LoadDefinition
            Case "SAVE": SaveDefinition
            Case "RUN": RunDefinition
            Case "CLEAR": ClearDefinition
            Case "CLOSE": Exit Sub
            Case Else: MsgBox "Choose a command from the list."
        End Select
    Loop
    Exit Sub
Failed:
    MsgBox "Action stopped: " & Err.Description, vbExclamation, "AccessTools"
End Sub

Private Function Ask(ByVal prompt As String, ByVal initial As String, ByRef answer As String) As Boolean
    Dim value As Variant
    value = Application.InputBox(prompt, "AccessTools", initial, Type:=2)
    If VarType(value) = vbBoolean Then Exit Function
    answer = CStr(value): Ask = True
End Function

Private Function PickBook() As Object
    Dim books As New Collection, book As Object, text As String, choice As String, i As Long
    For Each book In Application.Workbooks
        If Not book Is ThisWorkbook Then books.Add book: text = text & books.Count & ": " & book.Name & vbLf
    Next book
    If books.Count = 0 Then Err.Raise 5, , "Open a target workbook first (separate from the macro host)."
    If Len(text) > 200 Then Err.Raise 5, , "Too many open workbooks for this small dialog. Close unrelated workbooks."
    If Not Ask("Choose an open workbook by number:" & vbLf & text, "1", choice) Then Exit Function
    If Not IsNumeric(choice) Then Err.Raise 5, , "Invalid workbook number."
    i = CLng(choice)
    If CStr(i) <> choice Or i < 1 Or i > books.Count Then Err.Raise 5, , "Invalid workbook number."
    Set PickBook = books(i)
End Function

Private Sub EditDefinition(ByVal createNew As Boolean)
    Dim book As Object, sheet As Object, selected As Object, sheetName As String
    Dim fieldName As String, expected As String, message As String, severity As String, color As String, notes As String
    Dim d As CAtDefinition, expression As CAtDefinition, n As CAtNode, candidate As CAtPackage, issues As Collection, registry As New CAtOperationRegistry
    If Not createNew Then
        If configuration Is Nothing Then Err.Raise 5, , "Create or load a configuration first."
        CheckEditorProfile configuration
    End If
    Set book = PickBook(): If book Is Nothing Then Exit Sub
    If Not Ask("Worksheet name in " & book.Name, book.Worksheets(1).Name, sheetName) Then Exit Sub
    Set sheet = book.Worksheets(sheetName)
    If sheet.Visible <> -1 Then Err.Raise 5, , "Choose a visible worksheet."
    book.Activate: sheet.Activate
    ' Type 8 returns a real Range; cancellation raises 424, handled only here.
    On Error Resume Next
    Set selected = Application.InputBox("Select ONE field cell on the chosen worksheet", "AccessTools", Type:=8)
    Dim pickError As Long: pickError = Err.Number: Err.Clear
    On Error GoTo 0
    If selected Is Nothing Then
        If pickError <> 0 And pickError <> 424 Then Err.Raise pickError
        Exit Sub
    End If
    If Not selected.Parent Is sheet Or selected.Cells.CountLarge <> 1 Then Err.Raise 5, , "Select one cell on the chosen worksheet."
    fieldName = "Text field": expected = "expected": message = "Value must equal expected text."
    severity = "ERROR": color = "255": notes = "false"
    If Not createNew Then
        fieldName = AtDefinitionById(configuration, "Field", "FIELD_VALUE").DisplayName
        Set expression = AtDefinitionById(configuration, "Expression", "EXPRESSION")
        For Each n In expression.Nodes
            If n.Id = "EXPECTED" Then expected = CStr(n.Values("Value"))
        Next n
        Set d = AtDefinitionById(configuration, "Action", "ACTION_FAIL")
        message = AtValue(d, "Message"): severity = UCase$(AtValue(d, "Trigger"))
        color = AtValue(d, "FillColor"): notes = AtValue(d, "AddNote")
    End If
    If Not Ask("Field display name (type is Text)", fieldName, fieldName) Then Exit Sub
    If Not Ask("Condition: text equals (case sensitive, exact)", expected, expected) Then Exit Sub
    If Not Ask("Message on FAIL (short text; Writer heading included in validation)", message, message) Then Exit Sub
    If Not Ask("Severity: ERROR / WARNING / INFO", severity, severity) Then Exit Sub
    If Not Ask("Fill color: Excel RGB integer 0-16777215; red=255, yellow=65535", color, color) Then Exit Sub
    If Not AtLiteralValid("Int64", color) Then Err.Raise 5, , "Invalid RGB integer."
    If CDec(color) < 0 Or CDec(color) > 16777215 Then Err.Raise 5, , "RGB outside range."
    If Not Ask("Add Note: true / false. Unsupported .xlsx Notes reject the WHOLE action.", notes, notes) Then Exit Sub
    If notes <> "true" And notes <> "false" Then Err.Raise 5, , "Choose true or false explicitly."
    Set candidate = AtAdminDefinition(sheet.Name, selected.Address(False, False), fieldName, expected, message, severity, CLng(color), notes = "true")
    Set issues = AtValidate(candidate, registry.Catalog)
    If AtHasErrors(issues) Then Err.Raise 5, , DiagnosticsText(issues)
    If Not createNew Then
        candidate.Id = configuration.Id: candidate.Version = configuration.Version
        candidate.DisplayName = configuration.DisplayName
    End If
    Set configuration = candidate: Set targetBook = book
    MsgBox "Definition updated in memory. SAVE writes a NEW revision file."
End Sub

Private Sub CheckEditorProfile(ByVal p As CAtPackage)
    ' Reject unsupported configurations rather than silently discarding their AST.
    Dim d As CAtDefinition, expected As CAtPackage, expression As CAtDefinition, n As CAtNode, value As String
    If p.SchemaVersion <> "3.0.0" Or p.Definitions.Count <> 6 Then Err.Raise 5, , "This editor opens only its single Text equality profile."
    Set expression = AtDefinitionById(p, "Expression", "EXPRESSION")
    For Each n In expression.Nodes
        If n.Id = "EXPECTED" And n.Values.Exists("Value") Then value = CStr(n.Values("Value"))
    Next n
    Set d = AtDefinitionById(p, "Action", "ACTION_FAIL")
    Dim locator As CAtDefinition: Set locator = AtDefinitionById(p, "Locator", "LOCATION")
    Set expected = AtAdminDefinition(AtValue(locator, "Sheet"), AtValue(locator, "Address"), AtDefinitionById(p, "Field", "FIELD_VALUE").DisplayName, value, AtValue(d, "Message"), UCase$(AtValue(d, "Trigger")), CLng(AtValue(d, "FillColor")), AtValue(d, "AddNote") = "true")
    expected.Id = p.Id: expected.Version = p.Version: expected.DisplayName = p.DisplayName
    If AtToXml(expected) <> AtToXml(p) Then Err.Raise 5, , "Configuration exceeds the supported editor profile; no changes made."
End Sub

Private Sub LoadDefinition()
    Dim path As Variant, candidate As CAtPackage, issues As Collection, registry As New CAtOperationRegistry
    path = Application.GetOpenFilename("AccessTools XML (*.xml),*.xml")
    If VarType(path) = vbBoolean Then Exit Sub
    Set candidate = AtLoad(CStr(path), registry.Catalog, issues)
    If candidate Is Nothing Then Err.Raise 5, , DiagnosticsText(issues)
    If AtHasErrors(issues) Then Err.Raise 5, , DiagnosticsText(issues)
    CheckEditorProfile candidate
    Set configuration = candidate: Set targetBook = Nothing
    MsgBox "Loaded. RUN will ask for the target workbook explicitly."
End Sub

Private Sub SaveDefinition()
    Dim path As Variant, issues As Collection, registry As New CAtOperationRegistry
    If configuration Is Nothing Then Err.Raise 5, , "Create or load first."
    path = Application.GetSaveAsFilename("AccessTools-rule.xml", "AccessTools XML (*.xml),*.xml")
    If VarType(path) = vbBoolean Then Exit Sub
    If Not AtSaveNew(configuration, CStr(path), registry.Catalog, issues) Then Err.Raise 5, , DiagnosticsText(issues)
    MsgBox "Saved and round-trip verified: " & CStr(path)
End Sub

Private Sub EnsureJournal(ByVal book As Object)
    If Not journal Is Nothing Then
        If journalBook Is book Then Exit Sub
        Err.Raise 5, , "This session is bound to another workbook. Clear its owned marks before switching."
    End If
    Dim folder As Object, fso As Object, path As String
    Set folder = Application.FileDialog(4)
    folder.Title = "Choose SAFE LOCAL parent directory for a NEW ownership journal"
    If folder.Show <> -1 Then Err.Raise 5, , "Journal selection cancelled."
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(folder.SelectedItems(1), "AccessTools-" & fso.GetTempName)
    fso.CreateFolder path
    Set journal = AtNewOwnershipJournal(book, path): Set journalBook = book
    MsgBox "Live-session journal: " & path & vbLf & "Never automatically reused after session loss."
End Sub

Private Sub RunDefinition()
    Dim book As Object, resolver As CAtExcelResolver, context As CAtEvalContext, outcome As CAtRuleOutcome
    Dim issues As Collection, intents As Collection, plan As CAtActionPlan, receipts As Collection, receipt As CAtActionReceipt, text As String
    If configuration Is Nothing Then Err.Raise 5, , "Create or load first."
    Set book = PickBook(): If book Is Nothing Then Exit Sub
    Set resolver = AtNewExcelResolver(configuration, "TEMPLATE", AtDictionary())
    Set context = resolver.Resolve(book, issues)
    If context Is Nothing Then Err.Raise 5, , DiagnosticsText(issues)
    Set outcome = AtEvaluateRule(configuration, "RULE_EQUAL", context)
    If outcome.Status = "CANNOT_EVALUATE" Then
        MsgBox outcome.Status & " / " & outcome.Result.State & vbLf & DiagnosticsText(outcome.Result.Diagnostics) & vbLf & "Existing marks retained.", vbExclamation
        Exit Sub
    End If
    EnsureJournal book
    Set intents = AtOutcomeIntents(configuration, outcome)
    runNumber = runNumber + 1
    Set plan = AtPlanActions(book, intents, "RUN_" & runNumber, journal)
    text = "Condition: " & outcome.Status & vbLf & "Output plan: " & plan.Status & " " & plan.Code & " " & plan.Message
    If plan.Status = "READY" Then
        Set receipts = AtApplyPlan(plan, journal)
        For Each receipt In receipts
            text = text & vbLf & receipt.Status & " " & receipt.Code & " " & receipt.Message
        Next receipt
    End If
    If outcome.Status = "FAIL" Then text = text & vbLf & AtValue(AtDefinitionById(configuration, "Action", "ACTION_FAIL"), "Message")
    MsgBox text, vbInformation, "AccessTools results"
End Sub

Private Sub ClearDefinition()
    Dim receipts As Collection, receipt As CAtActionReceipt, text As String, complete As Boolean
    If journal Is Nothing Then Err.Raise 5, , "No owned marks in this session."
    runNumber = runNumber + 1: Set receipts = AtClearOwned(journalBook, journal, "CLEAR_" & runNumber)
    complete = True
    For Each receipt In receipts
        text = text & receipt.Status & " " & receipt.Code & vbLf
        If receipt.Status <> "ROLLED_BACK" And receipt.Status <> "SKIPPED" Then complete = False
    Next receipt
    MsgBox text
    If complete And Not journal.RecoveryRequired Then
        Set journal = Nothing: Set journalBook = Nothing
    End If
End Sub

Private Function DiagnosticsText(ByVal issues As Collection) As String
    Dim issue As Variant
    For Each issue In issues
        DiagnosticsText = DiagnosticsText & issue("Code") & ": " & issue("Message") & vbLf
    Next issue
End Function
