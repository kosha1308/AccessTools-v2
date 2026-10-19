Attribute VB_Name = "AtWriterTests"
Option Explicit

Private passed As Long, failed As Long, skipped As Long, transcript As String, counter As Long
Private book As Object, otherBook As Object, folder As String, root As String

Public Function AtRunWriterTests(ByVal repo As String, ByVal scratch As String) As String
    On Error GoTo Broken
    Dim journal As Object
    Set journal = Application.Run("'" & ThisWorkbook.Name & "'!AtNewOwnershipJournal", ThisWorkbook, scratch)
    passed = 0: failed = 0: skipped = 0: transcript = "": counter = 0: root = scratch
    Set book = Application.Workbooks.Add: book.Worksheets(1).Name = "Input"
    Set otherBook = Application.Workbooks.Add: otherBook.Worksheets(1).Name = "Independent"
    Check AtWUnhex(AtWHex(ChrW(-10179) & ChrW(-8638) & vbCrLf & ChrW(-1))) = ChrW(-10179) & ChrW(-8638) & vbCrLf & ChrW(-1), "journal UTF-16 encoding round-trips surrogate pairs CRLF and high code units"
    TestFill
    TestNotesAndMerges
    TestCapabilities
    TestConflicts
    TestRecovery
    Cleanup
    Dim file As Object
    Set file = CreateObject("Scripting.FileSystemObject").CreateTextFile(scratch & "\writer-assertions.txt", True, True)
    file.Write transcript: file.Close
    AtRunWriterTests = IIf(failed = 0, "PASS|Writer|", "FAIL|Writer|") & passed & "|" & failed & "|UNVERIFIED|" & skipped
    If failed > 0 Then AtRunWriterTests = AtRunWriterTests & "|" & transcript
    Exit Function
Broken:
    AtRunWriterTests = "FAIL|Writer|" & passed & "|" & Err.Number & "|" & Err.Description & "|" & transcript
    Cleanup
End Function

Private Sub Cleanup()
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    If Not otherBook Is Nothing Then otherBook.Close False
    Set book = Nothing: Set otherBook = Nothing
End Sub
Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    If ok Then passed = passed + 1 Else failed = failed + 1
    transcript = transcript & IIf(ok, "PASS ", "FAIL ") & label & vbCrLf
End Sub
Private Function Fresh(Optional ByVal workbook As Object = Nothing, Optional ByVal revisions As Long = 256) As CAtOwnershipJournal
    counter = counter + 1: folder = root & "\writer_" & counter
    CreateObject("Scripting.FileSystemObject").CreateFolder folder
    If workbook Is Nothing Then Set workbook = book
    Set Fresh = AtNewOwnershipJournal(workbook, folder, revisions)
End Function
Private Function Intent(ByVal key As String, ByVal address As String, Optional ByVal note As Boolean = False, Optional ByVal severity As String = "ERROR") As CAtActionIntent
    Dim i As New CAtActionIntent
    i.ActionInstanceKey = key: i.Sheet = "Input": i.Address = address: i.AddNote = note
    i.Severity = severity: i.Message = key: i.FillColor = RGB(250, 30, 40)
    Set Intent = i
End Function
Private Function Plan(ByVal journal As CAtOwnershipJournal, ByVal actions As Collection, Optional ByVal run As String = "RUN_TEST") As CAtActionPlan
    Set Plan = AtPlanActions(book, actions, run, journal)
End Function
Private Function Capture(ByVal address As String) As Object
    Dim code As String
    Set Capture = AtWCapture(book.Worksheets(1).Range(address), code)
End Function

Private Sub TestFill()
    Dim journal As CAtOwnershipJournal, actions As New Collection, planValue As CAtActionPlan, receipts As Collection, before As Object
    Dim cell As Object, active As Object, events As Boolean, screen As Boolean, alerts As Boolean, calculation As Long, count As Long
    Set cell = book.Worksheets(1).Range("A1"): cell.Formula = "=1+2": cell.Font.Bold = True
    cell.NumberFormat = "0.00": cell.Borders(9).LineStyle = 1
    book.SaveAs root & "\writer-synthetic.xls", 56
    Set before = Capture("A1"): Set journal = Fresh()
    actions.Add Intent("FILL", "A1")
    Set active = Application.ActiveWorkbook: count = Application.Workbooks.Count
    events = Application.EnableEvents: screen = Application.ScreenUpdating: alerts = Application.DisplayAlerts: calculation = Application.Calculation
    Set planValue = Plan(journal, actions)
    Check planValue.Status = "READY" And book.Saved, "planning is read-only on a saved workbook: " & planValue.Status & "/" & planValue.Code & "/" & planValue.Message
    If planValue.Status <> "READY" Then Err.Raise 5, , "Preflight failed: " & planValue.Code & " " & planValue.Message
    planValue.Actions(1)("Color") = "0"
    Set receipts = AtApplyPlan(planValue, journal)
    Check receipts(1).Status = "APPLIED" And cell.Interior.Color = RGB(250, 30, 40), "direct fill applied and plan getter is detached: " & receipts(1).Status & " " & receipts(1).Message
    If receipts(1).Status <> "APPLIED" Then
        Dim expected As Object, actual As Object, key As Variant
        Set expected = AtWUnpack(journal.Entries(1)("Next")): Set actual = Capture("A1")
        If Not actual Is Nothing Then
            For Each key In expected.Keys
                If actual(key) <> expected(key) Then transcript = transcript & key & " expected=" & expected(key) & " actual=" & actual(key) & vbCrLf
            Next key
        End If
        Err.Raise 5, , "Fill verification failed"
    End If
    Check journal.RevisionCount >= 3 And journal.Verify(), "write-ahead and actual-state acknowledgements persisted and verified"
    Check cell.Formula = "=1+2" And cell.Value2 = 3 And cell.Font.Bold And cell.NumberFormat = "0.00" And cell.Borders(9).LineStyle = 1, "writer preserves data formula font borders and number format"
    Check Not book.Saved, "writer does not auto-save or forge Saved=True"
    Check Application.ActiveWorkbook Is active And Application.Workbooks.Count = count, "explicit workbook writer does not activate or open books"
    Check Application.EnableEvents = events And Application.ScreenUpdating = screen And Application.DisplayAlerts = alerts And Application.Calculation = calculation, "writer leaves application flags unchanged"
    Set receipts = AtApplyPlan(Plan(journal, actions), journal)
    Check receipts(1).Status = "SKIPPED", "same action retry is idempotent"
    Set receipts = AtClearOwned(book, journal, "RUN_CLEAR")
    Check receipts(1).Status = "ROLLED_BACK" And AtWEqual(before, Capture("A1")), "no-fill baseline restored exactly"
    Set receipts = AtClearOwned(book, journal, "RUN_CLEAR")
    Check receipts.Count = 0, "repeated cleanup has no extra mutations"
    cell.Interior.Color = RGB(12, 34, 56): Set before = Capture("A1")
    Set receipts = AtApplyPlan(Plan(journal, actions), journal): Set receipts = AtClearOwned(book, journal, "RUN_CLEAR")
    Check receipts(1).Status = "ROLLED_BACK" And AtWEqual(before, Capture("A1")), "custom RGB baseline restored exactly"
End Sub

Private Sub TestNotesAndMerges()
    Dim journal As CAtOwnershipJournal, actions As New Collection, i As CAtActionIntent, receipts As Collection, before As Object, cell As Object, planValue As CAtActionPlan
    Set cell = book.Worksheets(1).Range("B5:C6"): cell.Merge: cell.Interior.Color = RGB(40, 50, 60)
    Set before = Capture("B5:C6"): Set journal = Fresh()
    Set i = Intent("LOW", "C6", True, "INFO"): i.FillColor = RGB(20, 180, 20): i.Message = "same message": actions.Add i
    Set i = Intent("HIGH", "B5", True, "ERROR"): i.Message = "same message " & ChrW(&H416): actions.Add i
    Set i = Intent("MID", "C5", True, "WARNING"): i.FillColor = RGB(200, 200, 0): actions.Add i
    Set planValue = Plan(journal, actions)
    Check planValue.Status = "READY" And planValue.Actions.Count = 1 And planValue.Actions(1)("Address") = "$B$5:$C$6", "merge normalization precedes grouping"
    Check planValue.Actions(1)("Messages") = "AccessTools" & vbLf & "[INFO] same message" & vbLf & "[ERROR] same message " & ChrW(&H416) & vbLf & "[WARNING] MID", "all messages retained in input order without text deduplication"
    Set receipts = AtApplyPlan(planValue, journal)
    Check receipts(1).Status = "APPLIED", "new legacy note and fill successfully verified: " & receipts(1).Message
    Check cell.Interior.Color = RGB(250, 30, 40), "ERROR fill has priority over WARNING and INFO"
    Check cell.Cells(1, 1).Comment.Text = planValue.Actions(1)("Messages"), "Unicode legacy note contains every grouped message"
    Set receipts = AtApplyPlan(Plan(journal, actions), journal)
    Check receipts(1).Status = "SKIPPED" And book.Worksheets(1).Comments.Count = 1, "repeat creates no duplicate note"
    Set i = actions(2): i.Message = "replacement"
    Set receipts = AtApplyPlan(Plan(journal, actions, "RUN_NEW"), journal)
    Check receipts(1).Status = "APPLIED" And InStr(cell.Cells(1, 1).Comment.Text, "replacement") > 0, "owned note replaced with a new run while preserving original baseline"
    Set receipts = AtClearOwned(book, journal, "RUN_CLEAR")
    Check receipts(1).Status = "ROLLED_BACK" And AtWEqual(before, Capture("B5:C6")), "merged fill restored and only own note removed"
    Check cell.Cells(1, 1).MergeArea.Address = "$B$5:$C$6", "merge structure remains unchanged"
End Sub

Private Sub TestCapabilities()
    Dim journal As CAtOwnershipJournal, actions As Collection, cell As Object, planValue As CAtActionPlan, receipts As Collection
    Dim first As Object, second As Object, count As Long, priority As Long, thread As Object, number As Long
    Set cell = book.Worksheets(1).Range("D1"): cell.Value2 = "user": cell.Interior.Color = RGB(250, 30, 40)
    cell.AddComment "AccessTools is a user note": cell.Comment.Shape.TextFrame.Characters(1, 5).Font.Bold = True
    Set journal = Fresh(): Set actions = New Collection: actions.Add Intent("USER_NOTE", "D1", True)
    Set planValue = Plan(journal, actions)
    Check planValue.Status = "FAILED" And planValue.Code = "CAPABILITY_UNSUPPORTED", "existing rich user note refuses whole action before fill"
    Set receipts = AtClearOwned(book, journal, "RUN_CLEAR")
    Check receipts.Count = 0 And cell.Comment.Text = "AccessTools is a user note" And cell.Comment.Shape.TextFrame.Characters(1, 5).Font.Bold, "color and AccessTools text do not confer ownership; user note formatting retained"
    cell.Comment.Delete
    Set first = cell.FormatConditions.Add(1, 3, "=1"): first.Interior.Color = RGB(1, 2, 3)
    Set second = cell.FormatConditions.Add(1, 3, "=2")
    count = cell.FormatConditions.Count: priority = first.Priority
    Set planValue = Plan(journal, actions)
    Check planValue.Code = "CAPABILITY_UNSUPPORTED" And cell.FormatConditions.Count = count And first.Priority = priority And first.Formula1 = "=1", "conditional formatting and its priorities are preserved by refusal"
    cell.FormatConditions.Delete
    book.Worksheets(1).Protect "synthetic"
    Check Plan(journal, actions).Code = "CAPABILITY_UNSUPPORTED", "protected sheet refused without unprotecting"
    book.Worksheets(1).Unprotect "synthetic"
    cell.Interior.ThemeColor = 4
    Check Plan(journal, actions).Code = "CAPABILITY_UNSUPPORTED", "theme fill is capability-gated"
    cell.Interior.Color = RGB(10, 20, 30): cell.Interior.Pattern = 5
    Check Plan(journal, actions).Code = "CAPABILITY_UNSUPPORTED", "patterned fill is capability-gated"
    cell.Interior.Pattern = -4142
    On Error Resume Next
    Set thread = cell.AddCommentThreaded("Synthetic threaded user comment")
    number = Err.Number: Err.Clear
    On Error GoTo 0
    If thread Is Nothing Then
        skipped = skipped + 1: transcript = transcript & "UNVERIFIED threaded creation unsupported: " & number & vbCrLf
    Else
        Check Plan(journal, actions).Code = "CAPABILITY_UNSUPPORTED" And cell.CommentThreaded.Text = "Synthetic threaded user comment", "threaded user comments refused and retained"
        thread.Delete
    End If
    Set cell = book.Worksheets(1).Range("E5:F6"): cell.Merge: cell.Interior.Color = RGB(1, 2, 3)
    book.Worksheets(1).Range("F6").Interior.Color = RGB(4, 5, 6)
    Set actions = New Collection: actions.Add Intent("MIXED_MERGE", "E5")
    Check Plan(journal, actions).Code = "CAPABILITY_UNSUPPORTED", "nonuniform merged fills refused instead of flattening original formatting"
    Check book.Worksheets(1).Range("F6").Interior.Color = RGB(4, 5, 6), "refusal preserves non-anchor fill"
    Set actions = New Collection: actions.Add Intent("RECTANGLE", "H1:I1")
    Check Plan(journal, actions).Code = "CAPABILITY_UNSUPPORTED", "ordinary multicell rectangles are outside bounded capability"
    Set actions = New Collection: actions.Add Intent("BAD_ADDRESS", "XFE1")
    Check Plan(journal, actions).Code = "PLAN_INVALID", "invalid Excel address rejected before write"
    Set actions = New Collection: actions.Add Intent("MISSING", "A1"): actions(1).Sheet = "Absent"
    Check Plan(journal, actions).Status = "FAILED", "missing worksheet rejected before write"
    Set actions = New Collection: actions.Add Intent("TWO_NOTES_A", "H1", True): actions.Add Intent("TWO_NOTES_B", "H2", True)
    Check Plan(journal, actions).Code = "CAPABILITY_UNSUPPORTED", "multiple note areas on one sheet refused before mutation to preserve shape ordering"
    Set cell = book.Worksheets(1).Range("H3"): cell.Value2 = String$(5000, "x")
    Set actions = New Collection: actions.Add Intent("LARGE_FILL", "H3"): actions.Add Intent("LARGE_NOTE", "H3", True)
    Set planValue = Plan(journal, actions)
    Check planValue.Code = "CAPABILITY_UNSUPPORTED" And cell.Comment Is Nothing And cell.Interior.Pattern = -4142, "grouped note budget checked even when first intent is fill-only"
    cell.ClearContents
    Set journal = Fresh(otherBook): Set actions = New Collection: actions.Add Intent("INDEPENDENT", "A1")
    actions(1).Sheet = "Independent"
    Set planValue = AtPlanActions(otherBook, actions, "RUN_SECOND", journal)
    Check planValue.Status = "READY" And otherBook.Path = "", "same writer plans a different unsaved template without a handler"
    Set receipts = AtApplyPlan(planValue, journal)
    Check receipts(1).Status = "APPLIED", "unsaved modern-format workbook supports safe direct fill"
    otherBook.SaveAs root & "\independent-writer.xlsx", 51
    Set receipts = AtClearOwned(otherBook, journal, "RUN_CLEAR")
    Check receipts(1).Status = "ROLLED_BACK", "live workbook identity survives user SaveAs without rebind or auto-save"
    actions(1).AddNote = True
    Set planValue = AtPlanActions(otherBook, actions, "RUN_SECOND", journal)
    If planValue.Status <> "READY" Then
        Check planValue.Code = "CAPABILITY_UNSUPPORTED", "modern note action fails closed when threaded inspection is unavailable"
    Else
        Set receipts = AtApplyPlan(planValue, journal)
        Check receipts(1).Status = "APPLIED", "modern host note capability works when inspection is available"
        Set receipts = AtClearOwned(otherBook, journal, "RUN_CLEAR")
    End If
End Sub

Private Sub TestConflicts()
    Dim journal As CAtOwnershipJournal, actions As Collection, planValue As CAtActionPlan, receipts As Collection, cell As Object, width As Double
    Dim font As Object, fontIndex As Long, fontRgb As Long, number As Long, fillRgb As Long
    Set cell = book.Worksheets(1).Range("A2"): cell.Formula = "=3+4"
    Set journal = Fresh(): Set actions = New Collection: actions.Add Intent("FORMULA", "A2")
    Set planValue = Plan(journal, actions): cell.Formula = "=100"
    Set receipts = AtApplyPlan(planValue, journal)
    Check receipts(1).Status = "CONFLICT" And journal.RevisionCount = 0 And cell.Formula = "=100", "stale plan rejects user formula change before any write"
    cell.Formula = "=3+4": Set receipts = AtApplyPlan(Plan(journal, actions), journal)
    cell.Interior.Color = RGB(55, 66, 77)
    Set receipts = AtClearOwned(book, journal, "RUN_CLEAR")
    Check receipts(1).Status = "CONFLICT" And cell.Interior.Color = RGB(55, 66, 77), "cleanup does not erase user's replacement fill"
    cell.Interior.Color = RGB(250, 30, 40): Set receipts = AtClearOwned(book, journal, "RUN_CLEAR")
    Check receipts(1).Status = "ROLLED_BACK", "cleanup succeeds after exact expected state is restored"
    Set actions = New Collection: actions.Add Intent("NOTE_STYLE", "A3", True)
    Set receipts = AtApplyPlan(Plan(journal, actions), journal): Set cell = book.Worksheets(1).Range("A3")
    width = cell.Comment.Shape.Width: cell.Comment.Shape.Width = width + 10
    Check Plan(journal, actions).Status = "CONFLICT", "owned note shape changes cause conflict"
    cell.Comment.Shape.Width = width
    Set font = cell.Comment.Shape.TextFrame.Characters.Font
    fontIndex = font.ColorIndex: fontRgb = font.Color
    font.Color = fontRgb
    Check Plan(journal, actions).Status = "CONFLICT", "same-RGB font palette linkage change detected"
    If fontIndex > 56 Then font.ColorIndex = -4105 Else font.ColorIndex = fontIndex
    On Error Resume Next
    font.ThemeColor = 1
    number = Err.Number: Err.Clear
    On Error GoTo 0
    If number = 0 Then
        Check Plan(journal, actions).Status = "CONFLICT", "note font theme linkage change detected independently of RGB"
        If fontIndex > 56 Then font.ColorIndex = -4105 Else font.ColorIndex = fontIndex
    Else
        skipped = skipped + 1: transcript = transcript & "UNVERIFIED note ThemeColor setter unavailable: " & number & vbCrLf
    End If
    fillRgb = cell.Comment.Shape.Fill.ForeColor.RGB
    cell.Comment.Shape.Fill.ForeColor.RGB = fillRgb
    Check Plan(journal, actions).Status = "CONFLICT", "same-RGB conversion of system-linked note fill to direct RGB is detected"
    cell.Comment.Shape.Fill.ForeColor.SchemeColor = 80
    cell.Comment.Shape.Shadow.Visible = -1
    Check Plan(journal, actions).Status = "CONFLICT", "unsupported shadow added to owned note prevents deletion"
    cell.Comment.Shape.Shadow.Visible = 0
    cell.Comment.Shape.TextFrame.Characters(1, 1).Font.Bold = True
    Check Plan(journal, actions).Status = "CONFLICT", "rich font edit in owned note causes conflict"
    cell.Comment.Shape.TextFrame.Characters(1, 1).Font.Bold = False
    Set receipts = AtClearOwned(book, journal, "RUN_CLEAR")
    Check receipts(1).Status = "ROLLED_BACK", "untouched supported note can be cleared after user undo"
    Set actions = New Collection: actions.Add Intent("OLD", "A4")
    Set receipts = AtApplyPlan(Plan(journal, actions), journal)
    Set actions = New Collection: actions.Add Intent("NEW", "A5")
    Set receipts = AtApplyPlan(Plan(journal, actions, "RUN_RECHECK"), journal)
    Check receipts.Count = 2 And book.Worksheets(1).Range("A4").Interior.Pattern = -4142 And book.Worksheets(1).Range("A5").Interior.Pattern = 1, "new plan reconciles obsolete owned regions using original journal targets"
    Set receipts = AtClearOwned(book, journal, "RUN_CLEAR")
End Sub

Private Sub TestRecovery()
    Dim journal As CAtOwnershipJournal, orphan As CAtOwnershipJournal, actions As Collection, receipts As Collection, before As Object
    Dim cell As Object, entry As Object, fso As Object, file As Object, savedFolder As String
    Set cell = book.Worksheets(1).Range("A10"): Set before = Capture("A10")
    Set journal = Fresh(): Set actions = New Collection: actions.Add Intent("INTERRUPT", "A10", True)
    Set receipts = AtApplyPlan(Plan(journal, actions), journal, 1)
    Check receipts(1).Status = "RECOVERY_REQUIRED" And journal.RecoveryRequired, "interruption between fill and note is not successful output"
    Check cell.Interior.Pattern = 1 And cell.Comment Is Nothing And journal.RevisionCount >= 2, "partial fill has acknowledged write-ahead evidence"
    Check Plan(journal, actions).Status = "RECOVERY_REQUIRED", "pending recovery blocks new annotation runs"
    Set receipts = AtRecoverOwned(book, journal, "RUN_RECOVERY")
    Check receipts(1).Status = "ROLLED_BACK" And AtWEqual(before, Capture("A10")) And Not journal.RecoveryRequired, "same-session recovery restores a verified partial fill"
    Set journal = Fresh(): savedFolder = folder
    Set receipts = AtApplyPlan(Plan(journal, actions), journal)
    Set orphan = AtNewOwnershipJournal(book, savedFolder)
    Check orphan.RecoveryRequired And orphan.EntryCount = 0, "new session cannot adopt durable files as ownership proof"
    Set receipts = AtRecoverOwned(book, orphan, "RUN_RECOVERY")
    Check receipts(1).Status = "RECOVERY_REQUIRED" And Not cell.Comment Is Nothing, "orphan journal requires manual recovery without deleting notes"
    Set entry = journal.Entries(1)
    entry("State") = "PENDING": entry("Phase") = "ADD_NOTE": entry("Expected") = entry("Before"): entry("Next") = ""
    journal.Persist entry: journal.RequireRecovery
    Set receipts = AtRecoverOwned(book, journal, "RUN_RECOVERY")
    Check receipts(1).Status = "CONFLICT" And journal.RecoveryRequired And Not cell.Comment Is Nothing, "unacknowledged note is not adopted by matching AccessTools text"
    Set journal = Fresh(): Set actions = New Collection: actions.Add Intent("COLLISION", "A11")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(folder & "\ATJ_" & journal.SessionId & "_0001.atj", False, True): file.Write "existing": file.Close
    Set receipts = AtApplyPlan(Plan(journal, actions), journal)
    Check receipts(1).Status = "RECOVERY_REQUIRED" And book.Worksheets(1).Range("A11").Interior.Pattern = -4142, "write-ahead file collision fails before Excel mutation"
    Set journal = Fresh(): Set actions = New Collection: actions.Add Intent("TAMPER", "A12")
    Set receipts = AtApplyPlan(Plan(journal, actions), journal)
    For Each file In fso.GetFolder(folder).Files
        Set file = fso.OpenTextFile(file.Path, 8, False, -1): file.Write "corrupt": file.Close: Exit For
    Next file
    Set receipts = AtClearOwned(book, journal, "RUN_CLEAR")
    Check receipts(1).Status = "RECOVERY_REQUIRED" And book.Worksheets(1).Range("A12").Interior.Pattern = 1, "modified durable journal blocks cleanup without erasing marks"
    Set journal = Fresh(): savedFolder = folder
    Set orphan = AtNewOwnershipJournal(book, savedFolder)
    Check orphan.RecoveryRequired, "exclusive journal lease blocks a second owner of an empty session directory"
    Set file = fso.CreateTextFile(folder & "\unknown.atj", False, True): file.Write "orphan": file.Close
    Set actions = New Collection: actions.Add Intent("UNKNOWN_FILE", "A13")
    Check Plan(journal, actions).Status = "RECOVERY_REQUIRED", "new unknown journal file blocks an already initialized session"
    Set journal = Fresh(Nothing, 1): Set before = Capture("A14")
    Set actions = New Collection: actions.Add Intent("RESERVE", "A14")
    Set receipts = AtApplyPlan(Plan(journal, actions), journal)
    Check receipts(1).Status = "RECOVERY_REQUIRED" And journal.RevisionCount = 1, "exhausted normal revision budget does not falsely acknowledge mutation"
    Set receipts = AtRecoverOwned(book, journal, "RUN_RECOVERY")
    Check receipts(1).Status = "ROLLED_BACK" And journal.RevisionCount > 1 And AtWEqual(before, Capture("A14")), "reserved journal capacity permits recovery after normal budget exhaustion"
    Set journal = Fresh(): Set actions = New Collection
    actions.Add Intent("STEP_ONE", "A15"): actions.Add Intent("STEP_TWO", "A16"): actions.Add Intent("STEP_THREE", "A17")
    Set receipts = AtApplyPlan(Plan(journal, actions), journal, 1)
    Check receipts.Count = 3 And receipts(1).Status = "APPLIED" And receipts(2).Status = "RECOVERY_REQUIRED" And receipts(3).Code = "BLOCKED_RECOVERY", "mid-plan failure preserves actual receipts and explicitly blocks remaining actions"
    Check receipts(2).ActionInstanceKey <> "" And receipts(3).Target <> "", "interrupted and blocked receipts retain action identity and target"
    Set receipts = AtRecoverOwned(book, journal, "RUN_RECOVERY")
    Check receipts(1).Status = "ROLLED_BACK" And book.Worksheets(1).Range("A15").Interior.Pattern = -4142 And book.Worksheets(1).Range("A16").Interior.Pattern = -4142, "partial multi-target run recovers only verified own changes"
End Sub
