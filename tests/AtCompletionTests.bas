Attribute VB_Name = "AtCompletionTests"
Option Explicit
Private passed As Long, failed As Long, log As String, book As Object, fso As Object
Public Function AtRunCompletionTests(ByVal repo As String, ByVal scratch As String) As String
    On Error GoTo Broken
    Dim store As Object, stream As Object
    Set store = Application.Run("'" & ThisWorkbook.Name & "'!AtNewJournalStore", scratch)
    passed = 0: failed = 0: log = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set book = Application.Workbooks.Add: book.Worksheets(1).Name = "Input"
    TestStore scratch
    TestEditor scratch
    TestLayout scratch
    book.Close False: Set book = Nothing
    Set stream = fso.CreateTextFile(scratch & "\completion-assertions.txt", True, True): stream.Write log: stream.Close
    AtRunCompletionTests = IIf(failed = 0, "PASS|COMPLETION|", "FAIL|COMPLETION|") & passed & "|" & failed & IIf(failed = 0, "", "|" & log)
    Exit Function
Broken:
    AtRunCompletionTests = "FAIL|COMPLETION|" & Err.Number & "|" & Err.Description & "|" & log
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
End Function
Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    If ok Then passed = passed + 1 Else failed = failed + 1
    log = log & IIf(ok, "PASS ", "FAIL ") & label & vbCrLf
End Sub
Private Sub TestStore(ByVal scratch As String)
    Dim store As CAtJournalStore, other As CAtJournalStore, journal As CAtOwnershipJournal
    Dim path As String, stream As Object, actions As New Collection, intent As New CAtActionIntent, receipts As Collection, blocked As Boolean
    Set store = AtNewJournalStore(scratch): path = store.NewDirectory()
    Check fso.FolderExists(path), "exclusive local directory"
    Set journal = AtNewOwnershipJournal(book, path): Set other = AtNewJournalStore(scratch)
    Check Not other.PruneCompleted(journal, path) And fso.FileExists(path & "\session.atj"), "another store cannot adopt cleanup ownership"
    Set stream = fso.CreateTextFile(path & "\user.txt", False, True): stream.Write "keep": stream.Close
    Check Not store.PruneCompleted(journal, path) And fso.FileExists(path & "\session.atj"), "unknown file prevents cleanup"
    fso.DeleteFile path & "\user.txt"
    Check store.PruneCompleted(journal, path) And Not fso.FolderExists(path), "verified unused lease cleaned"
    path = store.NewDirectory(): Set journal = AtNewOwnershipJournal(book, path)
    intent.ActionInstanceKey = "LIFECYCLE": intent.Sheet = "Input": intent.Address = "A1"
    intent.AddNote = False: intent.Message = "remark": intent.FillColor = 65535: intent.Severity = "ERROR": actions.Add intent
    book.Worksheets(1).Range("A1").Value2 = "wrong"
    Set receipts = AtApplyPlan(AtPlanActions(book, actions, "APPLY", journal), journal)
    Check receipts(1).Status = "APPLIED", "managed journal supports Writer"
    Check Not store.PruneCompleted(journal, path), "active marks retain journal"
    book.Worksheets(1).Range("A1").Interior.Color = RGB(1, 2, 3)
    Set receipts = AtClearOwned(book, journal, "CONFLICT")
    Check receipts(1).Status = "CONFLICT" And Not store.PruneCompleted(journal, path), "formatting conflict retains journal"
    Check book.Worksheets(1).Range("A1").Interior.Color = RGB(1, 2, 3), "foreign color preserved"
    path = store.NewDirectory(): Set journal = AtNewOwnershipJournal(book, path): journal.RequireRecovery
    Check Not store.PruneCompleted(journal, path) And fso.FolderExists(path), "interrupted journal retained"
    path = store.NewDirectory(): Set journal = AtNewOwnershipJournal(book, path)
    Set stream = fso.OpenTextFile(path & "\session.atj", 8, False, -1): stream.Write "changed": stream.Close
    Check Not store.PruneCompleted(journal, path) And fso.FileExists(path & "\session.atj"), "tampered manifest retained"
    path = store.NewDirectory(): Set journal = AtNewOwnershipJournal(book, path): fso.CreateFolder path & "\unknown"
    Check Not store.PruneCompleted(journal, path) And fso.FolderExists(path & "\unknown"), "unknown subfolder retained"
    On Error Resume Next
    Set other = AtNewJournalStore("relative"): blocked = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    Check blocked, "relative storage rejected"
End Sub
Private Sub TestEditor(ByVal scratch As String)
    Dim editor As CAtAdminEditor, preview As CAtAdminPreview, root As String, cell As Object, original As Long
    Dim directory As Object, stream As Object, retained As String, other As Object, clearText As String
    root = scratch & "\automatic": fso.CreateFolder root
    Set editor = AtNewAdminEditor(root): editor.SheetName = "Input": editor.Address = "B2": editor.ExpectedText = "yes"
    editor.AddNote = False: editor.Message = "Use yes": editor.Color = 65535
    Set cell = book.Worksheets(1).Range("B2"): cell.Value2 = "no": original = RGB(20, 30, 40): cell.Interior.Color = original
    Set preview = editor.Preview(book)
    Check fso.GetFolder(root).SubFolders.Count = 0 And cell.Interior.Color = original, "preview creates no journal"
    Set preview = editor.Apply(book, "")
    Check preview.Outcome.Status = "FAIL" And cell.Interior.Color = 65535 And fso.GetFolder(root).SubFolders.Count = 1, "FAIL applies without folder prompt"
    Set preview = editor.Apply(book, "")
    Check fso.GetFolder(root).SubFolders.Count = 1, "repeat reuses journal"
    cell.Formula = "=""yes""": Set preview = editor.Apply(book, "")
    Check preview.Outcome.Status = "PASS" And cell.Interior.Color = original And cell.Formula = "=""yes""", "PASS clears own fill preserving corrected formula"
    Check fso.GetFolder(root).SubFolders.Count = 0 And editor.NeedsJournal, "completed journal removed and binding released"
    cell.Value2 = "no": Set preview = editor.Apply(book, ""): editor.Clear book
    Check fso.GetFolder(root).SubFolders.Count = 0 And cell.Interior.Color = original, "clear removes completed files"
    editor.AddNote = True: book.SaveAs scratch & "\completion.xlsx", 51
    Set preview = editor.Apply(book, "")
    Check InStr(preview.Diagnostics, "CAPABILITY_UNSUPPORTED") > 0 And cell.Interior.Color = original, "unsupported Notes not downgraded"
    Check fso.GetFolder(root).SubFolders.Count = 0, "rejected action leaves no lease files"
    editor.AddNote = False: Set preview = editor.Apply(book, "")
    For Each directory In fso.GetFolder(root).SubFolders: retained = directory.Path: Next directory
    Set stream = fso.CreateTextFile(retained & "\user.txt", False, True): stream.Write "keep": stream.Close
    clearText = editor.Clear(book)
    Check editor.NeedsJournal And fso.FileExists(retained & "\user.txt") And cell.Interior.Color = original, "verified rollback releases binding while foreign files retained"
    Set other = Application.Workbooks.Add: other.Worksheets(1).Name = "Input": other.Worksheets(1).Range("B2").Value2 = "no"
    Set preview = editor.Apply(other, "")
    Check preview.Outcome.Status = "FAIL" And other.Worksheets(1).Range("B2").Interior.Color = 65535, "editor can switch book after retained completed evidence"
    clearText = editor.Clear(other): other.Close False
    book.SaveAs scratch & "\completion.xls", 56: editor.AddNote = True
    Set preview = editor.Apply(book, "")
    Check Not cell.Comment Is Nothing And cell.Interior.Color = 65535, "automatic journal supports proven xls Note profile"
    cell.Formula = "=""yes""": Set preview = editor.Apply(book, "")
    Check preview.Outcome.Status = "PASS" And cell.Comment Is Nothing And cell.Interior.Color = original And cell.Formula = "=""yes""", "xls corrected formula clears Note and fill safely"
    Check fso.GetFolder(root).SubFolders.Count = 1 And fso.FileExists(retained & "\user.txt"), "only intentionally retained evidence remains after successful cycles"
End Sub
Private Sub TestLayout(ByVal scratch As String)
    Dim sizes As Variant, size As Variant, editor As CAtAdminEditor, preview As CAtAdminPreview
    Set editor = AtNewAdminEditor(scratch): editor.SheetName = "Input": editor.Address = "B2": editor.ExpectedText = "yes"
    editor.SaveFile scratch & "\completion-ui.xml"
    Load AtAdminMain: AtUiLoadFile AtAdminMain, scratch & "\completion-ui.xml"
    Check AtAdminMain.cboSheet.Value = "Input" And AtAdminMain.cboBook.ListIndex = -1, "loaded sheet visible before workbook choice"
    sizes = Array(Array(884, 638), Array(600, 360), Array(480, 300))
    For Each size In sizes
        AtUiLayout AtAdminMain, CDbl(size(0)), CDbl(size(1))
        Check Fits(AtAdminMain) And AtAdminMain.Zoom = 100, "main bounds readable zoom " & size(0)
        Check AtAdminMain.pnlBody.ScrollHeight > AtAdminMain.pnlBody.Height And AtAdminMain.pnlBody.ScrollWidth >= 860, "main scroll reach " & size(0)
        Check AtAdminMain.cmdClose.Top >= AtAdminMain.pnlBody.Top + AtAdminMain.pnlBody.Height, "fixed actions " & size(0)
    Next size
    Check AtAdminMain.txtExpected.Font.Size >= 10 And AtAdminMain.cmdTest.Font.Size >= 10, "fonts not shrunk"
    Load AtAdminResults
    book.Worksheets(1).Range("B2").Value2 = "no"
    Set preview = editor.Preview(book): AtUiFillResult AtAdminResults, preview
    Check AtAdminResults.lblResult.ForeColor = RGB(180, 40, 40), "FAIL red"
    Check Not AtAdminResults.txtDiagnostic.Visible And Not AtAdminResults.lblDiagnostic.Visible, "diagnostics collapsed"
    Check ScrollReach(AtAdminResults.pnlBody), "collapsed output entirely reachable"
    Check InStr(AtAdminResults.txtDiagnostic.Text, "OK") = 0, "FAIL explanation without ambiguous OK"
    For Each size In sizes
        AtUiLayout AtAdminResults, CDbl(size(0)), CDbl(size(1))
        Check Fits(AtAdminResults) And AtAdminResults.Zoom = 100, "compact result bounds " & size(0)
    Next size
    AtUiToggleDiagnostics AtAdminResults
    Check AtAdminResults.txtDiagnostic.Visible And AtAdminResults.pnlBody.ScrollHeight = 458, "diagnostics expand with scroll reach"
    Check ScrollReach(AtAdminResults.pnlBody), "expanded diagnostics entirely reachable"
    AtUiToggleDiagnostics AtAdminResults
    Check Not AtAdminResults.txtDiagnostic.Visible, "diagnostics collapse"
    book.Worksheets(1).Range("B2").Value2 = "yes": Set preview = editor.Preview(book): AtUiFillResult AtAdminResults, preview
    Check AtAdminResults.lblResult.ForeColor = RGB(27, 119, 65) And Not AtAdminResults.cmdCopy.Enabled, "PASS green without remark"
    book.Worksheets(1).Range("B2").ClearContents: Set preview = editor.Preview(book): AtUiFillResult AtAdminResults, preview
    Check AtAdminResults.lblResult.ForeColor = RGB(151, 98, 0), "incomplete warning color"
    Unload AtAdminMain: Unload AtAdminResults
End Sub
Private Function Fits(ByVal form As Object) As Boolean
    Dim control As Object
    Fits = True
    For Each control In form.Controls
        If control.Parent Is form Then
            If control.Left < 0 Or control.Top < 0 Or control.Left + control.Width > form.InsideWidth + 1 Or control.Top + control.Height > form.InsideHeight + 1 Then Fits = False: log = log & "OUTSIDE " & control.Name & " parent=" & TypeName(control.Parent) & vbCrLf
        End If
    Next control
End Function
Private Function ScrollReach(ByVal frame As Object) As Boolean
    Dim control As Object
    ScrollReach = True
    For Each control In frame.Controls
        If control.Visible Then
            If control.Top + control.Height > frame.ScrollHeight Or control.Left + control.Width > frame.ScrollWidth Then ScrollReach = False
        End If
    Next control
End Function
