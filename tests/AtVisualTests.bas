Attribute VB_Name = "AtVisualTests"
Option Explicit
Private passed As Long, failed As Long, log As String, book As Object

Public Function AtRunVisualTests(ByVal repo As String, ByVal scratch As String) As String
    On Error GoTo Broken
    passed = 0: failed = 0: log = ""
    Dim editor As CAtAdminEditor, loaded As New CAtAdminEditor, p As CAtPackage, q As CAtPackage
    Dim preview As CAtAdminPreview, cell As Object, fso As Object, saved As Boolean, formula As String, count As Long
    Set editor = Application.Run("'" & ThisWorkbook.Name & "'!AtNewAdminEditor")
    Check editor.FieldName <> "", "editor defaults"
    editor.TemplateName = "Visual template": editor.FieldName = "Name": editor.SheetName = "Request": editor.Address = "B2"
    editor.ExpectedText = "expected": editor.Message = "Correct the name": editor.Color = 65535
    Set p = editor.Model
    Check p.SchemaVersion = "3.0.0", "adapter uses schema3"
    Check AtValue(AtDefinitionById(p, "Action", "ACTION_FAIL"), "Audience") = "FirstLine", "new annotations are internal, not initiator distribution"
    editor.SaveFile scratch & "\visual-rule.xml": loaded.LoadFile scratch & "\visual-rule.xml"
    Check AtToXml(p) = AtToXml(loaded.Model), "adapter save load full roundtrip"
    loaded.ExpectedText = "updated": loaded.Message = "Updated remark": loaded.Color = AtUiColor(2): loaded.Address = "D5"
    loaded.SaveFile scratch & "\visual-edited.xml": editor.LoadFile scratch & "\visual-edited.xml"
    Check editor.ExpectedText = "updated" And editor.Message = "Updated remark" And editor.Address = "D5" And editor.Color = AtUiColor(2), "loaded settings editable and persist"
    Check AtDefinitionById(editor.Model, "Field", "FIELD_VALUE").Id = "FIELD_VALUE" And AtDefinitionById(editor.Model, "Rule", "RULE_EQUAL").Id = "RULE_EQUAL", "logical identities stable"
    Set book = Application.Workbooks.Add(): book.Worksheets(1).Name = "Request"
    Set cell = book.Worksheets(1).Range("D5"): cell.Value2 = "wrong"
    book.Worksheets(1).Range("E5").Formula = "=1+2": formula = book.Worksheets(1).Range("E5").Formula
    book.Worksheets(1).Range("F5").AddComment "User note"
    book.SaveAs scratch & "\visual-target.xlsx", 51
    saved = book.Saved: count = Application.Workbooks.Count
    Set preview = editor.Preview(book)
    Check preview.Outcome.Status = "FAIL" And preview.ActualText = "wrong", "preview reads actual value and FAIL"
    Check preview.InitiatorText = "Updated remark", "copy payload excludes target and diagnostics"
    Check cell.Interior.Pattern = -4142 And cell.Comment Is Nothing And book.Saved = saved, "preview has no workbook or Saved side effects"
    Check book.Worksheets(1).Range("E5").Formula = formula And book.Worksheets(1).Range("F5").Comment.Text = "User note" And Application.Workbooks.Count = count, "preview preserves formula note and open workbook inventory"
    Set fso = CreateObject("Scripting.FileSystemObject"): fso.CreateFolder scratch & "\visual-journal"
    Set preview = editor.Apply(book, scratch & "\visual-journal")
    Check cell.Interior.Color = editor.Color, "explicit apply uses existing Writer"
    cell.Value2 = "updated": Set preview = editor.Preview(book)
    Check preview.Outcome.Status = "PASS" And preview.InitiatorText = "", "preview PASS has no initiator remark"
    Check cell.Interior.Color = editor.Color, "PASS preview does not clear previous marks"
    Set preview = editor.Apply(book, "")
    Check cell.Interior.Pattern = -4142 And cell.Value2 = "updated", "explicit apply rechecks and clears own marks"
    cell.Value2 = "wrong": Set preview = editor.Preview(book)
    cell.Value2 = "updated": Set preview = editor.Apply(book, "")
    Check preview.Outcome.Status = "PASS" And cell.Interior.Pattern = -4142, "apply never reuses stale displayed FAIL"
    cell.ClearContents: Set preview = editor.Preview(book)
    Check preview.Outcome.Status = "CANNOT_EVALUATE" And preview.InitiatorText = "", "EMPTY preview not successful"
    cell.Value2 = 42: Set preview = editor.Preview(book)
    Check preview.Outcome.Status = "CANNOT_EVALUATE", "INVALID preview not successful"
    editor.SheetName = "Missing": Set preview = editor.Preview(book)
    Check preview.Outcome.Status = "CANNOT_EVALUATE", "MISSING preview not successful"
    editor.SheetName = "Request": editor.AddNote = True: cell.Value2 = "wrong"
    Set preview = editor.Apply(book, "")
    Check InStr(preview.Diagnostics, "CAPABILITY_UNSUPPORTED") > 0 And cell.Interior.Pattern = -4142, "unsupported Notes explicit refusal without downgrade"
    Dim other As Object, blocked As Boolean, clearText As String
    Set other = Application.Workbooks.Add()
    On Error Resume Next
    clearText = editor.Clear(other): blocked = Err.Number <> 0: Err.Clear
    On Error GoTo Broken
    Check blocked, "clear requires explicit selection of bound workbook"
    other.Close False: Set other = Nothing
    clearText = editor.Clear(book)
    Check editor.NeedsJournal, "successful clear releases binding for another workbook"
    Load AtAdminMain: AtUiFill AtAdminMain: AtUiRefreshBooks AtAdminMain
    Check AtAdminMain.cboBook.ListIndex = -1, "UI requires explicit workbook choice"
    Check AtUiControlCount(AtAdminMain) > 30, "main form contains real controls"
    Check AtAdminMain.cmdTest.Caption = ChrW(&H41F) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H442) & ChrW(&H435) & ChrW(&H441) & ChrW(&H442) & ChrW(&H438) & ChrW(&H440) & ChrW(&H43E) & ChrW(&H432) & ChrW(&H430) & ChrW(&H442) & ChrW(&H44C), "Russian control captions preserved"
    AtAdminMain.cboBook.ListIndex = 0: AtAdminMain.cboSheet.ListIndex = 0
    AtAdminMain.txtField.Text = "Visual field": AtAdminMain.txtAddress.Text = "D5": AtAdminMain.txtExpected.Text = "text"
    AtAdminMain.cboColor.ListIndex = 2: AtAdminMain.chkNote.Value = True
    Dim card As String: card = AtAdminMain.lblWhat.Caption
    Check InStr(card, "Visual field") > 0 And InStr(card, "D5") > 0 And InStr(card, "text") > 0, "property edits update central card"
    AtAdminMain.cboMode.ListIndex = 1
    Check AtAdminMain.lblDetails.Visible And AtAdminMain.lblWhat.Caption = card And AtAdminMain.chkNote.Value, "advanced mode retains the same data"
    AtAdminMain.cboMode.ListIndex = 0
    Check Not AtAdminMain.lblDetails.Visible And AtAdminMain.cboColor.ListIndex = 2 And AtAdminMain.txtExpected.Text = "text", "simple mode retains settings"
    book.Worksheets.Add.Name = "Other"
    editor.SheetName = "Other": editor.AddNote = False
    editor.SaveFile scratch & "\visual-other-sheet.xml"
    AtUiLoadFile AtAdminMain, scratch & "\visual-other-sheet.xml"
    AtAdminMain.cboBook.ListIndex = 0
    Check AtAdminMain.cboSheet.Value = "Other", "load cannot replace persisted sheet with old UI selection"
    Check InStr(AtAdminMain.lblWhat.Caption, "Other!D5") > 0, "loaded sheet reflected in central card"
    Load AtAdminResults: AtUiFillResult AtAdminResults, preview
    Check AtAdminResults.txtRemark.Text = "Updated remark" And AtAdminResults.cmdCopy.Enabled, "results separates initiator text"
    Check InStr(AtAdminResults.txtDiagnostic.Text, "CAPABILITY_UNSUPPORTED") > 0, "internal diagnostics displayed separately"
    Check Not AtAdminMain.Visible And Not AtAdminResults.Visible, "form API tests do not claim interactive UI acceptance"
    Dim control As Object, fits As Boolean: fits = True
    For Each control In AtAdminMain.Controls
        If control.Left < 0 Or control.Top < 0 Or control.Left + control.Width > AtAdminMain.InsideWidth Or control.Top + control.Height > AtAdminMain.InsideHeight Then fits = False
    Next control
    Check fits, "main controls inside client bounds at design size"
    fits = True
    For Each control In AtAdminResults.Controls
        If control.Left < 0 Or control.Top < 0 Or control.Left + control.Width > AtAdminResults.InsideWidth Or control.Top + control.Height > AtAdminResults.InsideHeight Then fits = False
    Next control
    Check fits, "result controls inside client bounds at design size"
    AtAdminOpen
    Check AtAdminMain.Visible, "public launch macro actually opens modeless constructor"
    AtAdminMain.Hide
    Unload AtAdminMain: Unload AtAdminResults
    book.Close False: Set book = Nothing
    Dim file As Object: Set file = fso.CreateTextFile(scratch & "\visual-assertions.txt", True, True): file.Write log: file.Close
    AtRunVisualTests = IIf(failed = 0, "PASS|VISUAL|", "FAIL|VISUAL|") & passed & "|" & failed & IIf(failed = 0, "", "|" & log)
    Exit Function
Broken:
    AtRunVisualTests = "FAIL|VISUAL|" & Err.Number & "|" & Err.Description & "|" & log
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    If Not other Is Nothing Then other.Close False
End Function

Private Sub Check(ByVal condition As Boolean, ByVal label As String)
    If condition Then passed = passed + 1 Else failed = failed + 1
    log = log & IIf(condition, "PASS ", "FAIL ") & label & vbCrLf
End Sub
