Attribute VB_Name = "AtR5Tests"
Option Explicit
Private passed As Long, failed As Long, log As String, book As Object
Public Function AtRunR5Tests(ByVal repo As String, ByVal scratch As String) As String
    On Error GoTo Broken
    Dim editor As Object, stream As Object
    Set editor = Application.Run("'" & ThisWorkbook.Name & "'!AtNewAdminEditor", scratch)
    passed = 0: failed = 0: log = ""
    Set book = Application.Workbooks.Add: book.Worksheets(1).Name = "Input"
    TestCommands scratch
    TestNames scratch
    TestLayout
    book.Close False: Set book = Nothing
    Set stream = CreateObject("Scripting.FileSystemObject").CreateTextFile(scratch & "\r5-assertions.txt", True, True): stream.Write log: stream.Close
    AtRunR5Tests = IIf(failed = 0, "PASS|R5|", "FAIL|R5|") & passed & "|" & failed & IIf(failed = 0, "", "|" & log): Exit Function
Broken:
    AtRunR5Tests = "FAIL|R5|" & Err.Number & "|" & Err.Description & "|" & log
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
End Function
Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    If ok Then passed = passed + 1 Else failed = failed + 1
    log = log & IIf(ok, "PASS ", "FAIL ") & label & vbCrLf
End Sub
Private Sub TestCommands(ByVal scratch As String)
    Dim editor As CAtAdminEditor, result As CAtAdminPreview, root As String, cell As Object, moved As Object, fso As Object, saved As Boolean
    Set fso = CreateObject("Scripting.FileSystemObject"): root = scratch & "\r5-journals": fso.CreateFolder root
    Set editor = AtNewAdminEditor(root): editor.SheetName = "Input": editor.Address = "B2": editor.ExpectedText = "yes"
    editor.FieldName = "Title": editor.Message = "Use yes": editor.Color = 65535: editor.AddNote = False
    Set cell = book.Worksheets(1).Range("B2"): cell.Value2 = "no": cell.Interior.Color = RGB(20, 30, 40)
    book.Worksheets(1).Range("Z1").Formula = "=1+2": book.Worksheets(1).Range("Z2").AddComment "User note"
    book.SaveAs scratch & "\r5.xlsx", 51: saved = book.Saved
    Set result = AtUiRunMode(editor, book, "PREVIEW")
    Check result.Outcome.Status = "FAIL" And cell.Interior.Color = RGB(20, 30, 40) And book.Saved = saved, "PREVIEW is read only"
    Check fso.GetFolder(root).SubFolders.Count = 0, "PREVIEW creates no journal"
    Set result = AtUiRunMode(editor, book, "TEST")
    Check result.Outcome.Status = "FAIL" And cell.Interior.Color = 65535, "TEST FAIL immediately applies fill"
    Check InStr(result.Output, AtUiStatus("APPLIED")) > 0 And Not book.Saved, "actual Writer status visible without autosave"
    editor.Color = RGB(218, 232, 252): editor.Message = "Updated message"
    Set result = AtUiRunMode(editor, book, "TEST")
    Check cell.Interior.Color = editor.Color And result.Remark = "Updated message", "TEST updates color and message"
    cell.ClearContents: Set result = AtUiRunMode(editor, book, "TEST")
    Check result.Outcome.Status = "CANNOT_EVALUATE" And cell.Interior.Color = editor.Color, "incomplete TEST retains existing marks"
    cell.Value2 = "no": Set moved = book.Worksheets(1).Range("D4"): moved.Value2 = "no"
    editor.Address = "D4": Set result = AtUiRunMode(editor, book, "TEST")
    Check cell.Interior.Color = RGB(20, 30, 40) And moved.Interior.Color = editor.Color, "changing address reconciles old and new targets"
    moved.Formula = "=""yes""": Set result = AtUiRunMode(editor, book, "TEST")
    Check result.Outcome.Status = "PASS" And moved.Interior.Pattern = -4142 And moved.Formula = "=""yes""", "TEST PASS clears own marks preserving new formula"
    Check fso.GetFolder(root).SubFolders.Count = 0, "TEST PASS cleans completed journal"
    moved.Value2 = "no": Set result = AtUiRunMode(editor, book, "TEST")
    moved.Interior.Color = RGB(7, 8, 9): Set result = AtUiRunMode(editor, book, "TEST")
    Check result.Outcome.Status = "FAIL" And InStr(result.Diagnostics, "CONFLICT") > 0 And moved.Interior.Color = RGB(7, 8, 9), "rule FAIL separated from conflicting Writer"
    Check fso.GetFolder(root).SubFolders.Count = 1, "conflict evidence retained"
    Check book.Worksheets(1).Range("Z1").Formula = "=1+2" And book.Worksheets(1).Range("Z2").Comment.Text = "User note", "unrelated user formula and Note preserved"
    editor.Address = "E5": editor.AddNote = True: book.Worksheets(1).Range("E5").Value2 = "no"
    Set editor = AtNewAdminEditor(root): editor.SheetName = "Input": editor.Address = "E5": editor.ExpectedText = "yes": editor.AddNote = True
    Set result = AtUiRunMode(editor, book, "TEST")
    Check result.Outcome.Status = "FAIL" And InStr(result.Diagnostics, "CAPABILITY_UNSUPPORTED") > 0 And book.Worksheets(1).Range("E5").Interior.Pattern = -4142, "xlsx Note explicit whole action refusal during TEST"
    book.SaveAs scratch & "\r5.xls", 56
    Set result = AtUiRunMode(editor, book, "TEST")
    Set cell = book.Worksheets(1).Range("E5")
    Check result.Outcome.Status = "FAIL" And Not cell.Comment Is Nothing, "TEST supports proven xls fill and Note"
    editor.Message = "New internal message": editor.Color = RGB(218, 232, 252)
    Set result = AtUiRunMode(editor, book, "TEST")
    Check InStr(cell.Comment.Text, "New internal message") > 0 And cell.Interior.Color = editor.Color, "TEST updates own xls Note text and fill"
    cell.Formula = "=""yes""": Set result = AtUiRunMode(editor, book, "TEST")
    Check result.Outcome.Status = "PASS" And cell.Comment Is Nothing And cell.Interior.Pattern = -4142 And cell.Formula = "=""yes""", "xls TEST PASS clears Note without restoring old input"
End Sub
Private Sub TestNames(ByVal scratch As String)
    Dim value As String, previous As String, address As String, loaded As CAtAdminEditor, editor As CAtAdminEditor, nothingRange As Object, blocked As Boolean, other As Object
    Load AtAdminMain: AtUiFill AtAdminMain: AtUiRefreshBooks AtAdminMain
    AtAdminMain.txtAddress.Text = "B2": address = AtAdminMain.txtAddress.Text
    book.Worksheets(1).Range("A1").Value2 = "Название поля"
    Check AtUiReadFieldName(book, book.Worksheets(1).Range("A1")) = "Название поля", "ordinary title read"
    Check AtUiAcceptFieldName(AtAdminMain, book, book.Worksheets(1).Range("A1")), "title accepted by same UI path"
    Check AtAdminMain.txtField.Text = "Название поля" And InStr(AtAdminMain.lblWhat.Caption, "Название поля") > 0, "title and card updated"
    Check AtAdminMain.txtAddress.Text = address, "title selection independent of locator"
    book.Worksheets(1).Range("F1:H2").Merge: book.Worksheets(1).Range("F1").Value2 = "Merged title"
    Check AtUiReadFieldName(book, book.Worksheets(1).Range("G2")) = "Merged title", "merged interior uses top left value"
    Check AtUiReadFieldName(book, book.Worksheets(1).Range("F1:H2")) = "Merged title", "whole merged selection accepted"
    previous = AtAdminMain.txtField.Text
    Check Not AtUiAcceptFieldName(AtAdminMain, book, nothingRange) And AtAdminMain.txtField.Text = previous, "cancel preserves title"
    On Error Resume Next
    AtUiAcceptFieldName AtAdminMain, book, book.Worksheets(1).Range("J1"): blocked = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    Check blocked And AtAdminMain.txtField.Text = previous, "empty title rejected before mutation"
    book.Worksheets(1).Range("J1").Value2 = CVErr(2015): blocked = False
    On Error Resume Next
    value = AtUiReadFieldName(book, book.Worksheets(1).Range("J1")): blocked = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    Check blocked, "Excel error title rejected"
    Set other = Application.Workbooks.Add
    On Error Resume Next
    value = AtUiReadFieldName(book, other.Worksheets(1).Range("A1")): blocked = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    Check blocked, "foreign book title rejected": other.Close False
    Set editor = AtNewAdminEditor(scratch): editor.SheetName = "Input": editor.Address = address: editor.FieldName = AtAdminMain.txtField.Text
    editor.SaveFile scratch & "\r5-title.xml": Set loaded = AtNewAdminEditor(scratch): loaded.LoadFile scratch & "\r5-title.xml"
    Check loaded.FieldName = previous And AtDefinitionById(loaded.Model, "Field", "FIELD_VALUE").Id = "FIELD_VALUE", "selected title saved without changing FieldId"
    AtAdminMain.txtField.Text = "Edited manually"
    Check InStr(AtAdminMain.lblWhat.Caption, "Edited manually") > 0, "selected title remains editable"
    AtAdminMain.cboBook.ListIndex = 0
    AtAdminMain.cboSheet.ListIndex = -1
    Set other = AtUiPickerBook(AtAdminMain, "PICK_NAME")
    Check other Is book, "title command accepts explicit book without locator sheet"
    blocked = False
    On Error Resume Next
    Set other = AtUiPickerBook(AtAdminMain, "PICK"): blocked = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    Check blocked, "data locator picker still requires its own sheet"
    AtAdminMain.Hide: AtAdminOpen
    Check AtAdminMain.txtField.Text = "Edited manually" And AtAdminMain.Visible, "reopen preserves editor state"
    AtAdminOpen
    Check AtAdminMain.cboBook.ListIndex = 0 And AtAdminMain.txtAddress.Text = address, "repeated open preserves explicit book and locator"
    Check AtUiWatchingWindows And AtUiRaiseWithinExcel(AtAdminMain), "same-process form promoted without global topmost"
    book.Activate: book.Worksheets(1).Activate
    Check AtAdminMain.Visible And AtAdminMain.txtField.Text = "Edited manually", "book and sheet activation does not destroy modeless state"
    book.Worksheets(1).Range("B2").Select
    Check Application.ActiveWorkbook Is book And Application.ActiveCell.Address = "$B$2", "window promotion preserves Excel workbook and cell focus"
    AtUiPicking True
    AtAdminMain.Hide: book.Worksheets(1).Range("C3").Select
    Check Not AtAdminMain.Visible, "Excel events do not reshow form during range picker"
    AtUiPicking False: AtAdminMain.Show 0
    Dim count As Long, form As Object
    For Each form In VBA.UserForms: If TypeName(form) = "AtAdminMain" Then count = count + 1
    Next form
    Check count = 1, "repeated open uses one default instance"
    AtUiHideMain
    Check Not AtUiWatchingWindows And Not AtAdminMain.Visible, "close detaches native Office event receiver"
End Sub
Private Sub TestLayout()
    Dim sizes As Variant, size As Variant, form As Object, control As Object, inside As Boolean
    AtAdminMain.cboMode.ListIndex = 1
    AtAdminMain.txtExpected.Text = String$(160, "W")
    AtAdminMain.txtMessage.Text = String$(400, "M")
    sizes = Array(Array(1020, 720), Array(884, 638), Array(600, 360), Array(480, 300))
    For Each size In sizes
        AtUiLayout AtAdminMain, CDbl(size(0)), CDbl(size(1)): inside = True
        For Each control In AtAdminMain.pnlBody.Controls
            If control.Visible And control.Left + control.Width > AtAdminMain.pnlBody.InsideWidth Then inside = False
        Next control
        Check inside And AtAdminMain.pnlBody.ScrollBars <> 3 And AtAdminMain.pnlBody.ScrollLeft = 0, "no horizontal overflow " & size(0)
        Check AtAdminMain.cmdPreview.Top >= AtAdminMain.pnlBody.Top + AtAdminMain.pnlBody.Height And AtAdminMain.Zoom = 100, "fixed readable actions " & size(0)
    Next size
    Check InStr(AtAdminMain.lblAction.Caption, String$(400, "M")) > 0, "card retains full long message"
    AtUiScroll AtAdminMain, 1
    Check AtAdminMain.pnlBody.ScrollTop > 0, "safe page navigation scrolls content"
    AtUiScroll AtAdminMain, -1
    Check AtAdminMain.pnlBody.ScrollTop = 0, "page navigation returns to top"
    AtAdminMain.txtExpected.Text = "yes": AtAdminMain.txtMessage.Text = "Use yes": AtAdminMain.cboMode.ListIndex = 0
    AtUiLayout AtAdminMain, 1020, 720
    Check AtAdminMain.pnlBody.ScrollBars = 0 And Not AtAdminMain.cmdPageDown.Visible, "wide layout has no unnecessary scrollbars"
    Check AtAdminMain.cboBook.Left > AtAdminMain.lblWhat.Left And AtAdminMain.lstNavigation.Left < AtAdminMain.lblWhat.Left, "wide layout retains three columns"
    AtUiLayout AtAdminMain, 480, 300
    Check AtAdminMain.cboBook.Left = AtAdminMain.lblRuleTitle.Left And AtAdminMain.lblWhat.Top > AtAdminMain.chkNote.Top, "compact layout stacks sections"
    Unload AtAdminMain: Unload AtAdminResults
    Check VBA.UserForms.Count = 0 And Not AtUiWatchingWindows, "forms unload without retained instances or callback objects"
End Sub
