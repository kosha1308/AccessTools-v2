Attribute VB_Name = "modATUserFeedbackTests"
Option Explicit
Public Sub AT_TestUserFeedback()
    Dim fs As Object, folder As String, rulePath As String, supersetPath As String, stream As Object
    Dim settings As New ATNetworkSourceSettings, loaded As New ATNetworkSourceSettings, caught As Boolean
    Dim feedback As New ATNetworkUserFeedback, policy As ATNetworkPolicyReport, book As Workbook, ws As Worksheet
    Dim issues As New Collection, decisions As New Collection, checks As New Collection, report As New ATCheckReport, message As String
    Dim menu As Object
    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = fs.BuildPath(fs.GetSpecialFolder(2), "AccessTools-feedback-" & fs.GetTempName): fs.CreateFolder folder
    rulePath = folder & "\Правила.xlsx": supersetPath = folder & "\Выгрузка.xlsx"
    fs.CreateTextFile(rulePath).Close: fs.CreateTextFile(supersetPath).Close
    AT_Assert Not loaded.Load(folder), "source preferences missing is normal first selection"
    settings.Save rulePath, supersetPath, folder
    AT_Assert loaded.Load(folder) And loaded.Available, "source preferences survive a new object/session"
    AT_Assert loaded.RulebookPath = rulePath And loaded.SupersetPath = supersetPath, "Unicode source paths preserved"
    On Error Resume Next
    settings.Save rulePath, rulePath, folder: caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    AT_Assert caught, "preferences reject same source path before persisting"
    fs.DeleteFile supersetPath
    AT_Assert loaded.Load(folder) And Not loaded.Available And loaded.RulebookPath = rulePath, "removed export keeps remembered paths but requires reselection"
    Set stream = fs.CreateTextFile(folder & "\sources.ini", True, True): stream.Write "bad": stream.Close
    AT_Assert Not loaded.Load(folder) And Len(loaded.RulebookPath) = 0, "malformed preferences never retain stale paths"
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    ws.Range("A1").Value2 = "unchanged": ws.Range("A1").AddComment "Чужое примечание"
    ws.Range("E15:G15").Merge
    ws.Range("A1").Interior.Color = RGB(10, 20, 30)
    ws.Range("A1").FormatConditions.Add Type:=xlExpression, Formula1:="=1=2"
    AddFeedbackDecision issues, decisions, "A1", "Исправьте имя", True, False, True, "ERROR", False
    AddFeedbackDecision issues, decisions, "G15", "Заполните обоснование", True, False, True, "ERROR", False
    AddFeedbackDecision issues, decisions, "B2", "Только координатору", False, False, False, "ERROR", False
    AddFeedbackDecision issues, decisions, "C3", "Предупреждение", True, False, True, "WARNING", False
    AddFeedbackDecision issues, decisions, "D4", "Без отметки по правилам", True, False, True, "ERROR", True
    report.Initialize atEvalOK, "NETWORK_ACCESS_SUPERSET_POLICY", issues, checks
    Set policy = New ATNetworkPolicyReport: policy.Initialize report, decisions
    feedback.Prepare policy
    AT_Assert feedback.ErrorCount = 3 And InStr(feedback.UserErrorsText, "A1") > 0 And InStr(feedback.UserErrorsText, "G15") > 0, "copy payload includes all user errors with exact references"
    AT_Assert InStr(feedback.UserErrorsText, "Только координатору") = 0 And InStr(feedback.UserErrorsText, "Предупреждение") = 0, "coordinator errors and warnings excluded from user-error clipboard payload"
    AT_Assert feedback.ApplyNotes(ws, "RUN1") = 2, "unconfigured user input errors get notes at merged anchors"
    AT_Assert ws.Range("A1").FormatConditions.Count = 2, "error highlight overlays rather than replacing original formatting"
    If ws.Range("A1").FormatConditions.Count = 2 Then AT_Assert ws.Range("A1").FormatConditions(1).Interior.Color = RGB(255, 199, 206), "legacy red used for errors"
    If Not ws.Range("A1").Comment Is Nothing Then AT_Assert InStr(ws.Range("A1").Comment.Text, "Чужое примечание") > 0 And InStr(ws.Range("A1").Comment.Text, "Исправьте имя") > 0, "foreign note text retained alongside owned block"
    AT_Assert Not ws.Range("E15").Comment Is Nothing And ws.Range("D4").Comment Is Nothing, "merged target resolved and explicit no-annotation respected"
    feedback.ApplyNotes ws, "RUN2"
    AT_Assert ws.Range("A1").FormatConditions.Count = 2, "repeat replaces own highlight without accumulating rules"
    If Not ws.Range("A1").Comment Is Nothing Then AT_Assert InStr(ws.Range("A1").Comment.Text, "RUN1") = 0 And InStr(ws.Range("A1").Comment.Text, "RUN2") > 0, "repeat replaces only owned note block"
    AT_Assert ws.Range("A1").Value2 = "unchanged", "request cell values never edited"
    Set menu = AT_BuildFeedbackMenu(feedback.ErrorCount, False)
    AT_Assert menu.Controls.Count = 6 And menu.Controls(2).Enabled, "result popup exposes one-click user-error copying and result actions"
    AT_Assert InStr(menu.Controls(2).OnAction, ThisWorkbook.Name) > 0 And InStr(menu.Controls(2).OnAction, "AT_ResultCopy") > 0, "menu actions qualified to this addin instead of first version"
    menu.Delete: Set menu = Nothing
    Set menu = AT_BuildFeedbackMenu(0, False)
    AT_Assert Not menu.Controls(2).Enabled, "copy action disabled when there are no user errors"
    menu.Delete: Set menu = Nothing
    ws.Protect Password:="feedbackfixture"
    AT_Assert feedback.ApplyNotes(ws, "PROTECTED") = 0 And feedback.NoteFailures > 0, "protected sheet reports note failure without losing report or editing cells"
    ws.Unprotect Password:="feedbackfixture"
    feedback.ClearNotes ws
    ws.Range("E15").AddComment String$(32760, "x")
    feedback.ApplyNotes ws, "LONG"
    AT_Assert feedback.NoteFailures > 0 And Len(ws.Range("E15").Comment.Text) = 32760, "oversize combined note preserves full foreign text and reports omission"
    ws.Range("E15").Comment.Delete
    Set issues = New Collection: Set decisions = New Collection: Set checks = New Collection: Set report = New ATCheckReport
    report.Initialize atEvalOK, "NETWORK_ACCESS_SUPERSET_POLICY", issues, checks
    Set policy = New ATNetworkPolicyReport: policy.Initialize report, decisions: feedback.Prepare policy
    feedback.ApplyNotes ws, "CLEAN"
    AT_Assert ws.Range("A1").FormatConditions.Count = 1 And ws.Range("A1").Interior.Color = RGB(10, 20, 30), "clear removes only own highlights and restores original fill and conditional rule"
    AT_Assert ws.Range("A1").Comment.Text = "Чужое примечание" And ws.Range("E15").Comment Is Nothing, "fixed issues clear own notes and retain original foreign text"
    Set issues = New Collection: Set decisions = New Collection: Set checks = New Collection: Set report = New ATCheckReport
    AddFeedbackDecision issues, decisions, "A1:A2", "Неточный адрес", True, False, True, "ERROR", False
    report.Initialize atEvalOK, "NETWORK_ACCESS_SUPERSET_POLICY", issues, checks
    Set policy = New ATNetworkPolicyReport: policy.Initialize report, decisions: feedback.Prepare policy
    AT_Assert feedback.ApplyNotes(ws, "INVALID_REF") = 0 And feedback.NoteFailures > 0, "invalid cell reference omission stays visible after old-note cleanup"
Cleanup:
    On Error Resume Next
    If Not menu Is Nothing Then menu.Delete
    If Not book Is Nothing Then book.Close False
    If Not stream Is Nothing Then stream.Close
    If Len(folder) > 0 Then
        If StrComp(fs.GetParentFolderName(folder), CStr(fs.GetSpecialFolder(2)), vbTextCompare) = 0 And Left$(fs.GetFileName(folder), 21) = "AccessTools-feedback-" Then fs.DeleteFolder folder, True
    End If
    On Error GoTo 0
    Exit Sub
Failed:
    message = Err.Description: AT_Assert False, "user feedback unexpected error: " & message: Resume Cleanup
End Sub
Private Sub AddFeedbackDecision(ByVal issues As Collection, ByVal decisions As Collection, ByVal address As String, ByVal message As String, ByVal inUser As Boolean, ByVal inRequest As Boolean, ByVal userError As Boolean, ByVal severity As String, ByVal configured As Boolean)
    Dim issue As New ATCheckIssue, decision As New ATNetworkPolicyDecision
    issue.Initialize "NET.INPUT", "VALUE_INVALID", severity, "FIELD", "'Лист1'!" & address, message
    decision.Initialize issue, issue, "", inUser, inRequest, False, userError, configured
    issues.Add issue: decisions.Add decision
End Sub
