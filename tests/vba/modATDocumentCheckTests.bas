Attribute VB_Name = "modATDocumentCheckTests"
Option Explicit
Public Sub AT_TestDocumentChecks()
    Dim book As Workbook, ws As Worksheet, reader As New ATNetworkRequestReader, checker As New ATNetworkDocumentChecks
    Dim report As ATCheckReport, saved As ATCheckReport, source As ATNetworkReadResult
    Dim issue As ATCheckIssue, outcome As ATCheckOutcome, items As Collection, item As Variant, caught As Boolean
    Dim number As Long, description As String, malformed As New ATNetworkReadResult, context As New ATEvaluationContext
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    ws.Range("C7").Value2 = "Предоставить доступ": ws.Range("C8").Value2 = "Да"
    Set report = checker.Run(reader.Read(ws, "BLANK", "2026-10-16", "Europe/Moscow"))
    AT_Assert report.State = atEvalOK And report.ScopeId = "NETWORK_DOCUMENT", "document report scope explicit"
    AT_Assert report.Checks.Count = 6 And report.ErrorCount = 5 And report.WarningCount = 1, "blank grant collects all document issues"
    AT_Assert report.SkippedCount = 1 And report.HasErrors, "blank email syntax skipped without duplicate issue"
    AT_Assert report.Issues.Count = 6, "blank issues exact count"
    Set issue = report.Issues(1)
    AT_Assert issue.CheckId = "NET.REQUIRED_FIELDS" And issue.FieldId = "INITIATOR_FULL_NAME" And issue.CellRef = "'Лист1'!A5", "required issue actual field and source ref"
    AT_Assert InStr(issue.Message, "Заполните") > 0 And issue.Severity = "ERROR", "required issue localized"
    Set items = report.Issues: items.Remove 1
    AT_Assert report.Issues.Count = 6, "report issue collection copied"
    Set items = report.Checks: items.Remove 1
    AT_Assert report.Checks.Count = 6, "report check collection copied"
    On Error Resume Next
    issue.Initialize "X", "X", "ERROR", "X", "X", "X"
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo CleanupError
    AT_Assert caught, "published issue immutable"
    Set outcome = report.Checks(1)
    On Error Resume Next
    outcome.Initialize "X", "PASS", "X"
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo CleanupError
    AT_Assert caught, "published outcome immutable"
    On Error Resume Next
    report.Initialize atEvalOK, "OTHER", New Collection, New Collection
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo CleanupError
    AT_Assert caught, "report immutable"
    ws.Range("A5").Value2 = "Тестовый пользователь": ws.Range("D5").Value2 = "Отдел": ws.Range("E5").Value2 = "Должность"
    ws.Range("F5").Value2 = "user@example.test": ws.Range("C11").Value2 = "Рабочая задача"
    ws.Range("C12").NumberFormat = "@": ws.Range("C12").Value2 = "до 30.12.2026"
    Set source = reader.Read(ws, "VALID", "2026-10-16", "Europe/Moscow"): Set report = checker.Run(source): Set saved = report
    AT_Assert report.State = atEvalOK And report.ErrorCount = 0 And report.WarningCount = 0 And Not report.HasErrors, "valid grant with date words and no optional contact accepted"
    AT_Assert report.SkippedCount = 0 And report.Issues.Count = 0 And report.Checks.Count = 6, "six completed local checks"
    For Each outcome In report.Checks
        AT_Assert outcome.Status = "PASS", "valid local check " & outcome.CheckId
    Next outcome
    For Each item In Array(Array("Бессрочно", 0), Array("16.10.2026", 0), Array("до 17/10/2026", 0), Array("15.10.2026", 1), Array("31.02.2026", 1), Array("46100", 1), Array("", 1))
        ws.Range("C12").Value2 = item(0)
        Set report = checker.Run(reader.Read(ws, "DATES", "2026-10-16", "UTC"))
        AT_Assert report.ErrorCount = CLng(item(1)), "grant date validity against explicit RunDate " & CStr(item(0))
        If report.ErrorCount > 0 Then AT_Assert report.Issues(1).CellRef = "'Лист1'!C12", "validity issue source cell"
    Next item
    For Each item In Array("предоставить", " ПРЕДОСТАВИТЬ ДОСТУП ", "отозвать", " ОТОЗВАТЬ ДОСТУП ")
        ws.Range("C7").Value2 = item
        If Left$(AT_ExcelHeader(CStr(item)), 8) = "отозвать" Then ws.Range("C12").Value2 = CVErr(xlErrValue) Else ws.Range("C12").Value2 = "Бессрочно"
        Set report = checker.Run(reader.Read(ws, "ALIASES", "2026-10-16", "UTC"))
        AT_Assert report.ErrorCount = 0, "request type alias and revoke ignore validity " & CStr(item)
    Next item
    ws.Range("C7").Value2 = "неизвестный тип": ws.Range("C12").Value2 = CVErr(xlErrValue)
    Set report = checker.Run(reader.Read(ws, "UNKNOWN", "2026-10-16", "UTC"))
    AT_Assert report.ErrorCount = 1 And report.SkippedCount = 1, "unknown type does not imply grant or duplicate validity errors"
    ws.Range("C7").Value2 = "Предоставить доступ": ws.Range("C12").Value2 = "Бессрочно"
    ws.Range("F5").Value2 = "bad-mail": ws.Range("C8").Value2 = "Нет": ws.Range("C11").MergeArea.ClearContents
    Set report = checker.Run(reader.Read(ws, "MULTI", "2026-10-16", "UTC"))
    AT_Assert report.ErrorCount = 2 And report.WarningCount = 1 And report.Issues.Count = 3, "independent errors and warning all collected"
    AT_Assert report.Issues(1).FieldId = "EMAIL" And report.Issues(1).CellRef = "'Лист1'!F5", "email syntax issue precise"
    AT_Assert InStr(report.Summary, "Ошибок: 2") > 0 And InStr(report.Summary, "предупреждений: 1") > 0, "Russian report summary"
    ws.Range("A5").Value2 = CVErr(xlErrValue): ws.Range("F5").Value2 = CVErr(xlErrNA)
    ws.Range("C8").Value2 = CVErr(xlErrValue): ws.Range("C12").Value2 = CVErr(xlErrValue)
    ws.Range("C11").Value2 = "задача"
    Set report = checker.Run(reader.Read(ws, "CELLERR", "2026-10-16", "UTC"))
    AT_Assert report.State = atEvalOK And report.ErrorCount = 4 And report.Issues.Count = 4, "cell errors become one issue per affected field"
    For Each issue In report.Issues
        AT_Assert issue.Code = "CELL_ERROR", "cell error not replaced by empty/invalid message"
    Next issue
    ws.Range("A1").Value2 = "old template"
    Set report = checker.Run(reader.Read(ws, "", "bad", ""))
    AT_Assert report.State = atEvalError And report.Checks.Count = 1 And report.Issues.Count = 1, "structural rejection blocks all business checks"
    AT_Assert report.Issues(1).Code = "TEMPLATE_MISMATCH" And report.Checks(1).Status = "BLOCKED", "template mismatch retains reason"
    Set report = checker.Run(Nothing)
    AT_Assert report.State = atEvalError And report.Issues.Count = 1 And report.HasErrors, "missing input controlled report"
    context.Initialize "BAD", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    malformed.Initialize atEvalOK, context, Nothing
    Set report = checker.Run(malformed)
    AT_Assert report.State = atEvalError And report.Checks.Count = 1, "malformed input exposes no partial positive outcomes"
    AT_TestEmailSyntaxCorpus
Finish:
    book.Close False
    AT_Assert saved.ErrorCount = 0 And saved.Issues.Count = 0, "report survives edits and closed source"
    Exit Sub
CleanupError:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise number, , description
End Sub
Private Sub AT_TestEmailSyntaxCorpus()
    Dim item As Variant
    For Each item In Array("user@example.test", "USER_1@example-domain.test", "user+tag@sub.example.test", " user@example.test ", "a@..com", "a..b@example.test")
        AT_Assert AT_EmailSyntaxValid(CStr(item)), "compatible ASCII email accepted " & CStr(item)
    Next item
    For Each item In Array("", "bad", "a@@b.test", "a@b.t", "a b@example.test", "a@example.123", "a@example.test extra", "почта@example.test", "a@.com", "@example.test", "a@example.test" & vbLf & "x@y.test")
        AT_Assert Not AT_EmailSyntaxValid(CStr(item)), "invalid ASCII email rejected " & CStr(item)
    Next item
End Sub
Public Function AT_VerifyDocumentChecksReference(ByVal workbookName As String) As String
    Dim reader As New ATNetworkRequestReader, checker As New ATNetworkDocumentChecks, report As ATCheckReport
    Set report = checker.Run(reader.Read(Application.Workbooks(workbookName).Worksheets(1), "REFERENCE", "2026-10-16", "Europe/Moscow"))
    AT_VerifyDocumentChecksReference = "STATE=" & CStr(report.State) & "|CHECKS=" & CStr(report.Checks.Count) & "|ERRORS=" & CStr(report.ErrorCount) & "|WARNINGS=" & CStr(report.WarningCount)
End Function
