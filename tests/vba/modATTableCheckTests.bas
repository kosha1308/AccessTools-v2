Attribute VB_Name = "modATTableCheckTests"
Option Explicit
Public Sub AT_TestTableChecks()
    Dim book As Workbook, ws As Worksheet, reader As New ATNetworkRequestReader
    Dim checker As New ATNetworkTableChecks, localRunner As New ATNetworkLocalChecks
    Dim source As ATNetworkReadResult, report As ATCheckReport, protocols As New Collection
    Dim issue As ATCheckIssue, item As Variant, number As Long, description As String, i As Long, special As New Collection
    On Error GoTo CleanupError
    protocols.Add "TCP": protocols.Add "UDP"
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    Set source = reader.Read(ws, "TABLES", "2026-10-16", "UTC")
    AT_Assert source.Context.CellRef("TECHNICAL_ROWS_RAW") = "'Лист1'!A14" And source.Context.CellRef("ADDRESS_ROWS_RAW") = "'Лист1'!B16", "fresh collection header refs available for empty tables"
    Set report = checker.Run(source, protocols)
    AT_Assert report.State = atEvalOK And report.ScopeId = "NETWORK_TABLE_INPUT" And report.Checks.Count = 3, "table report scope and check count"
    AT_Assert report.ErrorCount = 2 And report.Issues(1).CellRef = "'Лист1'!A14" And report.Issues(2).CellRef = "'Лист1'!B16", "empty tables have real header issues not fake rows"
    ws.Range("A15:D15").Value2 = Array("192.0.2.1", "198.51.100.1", "tcp; UDP", "443;8443")
    ws.Range("E15").Value2 = "reason": ws.Range("A17").Value2 = "system"
    ws.Range("B17").Value2 = "192.0.2.1": ws.Range("C17").Value2 = "host.example.test": ws.Range("E17").Value2 = "role"
    Set report = checker.Run(reader.Read(ws, "VALID", "2026-10-16", "UTC"), protocols)
    AT_Assert report.State = atEvalOK And report.ErrorCount = 0 And report.WarningCount = 0 And report.SkippedCount = 0, "valid numeric tables and protocol list accepted"
    For Each item In Array("0", "0,443", "65536", "443;bad")
        ws.Range("D15").Value2 = item
        Set report = checker.Run(reader.Read(ws, "PORT", "2026-10-16", "UTC"), protocols)
        AT_Assert report.ErrorCount = 1 And report.Issues(1).FieldId = "TECH_DEST_PORT", "invalid business port " & CStr(item)
    Next item
    ws.Range("D15").Value2 = "1-65535": ws.Range("A15").Value2 = "10.20.30.1308"
    Set report = checker.Run(reader.Read(ws, "IP", "2026-10-16", "UTC"), protocols)
    AT_Assert report.ErrorCount = 1 And report.Issues(1).Code = "IP_INVALID", "numeric invalid IP not salvaged"
    ws.Range("A15").Value2 = "Группа серверов"
    Set report = checker.Run(reader.Read(ws, "GROUP", "2026-10-16", "UTC"), protocols)
    AT_Assert report.ErrorCount = 0 And report.WarningCount = 1 And report.Checks(1).Status = "SKIPPED", "textual IP group requires review not error or PASS"
    ws.Range("A15").Value2 = "192.0.2.1": ws.Range("C15").Value2 = "TCP|UDP"
    Set source = reader.Read(ws, "CATALOG", "2026-10-16", "UTC")
    Set report = checker.Run(source, Nothing)
    AT_Assert report.SkippedCount = 1 And report.WarningCount = 1 And report.ErrorCount = 0, "missing protocol catalog explicitly skipped"
    ws.Range("C15").Value2 = "TCP;UNKNOWN"
    Set report = checker.Run(reader.Read(ws, "PROTOCOL", "2026-10-16", "UTC"), protocols)
    AT_Assert report.ErrorCount = 1 And report.Issues(1).Code = "PROTOCOL_INVALID", "each protocol token checked"
    ws.Range("C15").Value2 = "; | ,"
    Set report = checker.Run(reader.Read(ws, "NO_TOKENS", "2026-10-16", "UTC"), protocols)
    AT_Assert report.ErrorCount = 1 And AT_TestIssueCodeCount(report, "PROTOCOL_INVALID") = 1, "separator-only protocol cannot pass membership"
    special.Add "HTTP-ALT"
    For Each item In Array("HTTP - ALT", "HTTP " & ChrW(&H2014) & " ALT")
        ws.Range("C15").Value2 = item
        Set report = checker.Run(reader.Read(ws, "LITERAL", "2026-10-16", "UTC"), special)
        AT_Assert report.ErrorCount = 1, "protocol catalog comparison never joins dash-separated tokens"
    Next item
    ws.Range("C15").Value2 = "http-alt"
    Set report = checker.Run(reader.Read(ws, "LITERAL_OK", "2026-10-16", "UTC"), special)
    AT_Assert report.ErrorCount = 0, "literal hyphenated protocol still accepted case-insensitively"
    Set report = checker.Run(reader.Read(ws, "", "bad", ""), protocols)
    AT_Assert report.State = atEvalError And InStr(report.Issues(1).Message, "CONFIG_") = 0, "blocked table message remains Russian without raw internal code"
    ws.Range("C15").Value2 = "TCP": ws.Range("A17").ClearContents: ws.Range("C17").MergeArea.ClearContents: ws.Range("E17").MergeArea.ClearContents
    Set report = checker.Run(reader.Read(ws, "LOWER", "2026-10-16", "UTC"), protocols)
    AT_Assert report.ErrorCount = 1 And report.WarningCount = 2, "hostname required system and role warnings"
    ws.Range("A15").Value2 = CVErr(xlErrValue): ws.Range("B17").Value2 = CVErr(xlErrValue)
    Set report = checker.Run(reader.Read(ws, "ERR", "2026-10-16", "UTC"), protocols)
    AT_Assert report.ErrorCount = 3 And AT_TestIssueCodeCount(report, "CELL_ERROR") = 2, "table Excel errors stay cell errors"
    ws.Range("A15").Value2 = "192.0.2.1": ws.Range("B17").Value2 = "192.0.2.1"
    ws.Rows(15).Resize(2).Insert
    ws.Range("A15:D15").Value2 = Array("192.0.2.1", "198.51.100.1", "TCP", "443")
    ws.Range("A17:D17").Value2 = Array("192.0.2.1", "198.51.100.1", "tcp", "443")
    ws.Range("E15:G17").Merge: ws.Range("E15").Value2 = "shared reason"
    Set source = reader.Read(ws, "DUP", "2026-10-16", "UTC"): Set report = checker.Run(source, protocols)
    AT_Assert AT_TestIssueCodeCount(report, "ROW_DUPLICATE") = 1, "duplicate warning across spacer with merged reason"
    AT_Assert source.Context.ResolveField("TECHNICAL_ROWS_RAW").Value.Count = 2 And source.Context.CellRef("ADDRESS_ROWS_RAW") = "'Лист1'!B18", "spacer skipped and moved header ref current"
    ws.Range("A15").Value2 = "GROUP|A": ws.Range("B15").Value2 = "GROUP_B"
    ws.Range("A17").Value2 = "GROUP": ws.Range("B17").Value2 = "A|GROUP_B"
    Set report = checker.Run(reader.Read(ws, "KEY", "2026-10-16", "UTC"), protocols)
    AT_Assert AT_TestIssueCodeCount(report, "ROW_DUPLICATE") = 0, "duplicate key delimiter collision avoided"
    Set report = localRunner.Run(source, protocols)
    AT_Assert report.ScopeId = "NETWORK_LOCAL" And report.Checks.Count = 9, "localRunner report aggregates six document and three table checks"
    protocols.Add "BAD|ENTRY"
    Set report = localRunner.Run(source, protocols)
    AT_Assert report.State = atEvalError And report.Checks.Count = 1 And report.Checks(1).Status = "BLOCKED", "bad catalog discards partial document PASS outcomes"
    ws.Range("A1").Value2 = "old template"
    Set report = localRunner.Run(reader.Read(ws, "BAD", "2026-10-16", "UTC"), protocols)
    AT_Assert report.State = atEvalError And report.Issues.Count = 1 And report.Issues(1).Code = "TEMPLATE_MISMATCH", "template rejection precedes catalog validation"
    book.Close False: Set book = Nothing
    Set report = checker.Run(source, Nothing)
    AT_Assert report.State = atEvalOK, "table source snapshot detached from closed workbook"
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1": AT_TestNetworkReferenceFixture ws
    ws.Rows(15).Resize(260).Insert: ws.Range("A15:A274").Value2 = "192.0.2.1"
    Set report = checker.Run(reader.Read(ws, "LIMIT", "2026-10-16", "UTC"), Nothing)
    AT_Assert report.State = atEvalError And report.Issues.Count = 1 And report.Checks.Count = 1, "issue limit blocks rather than publishes truncated success"
    book.Close False
    Exit Sub
CleanupError:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise number, , description
End Sub
Public Function AT_TestIssueCodeCount(ByVal report As ATCheckReport, ByVal code As String) As Long
    Dim issue As ATCheckIssue
    For Each issue In report.Issues
        If issue.Code = code Then AT_TestIssueCodeCount = AT_TestIssueCodeCount + 1
    Next issue
End Function
