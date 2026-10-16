Attribute VB_Name = "modATAddressMatchTests"
Option Explicit
Public Sub AT_TestAddressMatch()
    Dim book As Workbook, ws As Worksheet, reader As New ATNetworkRequestReader, matcher As New ATNetworkAddressMatch
    Dim localChecks As New ATNetworkLocalChecks, report As ATCheckReport, saved As ATNetworkReadResult
    Dim item As Variant, issue As ATCheckIssue, found As Long, i As Long, number As Long, description As String
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    For Each item In Array( _
        Array("192.0.2.1", "192.0.2.1", "PASS"), Array("192.0.2.1", "192.0.2.1/32", "FAIL"), _
        Array("192.0.2.0/24", "192.0.2.1", "FAIL"), Array("192.0.2.7/24", "192.0.2.0/24", "FAIL"), _
        Array("192.0.2.1-5", "192.0.2.1-192.0.2.5", "PASS"), Array("192.0.2.1-1", "192.0.2.1", "FAIL"), _
        Array("192.0.2.1-5", "192.0.2.1-3;192.0.2.4-5", "FAIL"), _
        Array("192.0.2.1,2", "192.0.2.2;192.0.2.1", "PASS"), Array("192.0.2.1|192.0.2.1", "192.0.2.1", "PASS"), _
        Array("192.0.2.1308", "192.0.2.13", "FAIL"), Array("192.0.2.1308", "192.0.2.1308", "PASS"), _
        Array("192.0.2.256", "", "FAIL"), Array("192.0.2.1oops", "192.0.2.1", "FAIL"), Array("192.0.2.1text", "192.0.2.1", "FAIL"), _
        Array("192.0.2.1;192.0.2.256;2", "192.0.2.1;192.0.2.256;192.0.2.2", "FAIL"), _
        Array("Группа серверов", "192.0.2.1", "SKIPPED"), Array("192.0.2.1", "Группа серверов", "SKIPPED"), _
        Array("192.0.x.x", "192.0.x.x", "SKIPPED"), Array("2001:db8::1", "2001:db8::1", "SKIPPED"), _
        Array("", "192.0.2.1", "SKIPPED"), Array("", "", "SKIPPED"), _
        Array("192.000.002.001", "192.0.2.1", "PASS"))
        ws.Range("A15").Value2 = item(0): ws.Range("B17").Value2 = item(1)
        Set report = matcher.Run(reader.Read(ws, "PAIR", "2026-10-16", "UTC"))
        AT_Assert report.State = atEvalOK, "declaration pair state " & CStr(item(0))
        If report.State = atEvalOK Then AT_Assert report.Checks(1).Status = CStr(item(2)), "declaration pair outcome " & CStr(item(0)) & " / " & CStr(item(1))
    Next item
    ws.Range("A15").Value2 = "192.0.2.1": ws.Range("B15").Value2 = "198.51.100.1": ws.Range("B17").Value2 = "203.0.113.1"
    Set report = matcher.Run(reader.Read(ws, "REFS", "2026-10-16", "UTC"))
    AT_Assert report.ErrorCount = 3 And report.Issues(1).Code = "ADDRESS_EXTRA" And report.Issues(1).CellRef = "'Лист1'!B17", "extra lower issue first with actual ref"
    AT_Assert report.Issues(2).Code = "ADDRESS_MISSING" And report.Issues(2).CellRef = "'Лист1'!A15" And report.Issues(3).CellRef = "'Лист1'!B15", "both upper fields retain missing issue refs"
    ws.Rows(15).Resize(2).Insert
    ws.Range("A15:A17").Merge: ws.Range("A15").Value2 = "192.0.2.1"
    ws.Range("B15").Value2 = "198.51.100.1": ws.Range("B17").Value2 = "198.51.100.1": ws.Range("B19").ClearContents
    Set saved = reader.Read(ws, "MERGE", "2026-10-16", "UTC"): Set report = matcher.Run(saved)
    found = 0
    For Each issue In report.Issues
        If issue.CellRef = "'Лист1'!A15" Then found = found + 1
    Next issue
    AT_Assert found = 1 And report.ErrorCount = 3, "merged declaration ref deduped without losing distinct source refs"
    ws.Range("A15").Value2 = CVErr(xlErrValue)
    Set report = matcher.Run(reader.Read(ws, "CELL_ERROR", "2026-10-16", "UTC"))
    AT_Assert report.State = atEvalOK And report.Checks(1).Status = "SKIPPED" And report.Issues.Count = 0, "cell error prevents false matching and false missing issues"
    Set report = localChecks.Run(saved, Nothing)
    AT_Assert report.Checks.Count = 10 And report.Checks(10).CheckId = "NET.ADDRESS_MATCH", "new exact match appended to existing nine checks"
    AT_Assert AT_CheckTitle("NET.ADDRESS_MATCH") = "Соответствие адресов двух таблиц", "existing writer receives new Russian label"
    ws.Range("A1").Value2 = "old template"
    Set report = matcher.Run(reader.Read(ws, "", "bad", ""))
    AT_Assert report.State = atEvalError And report.Issues(1).Code = "TEMPLATE_MISMATCH", "template rejection wins before declaration parsing"
    Set report = matcher.Run(Nothing)
    AT_Assert report.State = atEvalError And report.Checks.Count = 1, "missing matcher input controlled"
    Set report = matcher.Run(AT_TestMatchTextInput(String$(65536, " ")))
    AT_Assert report.State = atEvalOK And report.Checks(1).Status = "SKIPPED", "matcher accepts configured text boundary"
    Set report = matcher.Run(AT_TestMatchTextInput(String$(65537, " ")))
    AT_Assert report.State = atEvalError, "matcher rejects oversized typed text before tokenization"
    If report.State = atEvalError Then AT_Assert report.Issues(1).Code = "RUNTIME_LIMIT_TEXT", "typed matcher text limit stable code"
    book.Close False: Set book = Nothing
    Set report = matcher.Run(saved)
    AT_Assert report.ErrorCount = 3, "exact matching snapshot detached from edits and closed source"
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1": AT_TestNetworkReferenceFixture ws
    ws.Range("A15").Value2 = AT_TestDeclarationList(991)
    Set report = matcher.Run(reader.Read(ws, "LIMIT", "2026-10-16", "UTC"))
    AT_Assert report.State = atEvalError And report.Issues.Count = 1 And report.Checks(1).Status = "BLOCKED", "match issue cap gives blocked report rather than truncation"
    ws.Range("A15").Value2 = AT_TestDeclarationList(900)
    For i = 1 To 25
        ws.Cells(i + 16, 2).Value2 = "203.0.113." & CStr(i)
    Next i
    Set report = localChecks.Run(reader.Read(ws, "AGGREGATE_LIMIT", "2026-10-16", "UTC"), Nothing)
    AT_Assert report.State = atEvalError And report.Issues.Count = 1 And report.Checks.Count = 1, "aggregate issue cap discards all partial positives"
    book.Close False
    Exit Sub
CleanupError:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise number, , description
End Sub
Private Function AT_TestDeclarationList(ByVal count As Long) As String
    Dim i As Long
    For i = 1 To count
        If i > 1 Then AT_TestDeclarationList = AT_TestDeclarationList & ";"
        AT_TestDeclarationList = AT_TestDeclarationList & "10.1." & CStr(i \ 256) & "." & CStr(i Mod 256)
    Next i
End Function
Private Function AT_TestMatchTextInput(ByVal text As String) As ATNetworkReadResult
    Dim context As New ATEvaluationContext, result As New ATNetworkReadResult, row As New ATRuntimeRow
    Dim rows As New Collection, upper As New ATRuntimeValue, lower As New ATRuntimeValue
    context.Initialize "TYPED_BOUNDARY", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    row.Initialize "TECHNICAL_TABLE_RAW", "15"
    row.AddField "SOURCE_RAW", AT_RuntimeOK(AT_RuntimeScalar("TEXT", text)), "'Лист1'!A15"
    row.AddField "DESTINATION_RAW", AT_RuntimeOK(AT_RuntimeState("TEXT", atValueEmpty)), "'Лист1'!B15"
    rows.Add row: upper.InitializeRows "TECHNICAL_TABLE_RAW", rows
    lower.InitializeRows "ADDRESS_TABLE_RAW", New Collection
    context.AddField "TECHNICAL_ROWS_RAW", AT_RuntimeOK(upper)
    context.AddField "ADDRESS_ROWS_RAW", AT_RuntimeOK(lower)
    result.Initialize atEvalOK, context, Nothing: Set AT_TestMatchTextInput = result
End Function
