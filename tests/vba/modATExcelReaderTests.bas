Attribute VB_Name = "modATExcelReaderTests"
Option Explicit

Public Sub AT_TestExcelReader()
    Dim book As Workbook, ws As Worksheet, lower As Worksheet, wrong As Worksheet
    Dim spec As ATExcelTableSpec, layout As ATExcelTableLayout, reader As New ATExcelTableReader
    Dim result As ATEvalResult, field As ATEvalResult, saved As ATRuntimeValue
    Dim savedNumber As Long, savedText As String, last As Long
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = book.Worksheets(1): ws.Name = "AT's form"
    Set lower = book.Worksheets.Add: Set wrong = book.Worksheets.Add
    Set spec = AT_TestExcelSpec("TECH", 1, 3, 1, 5)
    spec.AddColumn "SRC_RAW", "Source", "INHERITED"
    spec.AddColumn "DST_RAW", "Destination", "INHERITED"
    spec.AddColumn "PROTOCOL_RAW", "Protocol", "INHERITED"
    spec.AddColumn "PORT_RAW", "Port", "INHERITED"
    spec.AddColumn "REASON_RAW", "Reason", "NONE"
    ws.Range("A1:E1").Value2 = Array("Source", "Destination", "Protocol", "Port", "Reason")
    ws.Range("A2").Value2 = "192.0.2.1": ws.Range("B2").Value2 = "198.51.100.2"
    ws.Range("C2").Value2 = "TCP": ws.Range("D2").Value2 = 443
    ws.Range("E2:E6").Merge: ws.Range("E2").Value2 = "reason"
    ws.Range("A5:A6").Merge: ws.Range("A5").Value2 = "bad-prefix 192.0.2.9" & vbLf & "203.0.113.4"
    ws.Range("B5").Value2 = "198.51.100.5": ws.Range("B6").Value2 = "198.51.100.6"
    ws.Range("C5:C6").Merge: ws.Range("C5").Value2 = "TCP"
    ws.Range("D5:D6").Merge: ws.Range("D5").Value2 = "80,443"
    ws.Range("A3").Formula = "="""""
    ws.Range("B4").Value2 = " " & ChrW(160) & vbTab & vbCrLf & ChrW(&H200B) & ChrW(&HFEFF)
    ws.Range("A2").Interior.Color = RGB(10, 20, 30)
    Set layout = spec.Locate(ws)
    Set result = reader.Read(ws, layout, 6)
    AT_Assert result.State = atEvalOK, "reader returns typed rows"
    If result.State <> atEvalOK Then GoTo Finish
    Set saved = result.Value
    AT_Assert saved.Count = 3, "merged reason and unicode/formula spacers do not create connectivity"
    AT_Assert saved.Row(1).RowId = "2" And saved.Row(2).RowId = "5" And saved.Row(3).RowId = "6", "reader preserves physical row ids"
    AT_Assert saved.Row(3).CellRef("SRC_RAW") = "'AT''s form'!A5", "merged source anchor metadata escaped"
    Set field = saved.Row(3).ResolveField("SRC_RAW")
    AT_Assert field.Value.Scalar = "bad-prefix 192.0.2.9" & vbLf & "203.0.113.4", "raw invalid multiline address retained whole"
    Set field = saved.Row(1).ResolveField("PORT_RAW")
    AT_Assert field.Value.Scalar = "443", "numeric Excel port becomes invariant raw text"
    AT_Assert ws.Range("A3").Formula = "=""""" And ws.Range("E2:E6").MergeCells, "reader preserves formula and merges"
    AT_Assert ws.Range("A2").Interior.Color = RGB(10, 20, 30), "reader preserves formatting"
    ws.Range("A2").Value2 = "changed"
    Set field = saved.Row(1).ResolveField("SRC_RAW")
    AT_Assert field.Value.Scalar = "192.0.2.1", "snapshot detached from workbook mutation"
    AT_TestExcelRowsDSL saved
    Set result = reader.Read(wrong, layout, 6)
    AT_Assert result.Diagnostics(1).Code = "EXCEL_LAYOUT_SHEET_MISMATCH", "reader rejects wrong sheet"
    Set result = reader.Read(ws, layout, 10002)
    AT_Assert result.Diagnostics(1).Code = "EXCEL_ROW_LIMIT", "reader bounds physical rows before bulk read"
    Set result = reader.Read(ws, layout, 5)
    AT_Assert result.Diagnostics(1).Code = "EXCEL_MERGE_BOUNDS", "reader rejects merge crossing lower table boundary"
    Set result = reader.Read(ws, layout, 1)
    AT_Assert result.State = atEvalOK And result.Value.Count = 0, "empty table and reuse after errors"
    Set spec = AT_TestExcelSpec("ADDR", 1, 3, 1, 4)
    spec.AddColumn "SYSTEM_RAW", "System", "DIRECT"
    spec.AddColumn "ADDRESS_RAW", "Address", "DIRECT"
    spec.AddColumn "HOST_RAW", "Hostname", "DIRECT"
    spec.AddColumn "ROLE_RAW", "Role", "DIRECT"
    lower.Range("A1:D1").Value2 = Array("System", "Address", "Hostname", "Role")
    lower.Range("A2:A5").Merge: lower.Range("A2").Value2 = "system"
    lower.Range("D2:D5").Merge: lower.Range("D2").Value2 = "role"
    lower.Range("B2").Value2 = "192.0.2.10": lower.Range("B5").Value2 = "203.0.113.10"
    lower.Range("C6").Value2 = "only-host"
    lower.Range("D7").Value2 = "separate role"
    lower.Range("B8").Value2 = CVErr(xlErrNA)
    lower.Range("B1048576").Interior.Color = vbYellow
    lower.Range("C1048575").Formula = "="""""
    Set layout = spec.Locate(lower)
    last = reader.LastValueRow(lower, layout)
    AT_Assert last = 8, "last row ignores formatting and empty formula at sheet end"
    Set result = reader.Read(lower, layout, last)
    AT_Assert result.State = atEvalOK, "reader accepts lower rows with a field error"
    AT_Assert result.Value.Count = 5, "lower inherited aux skips blanks and separate aux retained"
    AT_Assert result.Value.Row(2).RowId = "5" And result.Value.Row(5).RowId = "8", "lower blank does not stop extraction"
    Set field = result.Value.Row(2).ResolveField("SYSTEM_RAW")
    AT_Assert field.Value.Scalar = "system", "lower auxiliary merge inherited on real address row"
    Set field = result.Value.Row(3).ResolveField("ADDRESS_RAW")
    AT_Assert field.Value.ValueState = atValueEmpty, "missing cell content represented as EMPTY"
    Set field = result.Value.Row(5).ResolveField("ADDRESS_RAW")
    AT_Assert field.State = atEvalError And field.Diagnostics(1).Code = "EXCEL_CELL_ERROR", "Excel error does not disappear or become text"
    lower.Range("A2:D8").UnMerge: lower.Range("A2:D8").ClearContents
    AT_Assert reader.LastValueRow(lower, layout) = 1, "empty lower table has header boundary"
    Set result = reader.Read(lower, layout, 0)
    AT_Assert result.Diagnostics(1).Code = "EXCEL_DATA_BOUNDS", "reader rejects inverted bounds"
    AT_TestExcelReaderWidth reader, wrong
    AT_TestExcelMergeEdges reader, wrong
Finish:
    book.Close False
    If Not saved Is Nothing Then
        Set field = saved.Row(1).ResolveField("SRC_RAW")
        AT_Assert field.Value.Scalar = "192.0.2.1", "snapshot remains usable after source workbook closes"
    End If
    Exit Sub
CleanupError:
    savedNumber = Err.Number: savedText = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise savedNumber, , savedText
End Sub

Private Sub AT_TestExcelMergeEdges(ByVal reader As ATExcelTableReader, ByVal ws As Worksheet)
    Dim spec As ATExcelTableSpec, layout As ATExcelTableLayout, result As ATEvalResult
    ws.Cells.Clear
    ws.Range("A1").Value2 = "Source": ws.Range("B1").Value2 = "Destination"
    Set spec = AT_TestExcelSpec("T", 1, 1, 1, 2)
    spec.AddColumn "A", "Source", "INHERITED": spec.AddColumn "B", "Destination", "INHERITED"
    Set layout = spec.Locate(ws)
    ws.Range("A2:B2").Merge: ws.Range("A2").Value2 = "192.0.2.1"
    Set result = reader.Read(ws, layout, 2)
    AT_Assert result.State = atEvalError, "horizontal merge cannot fabricate another field"
    If result.State = atEvalError Then AT_Assert result.Diagnostics(1).Code = "EXCEL_MERGE_COLUMNS", "cross-field horizontal merge explicit error"
    ws.Range("A2:B2").UnMerge: ws.Range("A2").ClearContents
    ws.Range("A2:A5").Merge: ws.Range("A2").Value2 = "192.0.2.1"
    AT_Assert reader.LastValueRow(ws, layout) = 5, "last value covers vertical merged key extent"
    Set result = reader.Read(ws, layout, reader.LastValueRow(ws, layout))
    AT_Assert result.State = atEvalOK, "merged final value can be read using discovered boundary"
    If result.State = atEvalOK Then AT_Assert result.Value.Count = 4, "last merged key inherits through final physical row"
    Set spec = AT_TestExcelSpec("D", 1, 1, 1, 2)
    spec.AddColumn "A", "Source", "DIRECT": spec.AddColumn "B", "Destination", "DIRECT"
    Set layout = spec.Locate(ws)
    Set result = reader.Read(ws, layout, reader.LastValueRow(ws, layout))
    AT_Assert result.State = atEvalOK, "last direct auxiliary merge accepted"
    If result.State = atEvalOK Then AT_Assert result.Value.Count = 1, "auxiliary extent does not create blank records"
    ws.Range("A2:A5").UnMerge: ws.Range("A2").ClearContents
    ws.Range("B1").ClearContents: ws.Range("C1").Value2 = "Destination"
    Set spec = AT_TestExcelSpec("H", 1, 1, 1, 3)
    spec.AddColumn "A", "Source", "INHERITED": spec.AddColumn "B", "Destination", "INHERITED"
    Set layout = spec.Locate(ws)
    ws.Range("A2:B2").Merge: ws.Range("A2").Value2 = "192.0.2.20"
    ws.Range("C2").Value2 = "198.51.100.20"
    Set result = reader.Read(ws, layout, 2)
    AT_Assert result.State = atEvalOK, "horizontal merge within one logical field remains legal"
    If result.State = atEvalOK Then
        AT_Assert result.Value.Row(1).ResolveField("B").Value.Scalar = "198.51.100.20", "owned horizontal merge cannot alter adjacent field"
    End If
    ws.Range("A2:B2").UnMerge: ws.Range("A2:C2").ClearContents
    ws.Range("C10001").Value2 = "198.51.100.30"
    Set result = reader.Read(ws, layout, 10001)
    AT_Assert result.State = atEvalOK, "exact physical-row limit accepted"
    If result.State = atEvalOK Then AT_Assert result.Value.Row(1).RowId = "10001", "last allowed row retained after blanks"
End Sub

Private Sub AT_TestExcelRowsDSL(ByVal rows As ATRuntimeValue)
    Dim schema As New ATCompileContext, compiler As New ATCompiler
    Dim context As New ATEvaluationContext, evaluator As New ATEvaluator, result As ATEvalResult
    schema.AddField "ROWS", "COLLECTION<ROW<TECH>>"
    schema.AddField "PROTOCOL_RAW", "TEXT", "TECH"
    context.Initialize "EXCEL_TEST", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    context.AddField "ROWS", AT_RuntimeOK(rows)
    Set result = evaluator.Evaluate(compiler.Compile("COUNT(@ROWS) = 3 AND ALLROWS(@ROWS AS t, @t.PROTOCOL_RAW = ""TCP"")", schema), context)
    AT_Assert result.State = atEvalOK, "extracted rows compile and evaluate in existing runtime"
    If result.State = atEvalOK Then AT_Assert result.Value.Scalar, "extracted connectivity rules agree with real rows"
End Sub

Private Sub AT_TestExcelReaderWidth(ByVal reader As ATExcelTableReader, ByVal ws As Worksheet)
    Dim spec As ATExcelTableSpec, layout As ATExcelTableLayout, result As ATEvalResult
    Set spec = AT_TestExcelSpec("WIDE", 1, 1, 1, 120)
    spec.AddColumn "A", "Left", "DIRECT": spec.AddColumn "B", "Right", "DIRECT"
    ws.Cells(1, 1).Value2 = "Left": ws.Cells(1, 120).Value2 = "Right"
    Set layout = spec.Locate(ws)
    Set result = reader.Read(ws, layout, 10001)
    AT_Assert result.Diagnostics(1).Code = "EXCEL_CELL_LIMIT", "reader bounds wide rectangle before allocation"
    Set spec = AT_TestExcelSpec("SINGLE", 1, 1, 1, 1)
    spec.AddColumn "A", "Left", "DIRECT"
    Set layout = spec.Locate(ws)
    ws.Cells(2, 1).Value2 = 0
    Set result = reader.Read(ws, layout, 2)
    AT_Assert result.Value.Count = 1, "single-cell block and zero participate"
    AT_Assert result.Value.Row(1).ResolveField("A").Value.Scalar = "0", "zero retained rather than blank"
End Sub
