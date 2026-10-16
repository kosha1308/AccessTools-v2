Attribute VB_Name = "modATNetworkNormalizeTests"
Option Explicit
Public Sub AT_TestNetworkNormalize()
    Dim book As Workbook, ws As Worksheet, reader As New ATNetworkRequestReader
    Dim adapter As New ATNetworkRequestNormalizer, source As ATNetworkReadResult, result As ATNetworkReadResult
    Dim context As ATEvaluationContext, saved As ATEvaluationContext, rows As ATRuntimeValue
    Dim item As Variant, caught As Boolean, number As Long, description As String
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    ws.Range("A15:D15").Value2 = Array("192.0.2.0/24", "198.51.100.1", " tcp ", "443;8443")
    ws.Range("E15").Value2 = "reason" & vbLf & "line2"
    ws.Range("A17").Value2 = "system": ws.Range("B17").Value2 = "192.0.2.1-192.0.2.2"
    ws.Range("C17").Value2 = "host.example.test": ws.Range("E17").Value2 = " role "
    ws.Range("C10").Value2 = "db_test": ws.Range("C12").Value2 = 46100#
    Set source = reader.Read(ws, "NORMALIZE", "2026-10-16", "Europe/Moscow")
    Set result = adapter.Normalize(source, "DMY_DOT", "Бессрочно")
    AT_Assert result.State = atEvalOK, "reader to typed normalization" & NormalizeDetails(result)
    If result.State <> atEvalOK Then GoTo Finish
    Set context = result.Context: Set saved = context.Snapshot
    AT_Assert context.ResolveField("DB_ACCOUNT").Value.Scalar = "db_test" And context.CellRef("DB_ACCOUNT") = "'Лист1'!C10", "document value and ref retained"
    AT_Assert context.ResolveField("JUSTIFICATION").Value.ValueState = atValueEmpty, "document blank preserved"
    AT_Assert context.ResolveField("VALIDITY_DATE").Value.DataType = "DATE" And CDbl(context.ResolveField("VALIDITY_DATE").Value.Scalar) = 46100#, "native validity serial converted"
    AT_Assert Not context.ResolveField("VALIDITY_IS_UNLIMITED").Value.Scalar, "date is not unlimited"
    AT_Assert context.CellRef("VALIDITY_DATE") = "'Лист1'!C12", "typed validity ref"
    Set rows = context.ResolveField("TECHNICAL_ROWS").Value
    AT_Assert rows.Count = 1 And rows.Row(1).RowId = "15" And rows.RowTableId = "TECHNICAL_TABLE", "typed technical row identity"
    AT_Assert rows.Row(1).ResolveField("TECH_PROTOCOL").Value.Scalar = "TCP", "protocol canonicalized without allowlist invention"
    AT_Assert rows.Row(1).ResolveField("TECH_DEST_PORT").Value.Domain.Count = 2, "port set parsed"
    AT_Assert rows.Row(1).ResolveField("TECH_SOURCE_IP").Value.DeclarationKeys(1) = "CIDR:192.0.2.0/24", "IP declaration preserved"
    AT_Assert rows.Row(1).CellRef("TECH_DEST_PORT") = "'Лист1'!D15", "typed row actual source ref"
    AT_Assert rows.Row(1).ResolveField("TECH_REASON").Value.Scalar = "reason" & vbLf & "line2", "reason lines unchanged"
    Set rows = context.ResolveField("ADDRESS_ROWS").Value
    AT_Assert rows.Count = 1 And rows.Row(1).RowId = "17" And rows.RowTableId = "ADDRESS_TABLE", "typed lower row identity"
    AT_Assert rows.Row(1).ResolveField("ROLE").Value.Scalar = " role ", "auxiliary raw spacing retained"
    AT_TestNormalizedDSL context
    AT_Assert source.Context.ResolveField("TECHNICAL_ROWS_RAW").Value.Row(1).ResolveField("PROTOCOL_RAW").Value.Scalar = " tcp ", "normalization never mutates raw source"
    On Error Resume Next
    context.AddField "EXTRA", AT_RuntimeOK(AT_RuntimeScalar("TEXT", "x"))
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo CleanupError
    AT_Assert caught, "normalized snapshot frozen"
    ws.Range("D15").Value2 = "443;bad": ws.Range("B17").Value2 = "192.0.2.1;bad"
    ws.Range("C10").Value2 = CVErr(xlErrValue): ws.Range("C12").Value2 = CVErr(xlErrNA)
    Set result = adapter.Normalize(reader.Read(ws, "INVALID", "2026-10-16", "UTC"), "ISO", "Бессрочно")
    AT_Assert result.State = atEvalOK, "invalid data remains usable context for full issue collection"
    Set context = result.Context
    AT_Assert context.ResolveField("DB_ACCOUNT").State = atEvalError, "document Excel error retained"
    AT_Assert context.ResolveField("VALIDITY_DATE").State = atEvalError And context.ResolveField("VALIDITY_IS_UNLIMITED").State = atEvalError, "validity Excel error not masked"
    AT_Assert context.ResolveField("TECHNICAL_ROWS").Value.Row(1).ResolveField("TECH_DEST_PORT").Value.ValueState = atValueInvalid, "invalid mixed port list never partially salvaged"
    AT_Assert context.ResolveField("ADDRESS_ROWS").Value.Row(1).ResolveField("ADDRESS_IP").Value.ValueState = atValueInvalid, "invalid mixed IP list never partially salvaged"
    ws.Range("A15").Value2 = CVErr(xlErrValue): ws.Range("D15").ClearContents
    ws.Range("C15").Value2 = " new-protocol "
    For Each item In Array(Array(" БЕССРОЧНО ", False, atValueMissing, True), Array("16.10.2026", False, atValuePresent, False), Array("46100", False, atValueInvalid, False), Array(True, False, atValueInvalid, False), Array(0#, True, atValuePresent, False), Array(60#, False, atValueInvalid, False), Array("", False, atValueEmpty, False))
        ws.Range("C12").NumberFormat = "General": ws.Range("C12").MergeArea.ClearContents
        If VarType(item(0)) = vbString Then ws.Range("C12").NumberFormat = "@"
        ws.Range("C12").Value2 = item(0): book.Date1904 = CBool(item(1))
        Set result = adapter.Normalize(reader.Read(ws, "VARIANT", "2026-10-16", "UTC"), "DMY_DOT", "Бессрочно")
        AT_Assert result.State = atEvalOK, "validity variant extraction"
        Set context = result.Context
        AT_Assert context.ResolveField("VALIDITY_DATE").Value.ValueState = CLng(item(2)), "validity type aware value state"
        AT_Assert context.ResolveField("VALIDITY_IS_UNLIMITED").Value.Scalar = CBool(item(3)), "explicit unlimited classification"
        AT_Assert context.ResolveField("TECHNICAL_ROWS").Value.Row(1).ResolveField("TECH_SOURCE_IP").State = atEvalError, "row Excel error preserved"
        AT_Assert context.ResolveField("TECHNICAL_ROWS").Value.Row(1).ResolveField("TECH_DEST_PORT").Value.ValueState = atValueEmpty, "typed blank port retained"
        AT_Assert context.ResolveField("TECHNICAL_ROWS").Value.Row(1).ResolveField("TECH_PROTOCOL").Value.Scalar = "NEW-PROTOCOL", "unregistered protocol not invented as invalid"
    Next item
    For Each item In Array(Array("AUTO", "Бессрочно", "CONFIG_DATE_FORMAT"), Array("ISO", "", "CONFIG_UNLIMITED_KEYWORD"), Array("ISO", "46100", "CONFIG_UNLIMITED_KEYWORD"), Array("ISO", "2026-10-16", "CONFIG_UNLIMITED_KEYWORD"), Array("ISO", ChrW(160) & "2026-10-16", "CONFIG_UNLIMITED_KEYWORD"), Array("ISO", "46100" & ChrW(&H200B), "CONFIG_UNLIMITED_KEYWORD"), Array("DMY_DOT", ChrW(160) & "16.10.2026", "CONFIG_UNLIMITED_KEYWORD"))
        Set result = adapter.Normalize(source, CStr(item(0)), CStr(item(1)))
        AT_Assert result.State = atEvalError, "normalization config fail closed"
        If result.State = atEvalError Then AT_Assert result.Diagnostics(1).Code = CStr(item(2)), "normalization config error code"
        On Error Resume Next
        Set context = result.Context
        caught = Err.Number <> 0: Err.Clear
        On Error GoTo CleanupError
        AT_Assert caught, "config failure exposes no partial context"
    Next item
    ws.Range("A1").Value2 = "old template"
    Set result = adapter.Normalize(reader.Read(ws, "", "bad", ""), "AUTO", "")
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "TEMPLATE_MISMATCH", "source rejection precedes bad normalization policy"
    Set result = adapter.Normalize(Nothing, "ISO", "Бессрочно")
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "NETWORK_INPUT_INVALID", "missing source controlled"
    AT_TestMalformedNormalization source
Finish:
    book.Close False
    If Not saved Is Nothing Then AT_Assert saved.ResolveField("TECHNICAL_ROWS").Value.Row(1).ResolveField("TECH_PROTOCOL").Value.Scalar = "TCP", "typed snapshot detached from source edits and closure"
    Exit Sub
CleanupError:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise number, , description
End Sub
Private Function NormalizeDetails(ByVal result As ATNetworkReadResult) As String
    If result.State <> atEvalOK Then NormalizeDetails = " " & result.Diagnostics(1).Code
End Function
Public Function AT_VerifyNetworkNormalizationReference(ByVal workbookName As String) As String
    Dim reader As New ATNetworkRequestReader, adapter As New ATNetworkRequestNormalizer
    Dim result As ATNetworkReadResult, context As ATEvaluationContext
    Set result = adapter.Normalize(reader.Read(Application.Workbooks(workbookName).Worksheets(1), "REFERENCE", "2026-10-16", "Europe/Moscow"), "DMY_DOT", "Бессрочно")
    If result.State <> atEvalOK Then
        AT_VerifyNetworkNormalizationReference = "ERROR|" & result.Diagnostics(1).Code
    Else
        Set context = result.Context
        AT_VerifyNetworkNormalizationReference = "OK|TECH=" & CStr(context.ResolveField("TECHNICAL_ROWS").Value.Count) & "|ADDR=" & CStr(context.ResolveField("ADDRESS_ROWS").Value.Count) & "|UNLIMITED=" & CStr(context.ResolveField("VALIDITY_IS_UNLIMITED").Value.Scalar)
    End If
End Function
Private Sub AT_TestNormalizedDSL(ByVal context As ATEvaluationContext)
    Dim schema As New ATCompileContext, compiler As New ATCompiler, evaluator As New ATEvaluator
    Dim result As ATEvalResult, entry As ATTraceEntry, located As Boolean
    schema.AddField "TECHNICAL_ROWS", "COLLECTION<ROW<TECHNICAL_TABLE>>"
    schema.AddField "ADDRESS_ROWS", "COLLECTION<ROW<ADDRESS_TABLE>>": schema.AddField "ADDRESS_IP", "IP_EXPRESSION", "ADDRESS_TABLE"
    schema.AddField "TECH_SOURCE_IP", "IP_EXPRESSION", "TECHNICAL_TABLE": schema.AddField "TECH_DEST_PORT", "PORT_SET", "TECHNICAL_TABLE"
    schema.AddField "VALIDITY_DATE", "DATE"
    Set result = evaluator.Evaluate(compiler.Compile("ALLROWS(@TECHNICAL_ROWS AS t, ANYROW(@ADDRESS_ROWS AS a, @t.TECH_SOURCE_IP CONTAINS @a.ADDRESS_IP) AND @t.TECH_DEST_PORT COVER: 443,8443) AND @VALIDITY_DATE >= DATE(""2026-01-01"")", schema), context)
    AT_Assert result.State = atEvalOK, "typed read pipeline feeds domain/date DSL" & AT_RuntimeTestDetails(result)
    If result.State = atEvalOK Then AT_Assert result.Value.Scalar, "typed read pipeline domain/date predicate true"
    For Each entry In result.Trace
        If entry.FieldId = "TECH_DEST_PORT" And entry.CellRef = "'Лист1'!D15" Then located = True
    Next entry
    AT_Assert located, "typed pipeline retains DSL row trace source"
End Sub
Private Sub AT_TestMalformedNormalization(ByVal original As ATNetworkReadResult)
    Dim adapter As New ATNetworkRequestNormalizer, malformed As New ATNetworkReadResult
    Dim context As New ATEvaluationContext, result As ATNetworkReadResult, rows As New Collection
    Dim value As New ATRuntimeValue, row As ATRuntimeRow, i As Long
    context.Initialize "BAD", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    malformed.Initialize atEvalOK, context, Nothing
    Set result = adapter.Normalize(malformed, "ISO", "Бессрочно")
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "NETWORK_INPUT_INVALID", "missing raw schema controlled"
    For i = 1 To 10001
        Set row = New ATRuntimeRow: row.Initialize "TECHNICAL_TABLE_RAW", CStr(i): rows.Add row
    Next i
    value.InitializeRows "TECHNICAL_TABLE_RAW", rows
    Set malformed = AT_OverrideNormalizationInput(original.Context, "TECHNICAL_ROWS_RAW", AT_RuntimeOK(value))
    Set result = adapter.Normalize(malformed, "ISO", "Бессрочно")
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "RUNTIME_LIMIT_LIST", "oversized raw rows rejected before processing"
    Set malformed = AT_OverrideNormalizationInput(original.Context, "TECHNICAL_ROWS_RAW", AT_RuntimeOK(AT_RuntimeScalar("TEXT", "wrong")))
    Set result = adapter.Normalize(malformed, "ISO", "Бессрочно")
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "NETWORK_INPUT_INVALID", "wrong raw table type rejected"
    Set malformed = AT_OverrideNormalizationInput(original.Context, "VALIDITY_VALUE2_KIND", AT_RuntimeOK(AT_RuntimeScalar("TEXT", "UNKNOWN")))
    Set result = adapter.Normalize(malformed, "ISO", "Бессрочно")
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "NETWORK_INPUT_INVALID", "unknown native metadata rejected"
    Set malformed = AT_OverrideNormalizationInput(original.Context, "DB_ACCOUNT_RAW", AT_RuntimeFailure(atEvalError, "CONFIG_TEST", Nothing))
    Set result = adapter.Normalize(malformed, "ISO", "Бессрочно")
    AT_Assert result.State = atEvalError, "raw document system failure rejects whole snapshot"
    Set row = New ATRuntimeRow: row.Initialize "TECHNICAL_TABLE_RAW", "15"
    row.AddField "SOURCE_RAW", AT_RuntimeFailure(atEvalError, "CONFIG_TEST", Nothing)
    row.AddField "DESTINATION_RAW", AT_RuntimeOK(AT_RuntimeScalar("TEXT", "198.51.100.1"))
    row.AddField "PROTOCOL_RAW", AT_RuntimeOK(AT_RuntimeScalar("TEXT", "TCP"))
    row.AddField "PORT_RAW", AT_RuntimeOK(AT_RuntimeScalar("TEXT", "443"))
    row.AddField "REASON_RAW", AT_RuntimeOK(AT_RuntimeScalar("TEXT", "reason"))
    Set rows = New Collection: rows.Add row: Set value = New ATRuntimeValue
    value.InitializeRows "TECHNICAL_TABLE_RAW", rows
    Set malformed = AT_OverrideNormalizationInput(original.Context, "TECHNICAL_ROWS_RAW", AT_RuntimeOK(value))
    Set result = adapter.Normalize(malformed, "ISO", "Бессрочно")
    AT_Assert result.State = atEvalError, "raw row system failure rejects whole snapshot"
End Sub
Private Function AT_OverrideNormalizationInput(ByVal source As ATEvaluationContext, ByVal id As String, ByVal replacement As ATEvalResult) As ATNetworkReadResult
    Dim context As New ATEvaluationContext, spec As ATExcelFieldSpec, specs As Collection
    Dim fieldId As Variant, value As ATEvalResult, result As New ATNetworkReadResult
    context.Initialize source.RunId, source.RunDate, source.TimezoneId, source.LocaleId, source.ComparisonMode
    Set specs = AT_NetworkDocumentFieldsV3()
    For Each spec In specs
        Set value = source.ResolveField(spec.FieldId)
        If spec.FieldId = id Then Set value = replacement
        context.AddField spec.FieldId, value, source.CellRef(spec.FieldId)
    Next spec
    For Each fieldId In Array("TECHNICAL_ROWS_RAW", "ADDRESS_ROWS_RAW", "WORKBOOK_DATE1904", "VALIDITY_VALUE2_KIND", "VALIDITY_SERIAL")
        Set value = source.ResolveField(CStr(fieldId))
        If CStr(fieldId) = id Then Set value = replacement
        context.AddField CStr(fieldId), value
    Next fieldId
    result.Initialize atEvalOK, context, Nothing: Set AT_OverrideNormalizationInput = result
End Function
