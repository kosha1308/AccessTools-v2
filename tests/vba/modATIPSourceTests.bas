Attribute VB_Name = "modATIPSourceTests"
Option Explicit
Public Sub AT_TestIPSource()
    Dim c As New ATIPSourceCompiler, loaded As ATIPSourceResult, q As ATIPMatchResult, index As ATIPSourceIndex
    Dim rows As Variant, v As Variant, b As ATEvalBudget, limits As ATEvalLimits, caught As Boolean
    rows = Array(Array("0.0.0.0/0", "", "whole"), Array("192.0.2.0/24", "192.0.2.1", "network"), Array("192.0.2.1-192.0.2.3", "", "range"), Array("192.0.2.4,5", "", "compact"), Array("203.0.113.5", "", "other"))
    Set loaded = c.Compile(AT_TestIPRaw("SHARED", rows), "SHARED", AT_TestIPFields(), "file|100|mtime1")
    AT_Assert loaded.State = atEvalOK, "shared IP source compiles"
    If loaded.State <> atEvalOK Then Exit Sub
    Set index = loaded.Index
    AT_Assert index.SourceRowCount = 5 And index.SourceId = "SHARED" And index.Fingerprint = "file|100|mtime1", "immutable source metadata"
    Set q = index.Query("192.0.2.2", AT_TestInitiatorBudget())
    AT_Assert q.State = atEvalOK And q.Matches.Count = 3 And q.SourceComplete, "nested intervals preserve broad and narrow matches"
    AT_Assert q.Matches.Row(1).RowId = "2" And q.Matches.Row(2).RowId = "3" And q.Matches.Row(3).RowId = "4", "matches ordered by original source rows"
    AT_Assert q.Matches.Row(2).CellRef("ADDR1") = "'IP source'!A3" And AT_PolicyCell(q.Matches.Row(2), "LABEL", AT_TestInitiatorBudget()) = "network", "all raw fields and source refs retained"
    Set q = index.Query("192.0.2.2;192.0.2.4", AT_TestInitiatorBudget())
    AT_Assert q.State = atEvalOK And q.Matches.Count = 4, "multiple query intervals deduplicate repeated source matches"
    For Each v In Array("0.0.0.0", "255.255.255.255")
        Set q = index.Query(CStr(v), AT_TestInitiatorBudget())
        AT_Assert q.State = atEvalOK And q.Matches.Count = 1, "IPv4 extreme boundaries covered without expansion"
    Next v
    Set q = index.Query("192.0.2.255-192.0.3.0", AT_TestInitiatorBudget())
    AT_Assert q.Matches.Count = 2, "range overlap includes boundary, not full-policy coverage"
    Set q = index.Query("192.0.2.4/32", AT_TestInitiatorBudget())
    AT_Assert q.Matches.Count = 3, "compact source continuation and /32 query"
    On Error Resume Next
    q.Matches.Row(1).AddField "MUTATE", AT_RuntimeOK(AT_RuntimeScalar("TEXT", "bad"))
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught, "matched source rows frozen"
    For Each v In Array("", "group A", "192.0.2.1;bad", "192.0.2.999", "2001:db8::1")
        Set q = index.Query(CStr(v), AT_TestInitiatorBudget())
        AT_Assert q.State = atEvalManualReview, "invalid query yields no partial matches"
    Next v
    Set q = index.Query("192.0.2.1", Nothing)
    AT_Assert q.State = atEvalError, "query budget required"
    Set b = AT_TestInitiatorBudget(): b.Step 200000
    Set q = index.Query("192.0.2.1", b)
    AT_Assert q.State = atEvalError And q.Diagnostics(1).Code = "RUNTIME_LIMIT_STEPS", "exhausted query work fails explicitly"
    Set limits = New ATEvalLimits: limits.MaxListItems = 1: Set b = New ATEvalBudget: b.Initialize limits
    Set q = index.Query("192.0.2.1", b)
    AT_Assert q.State = atEvalError And q.Diagnostics(1).Code = "RUNTIME_LIMIT_LIST", "too many matches cannot truncate"
    rows = Array(Array("192.0.2.1;bad", "203.0.113.1", "bad cell valid second"), Array(CVErr(xlErrValue), "", "cell error"), Array("", "", "blank addresses"))
    Set loaded = c.Compile(AT_TestIPRaw("SHARED", rows), "SHARED", AT_TestIPFields(), "changed")
    AT_Assert loaded.State = atEvalOK And loaded.Index.IssueCount = 2, "unparsed source cells explicit while valid cells retained"
    Set q = loaded.Index.Query("192.0.2.1", AT_TestInitiatorBudget())
    AT_Assert q.State = atEvalOK And q.Matches.Count = 0 And Not q.SourceComplete, "partial malformed cell never salvaged or silently complete"
    Set q = loaded.Index.Query("203.0.113.1", AT_TestInitiatorBudget())
    AT_Assert q.Matches.Count = 1 And Not q.SourceComplete, "valid second cell still matches with incomplete-source flag"
    AT_Assert loaded.Index.DataIssues.Value.Row(1).CellRef("VALUE") = "'IP source'!A2", "source data issue points to actual bad cell"
    AT_Assert AT_PolicyCell(loaded.Index.DataIssues.Value.Row(1), "VALUE", AT_TestInitiatorBudget()) = "192.0.2.1;bad", "bad original value retained whole"
    Set loaded = c.Compile(AT_TestIPRaw("SHARED", Array()), "SHARED", AT_TestIPFields(), "empty")
    Set q = loaded.Index.Query("192.0.2.1", AT_TestInitiatorBudget())
    AT_Assert q.State = atEvalOK And q.Matches.Count = 0 And q.SourceComplete, "valid empty source has explicit empty matches"
    AT_Assert c.Compile(Nothing, "SHARED", AT_TestIPFields(), "x").State = atEvalError, "missing source cannot compile"
    AT_Assert c.Compile(AT_TestIPRaw("WRONG", Array()), "SHARED", AT_TestIPFields(), "x").State = atEvalError, "source type validated"
    AT_Assert c.Compile(AT_TestIPRaw("SHARED", Array()), "SHARED", AT_TestIPFields(), "").State = atEvalError, "source fingerprint mandatory"
    AT_TestIPSourceSchema c
    AT_TestIPIndexCorpus
    AT_TestSupersetIPIndex
End Sub
Private Sub AT_TestIPIndexCorpus()
    Dim records(0 To 63) As Variant, lo(0 To 63) As Long, hi(0 To 63) As Long, i As Long, x As Long, expected As Long
    Dim c As New ATIPSourceCompiler, loaded As ATIPSourceResult, q As ATIPMatchResult
    For i = 0 To 63
        lo(i) = (i * 7) Mod 32: hi(i) = lo(i) + (i Mod 5)
        records(i) = Array(AT_IPText(lo(i)) & "-" & AT_IPText(hi(i)), "", CStr(i))
    Next i
    Set loaded = c.Compile(AT_TestIPRaw("SHARED", records), "SHARED", AT_TestIPFields(), "corpus")
    For x = 0 To 36
        expected = 0
        Set q = loaded.Index.Query(AT_IPText(x), AT_TestInitiatorBudget())
        For i = 0 To 63
            If lo(i) <= x And hi(i) >= x Then
                expected = expected + 1
                AT_Assert q.Matches.Row(expected).RowId = CStr(i + 2), "indexed overlapping interval corpus preserves every source row in order"
            End If
        Next i
        AT_Assert q.Matches.Count = expected, "indexed overlap equals finite brute-force reference"
    Next x
    Set loaded = c.Compile(AT_TestIPRaw("SHARED", Array(Array("192.0.x.1", "", "mask"))), "SHARED", AT_TestIPFields(), "mask")
    AT_Assert loaded.State = atEvalOK And loaded.Index.IntervalCount = 256, "nontrailing octet mask remains compact interval index"
    Set q = loaded.Index.Query("192.0.255.1", AT_TestInitiatorBudget())
    AT_Assert q.Matches.Count = 1, "nontrailing mask final endpoint"
    Set q = loaded.Index.Query("192.0.255.2", AT_TestInitiatorBudget())
    AT_Assert q.Matches.Count = 0, "nontrailing mask excludes gaps"
End Sub
Private Sub AT_TestSupersetIPIndex()
    Dim book As Workbook, ws As Worksheet, reader As New ATNetworkSourceReader, c As New ATIPSourceCompiler
    Dim loaded As ATIPSourceResult, q As ATIPMatchResult, fields As New Collection, n As Long, description As String
    On Error GoTo Failed
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Superset test"
    ws.Range("B1:I1").Value2 = Array("FQDN", "SOURCE:ТЕЛЕГРАФ", "SOURCE: ЦХД", "SOURCE: AWX", "ПОЛНОЕ ИМЯ ИС", "СЕРВИС", "СКУФ ГРУППА ТП", "КОМАНДА ПОДДЕРЖКИ")
    ws.Range("B2:I2").Value2 = Array("host.test", "192.0.2.1", "192.0.2.2", "192.0.2.3", "IS", "Service", "SKUF", "Support")
    ws.Range("B3:C3").Value2 = Array("hidden.test", "203.0.113.1")
    ws.Range("B1:I3").AutoFilter Field:=1, Criteria1:="host.test"
    fields.Add "SOURCE_TELEGRAPH": fields.Add "SOURCE_CXD": fields.Add "SOURCE_AWX"
    Set loaded = c.Compile(reader.Read(ws, "SUPERSET"), "SUPERSET", fields, "provided schema synthetic fixture")
    Set q = loaded.Index.Query("192.0.2.3", AT_TestInitiatorBudget())
    AT_Assert q.Matches.Count = 1 And q.Matches.Row(1).CellRef("SOURCE_AWX") = "'Superset test'!E2", "actual third SOURCE column indexed"
    AT_Assert AT_PolicyCell(q.Matches.Row(1), "IS_NAME", AT_TestInitiatorBudget()) = "IS" And AT_PolicyCell(q.Matches.Row(1), "SKUF_GROUP", AT_TestInitiatorBudget()) = "SKUF" And AT_PolicyCell(q.Matches.Row(1), "SUPPORT_TEAM", AT_TestInitiatorBudget()) = "Support", "source payload available for final enrichment"
    Set q = loaded.Index.Query("203.0.113.1", AT_TestInitiatorBudget())
    AT_Assert q.Matches.Count = 1 And ws.FilterMode And ws.Rows(3).Hidden, "source index includes filtered rows without mutating filter"
    book.Close False: Set book = Nothing: Exit Sub
Failed:
    n = Err.Number: description = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise n, "AT_TestSupersetIPIndex", description
End Sub
Private Sub AT_TestIPSourceSchema(ByVal c As ATIPSourceCompiler)
    Dim fields As New Collection, row As New ATRuntimeRow, rows As New Collection, value As New ATRuntimeValue
    Dim loaded As ATIPSourceResult
    fields.Add "MISSING"
    Set loaded = c.Compile(AT_TestIPRaw("SHARED", Array(Array("192.0.2.1", "", "x"))), "SHARED", fields, "schema")
    AT_Assert loaded.State = atEvalError, "missing address field is structural, not silently incomplete source"
    row.Initialize "SHARED_RAW", "2": row.AddField "ADDR1", AT_RuntimeOK(AT_RuntimeScalar("NUMBER", CDbl(123)))
    row.AddField "ADDR2", AT_RuntimeOK(AT_RuntimeState("TEXT", atValueEmpty)): rows.Add row
    value.InitializeRows "SHARED_RAW", rows
    Set loaded = c.Compile(AT_RuntimeOK(value), "SHARED", AT_TestIPFields(), "type")
    AT_Assert loaded.State = atEvalError, "non-text address field rejected structurally"
    Set fields = New Collection: fields.Add "ADDR1": fields.Add "addr1"
    Set loaded = c.Compile(AT_TestIPRaw("SHARED", Array()), "SHARED", fields, "duplicates")
    AT_Assert loaded.State = atEvalError, "duplicate address-field metadata rejected"
End Sub
Public Sub AT_TestIPSourceCache()
    Dim cache As New ATIPSourceCache, a As ATIPSourceResult, b As ATIPSourceResult, raw As ATEvalResult, i As Long
    AT_Assert cache.Count = 0, "cache startup has no source work"
    Set raw = AT_TestIPRaw("SHARED", Array(Array("192.0.2.1", "", "first")))
    Set a = cache.GetOrCompile(raw, "SHARED", AT_TestIPFields(), "same")
    AT_Assert a.State = atEvalOK, "cache compiles first consumer"
    If a.State <> atEvalOK Then Exit Sub
    Set b = cache.GetOrCompile(raw, "SHARED", AT_TestIPFields(), "same")
    AT_Assert a.Index Is b.Index, "NETWORK_ACCESS and VPN reuse exactly one immutable index"
    Set b = cache.GetOrCompile(AT_TestIPRaw("SHARED", Array(Array("203.0.113.1", "", "new"))), "SHARED", AT_TestIPFields(), "new size or mtime")
    AT_Assert Not a.Index Is b.Index And b.Index.Query("192.0.2.1", AT_TestInitiatorBudget()).Matches.Count = 0, "changed source fingerprint invalidates old contents"
    Set b = cache.GetOrCompile(Nothing, "SHARED", AT_TestIPFields(), "bad replacement")
    AT_Assert b.State = atEvalError And cache.Count = 0, "bad replacement cannot silently reuse stale index"
    Dim fields As New Collection
    Set raw = AT_TestIPRaw("SHARED", Array(Array("192.0.2.1", "203.0.113.1", "schema")))
    Set a = cache.GetOrCompile(raw, "SHARED", AT_TestIPFields(), "schema")
    fields.Add "ADDR2"
    Set b = cache.GetOrCompile(raw, "SHARED", fields, "schema")
    AT_Assert Not a.Index Is b.Index And b.Index.Query("192.0.2.1", AT_TestInitiatorBudget()).Matches.Count = 0, "address schema participates in cache identity"
    For i = 1 To 6
        Set b = cache.GetOrCompile(AT_TestIPRaw("S" & i, Array()), "S" & i, AT_TestIPFields(), "x")
        AT_Assert b.State = atEvalOK, "bounded cache accepts different registered sources"
    Next i
    AT_Assert cache.Count = 4, "cache memory bounded by four source/schema entries"
    AT_TestIPSourceScale
End Sub
Private Sub AT_TestIPSourceScale()
    Dim records(0 To 10000) As Variant, i As Long, c As New ATIPSourceCompiler, loaded As ATIPSourceResult
    Dim limits As New ATEvalLimits, budget As New ATEvalBudget, q As ATIPMatchResult
    For i = 0 To 10000: records(i) = Array(AT_IPText(i + 1), "", "row" & CStr(i + 1)): Next i
    Set loaded = c.Compile(AT_TestIPRaw("SHARED", records), "SHARED", AT_TestIPFields(), "scale10001")
    AT_Assert loaded.State = atEvalOK And loaded.Index.SourceRowCount = 10001, "source indexing supports more than generic ten-thousand-list default"
    limits.MaxSteps = 1000: budget.Initialize limits
    Set q = loaded.Index.Query(AT_IPText(10001), budget)
    AT_Assert q.State = atEvalOK And q.Matches.Count = 1 And q.Matches.Row(1).RowId = "10002", "indexed tail search fits budget too small for full source scan"
End Sub
Public Function AT_TestIPFields() As Collection
    Dim r As New Collection: r.Add "ADDR1": r.Add "ADDR2": Set AT_TestIPFields = r
End Function
Public Function AT_TestIPRaw(ByVal kind As String, ByVal records As Variant) As ATEvalResult
    Dim row As ATRuntimeRow, rows As New Collection, fields As Variant, record As Variant, i As Long, number As Long
    Dim value As New ATRuntimeValue, field As ATEvalResult
    fields = Array("ADDR1", "ADDR2", "LABEL")
    For Each record In records
        number = number + 1: Set row = New ATRuntimeRow: row.Initialize kind & "_RAW", CStr(number + 1)
        For i = 0 To 2
            If IsError(record(i)) Then
                Set field = AT_RuntimeFailure(atEvalError, "EXCEL_CELL_ERROR", Nothing)
            ElseIf Len(CStr(record(i))) = 0 Then
                Set field = AT_RuntimeOK(AT_RuntimeState("TEXT", atValueEmpty))
            Else
                Set field = AT_RuntimeOK(AT_RuntimeScalar("TEXT", CStr(record(i))))
            End If
            row.AddField CStr(fields(i)), field, "'IP source'!" & Chr$(65 + i) & CStr(number + 1)
        Next i
        rows.Add row
    Next record
    value.InitializeRows kind & "_RAW", rows: Set AT_TestIPRaw = AT_RuntimeOK(value)
End Function
