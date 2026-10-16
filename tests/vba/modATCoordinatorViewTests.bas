Attribute VB_Name = "modATCoordinatorViewTests"
Option Explicit
Public Sub AT_TestCoordinatorView()
    Dim sourceBook As Workbook, ws As Worksheet, output As Workbook, summary As Worksheet, view As Worksheet
    Dim registry As ATNetworkRegistryResult, planner As New ATNetworkRegistryPlanner, reader As New ATNetworkRequestReader
    Dim compiler As New ATIPSourceCompiler, loaded As ATIPSourceResult, writer As New ATCoordinatorWorkbookWriter
    Dim savedNumber As Long, savedDescription As String, records As Variant, i As Long, previousEvents As Boolean
    Dim enricher As New ATSupersetEnricher, enriched As ATEvalResult, spec As ATExcelTableSpec, ids As New Collection
    Dim bookCount As Long, caught As Boolean, code As String
    On Error GoTo Failed
    Set sourceBook = Application.Workbooks.Add(xlWBATWorksheet): Set ws = sourceBook.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    ws.Range("A15:D15").Value2 = Array("192.0.2.1", "192.0.2.0/24", "TCP", "443")
    ws.Range("B17").Value2 = "192.0.2.1": ws.Range("C17").Value2 = "absent.test"
    ws.Range("B18").Value2 = "203.0.113.1"
    records = Array(Array("alpha.test", "192.0.2.1", "", "", "IS A", "Svc A", "Group A", "Team A"), Array("beta.test", "192.0.2.1", "", "", "IS B", "Svc B", "Group B", "Team B"), Array("alpha.test", "192.0.2.2", "", "", "IS A", "Svc A", "Group A", "=1+1"))
    Set loaded = compiler.Compile(AT_TestRegistryRaw(records), "SUPERSET", AT_TestRegistryFields(), "view fixture")
    Set registry = planner.Run(reader.Read(ws, "VIEW", "2026-10-16", "UTC"), loaded)
    Set output = writer.Create(registry)
    AT_Assert Not output Is Nothing, "coordinator writer creates output workbook"
    If output Is Nothing Then GoTo Cleanup
    Set summary = output.Worksheets(1): Set view = output.Worksheets(2)
    AT_Assert output.Worksheets.Count = 2 And summary.Name = "Сводка координатора" And view.Name = "Superset", "coordinator output tab order"
    AT_Assert output.Names("ATVIEW_SUPERSET").RefersToRange.Address = "$A$1", "temporary result marked ATVIEW_SUPERSET"
    AT_Assert view.Cells(8, 1).Value2 = "2" And Len(view.Cells(9, 1).Value2) = 0, "physical source rows deduplicated not repeated per summary"
    AT_Assert view.Cells(6, 1).Value2 = "1,2" And view.Cells(7, 1).Value2 = "1,2", "overlap source row preserves every summary link"
    AT_Assert view.Cells(6, 2).Value2 = "alpha.test" And view.Cells(7, 2).Value2 = "beta.test", "matched original business rows stay in source order"
    AT_Assert summary.Cells(6, 4).Value2 = 2 And summary.Cells(7, 4).Value2 = 3 And summary.Cells(8, 4).Value2 = 0, "summary counts actual linked output rows"
    AT_Assert summary.Cells(6, 8).Value2 = "Svc A" & vbLf & "Svc B", "summary service comes from all actual linked rows"
    AT_Assert summary.Cells(6, 9).Value2 = "IS A" & vbLf & "IS B", "summary preserves multiple systems"
    AT_Assert summary.Cells(6, 10).Value2 = "Group A" & vbLf & "Group B", "summary SKUF groups preserved"
    AT_Assert summary.Cells(6, 11).Value2 = "Team A" & vbLf & "Team B", "summary support teams preserved"
    AT_Assert Len(summary.Cells(6, 7).Value2) = 0, "ambiguous hostname not chosen arbitrarily"
    AT_Assert summary.Cells(8, 3).Value2 = "НЕ НАЙДЕН В SUPERSET" And Len(summary.Cells(8, 8).Value2) = 0, "missing address retained with empty source enrichment"
    AT_Assert InStr(summary.Cells(6, 12).Value2, "A15") > 0 And InStr(summary.Cells(6, 12).Value2, "B17") > 0, "summary preserves both request references"
    AT_Assert Not view.Cells(8, 9).HasFormula And view.Cells(8, 9).Value2 = "=1+1", "formula-like source text safely retained as literal"
    AT_Assert InStr(summary.Cells(3, 1).Value2, "ещё не") > 0 And Len(summary.Cells(6, 5).Value2) = 0, "pending rule annotations explicitly labelled rather than fabricated"
    AT_Assert ws.Range("A15").Value2 = "192.0.2.1" And ws.Range("B18").Value2 = "203.0.113.1", "request untouched by writer"
    AT_Assert summary.AutoFilterMode And view.AutoFilterMode And Not output.Saved, "output filter controls and explicit save required"
    AT_Assert summary.Cells(8, 3).Interior.Color = RGB(255, 242, 204), "missing result keeps warning highlight after row striping"
    ids.Add "1": Set spec = AT_ViewSpec()
    view.Rows(7).Hidden = True
    Set enriched = enricher.Read(view, spec.Locate(view), 8, ids)
    AT_Assert enriched.Value.Row(1).ResolveField("SERVICE").Value.Scalar = "Svc A", "actual output visibility controls enrichment without snapshot shortcuts"
    output.Close False: Set output = Nothing
    ws.Range("A15").Value2 = "192.0.2.1,2,3,4,5,6,7,8,9,10": ws.Range("B15").ClearContents: ws.Range("B18").ClearContents
    records(2) = Array("only10.test", "192.0.2.10", "", "", "IS 10", "Svc 10", "Group 10", "+literal")
    Set loaded = compiler.Compile(AT_TestRegistryRaw(records), "SUPERSET", AT_TestRegistryFields(), "exact links")
    Set registry = planner.Run(reader.Read(ws, "TEN", "2026-10-16", "UTC"), loaded)
    Set output = writer.Create(registry)
    AT_Assert output.Worksheets(1).Cells(6, 4).Value2 = 2 And output.Worksheets(1).Cells(15, 4).Value2 = 1, "summary1 never consumes summary10 source links"
    AT_Assert output.Worksheets(1).Cells(15, 7).Value2 = "only10.test", "unambiguous source hostname displayed"
    AT_Assert output.Worksheets(2).Cells(8, 9).Value2 = "+literal" And Not output.Worksheets(2).Cells(8, 9).HasFormula, "plus-prefixed source literal safe"
    output.Close False: Set output = Nothing
    Set registry = planner.Run(reader.Read(ws, "NO_SOURCE", "2026-10-16", "UTC"), Nothing)
    Set output = writer.Create(registry)
    AT_Assert output.Worksheets(1).Cells(6, 3).Value2 = "НЕ ПОДТВЕРЖДЕНО" And Len(output.Worksheets(2).Cells(6, 1).Value2) = 0, "unavailable source produces unconfirmed summary not false not-found"
    output.Close False: Set output = Nothing
    records(0) = Array("bad-service.test", "192.0.2.1", "", "", "IS", CVErr(xlErrValue), "", "")
    Set loaded = compiler.Compile(AT_TestRegistryRaw(records), "SUPERSET", AT_TestRegistryFields(), "source field error")
    Set registry = planner.Run(reader.Read(ws, "ERROR", "2026-10-16", "UTC"), loaded)
    bookCount = Application.Workbooks.Count: previousEvents = Application.EnableEvents
    Application.EnableEvents = False
    On Error Resume Next
    Set output = writer.Create(registry): code = Err.Description: caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    AT_Assert caught And code = "EXCEL_CELL_ERROR" And output Is Nothing, "source service error blocks incomplete final workbook"
    AT_Assert Application.Workbooks.Count = bookCount And Not Application.EnableEvents, "failed writer closes only own book and restores disabled events"
    Application.EnableEvents = previousEvents
    Set registry = planner.Run(Nothing, loaded)
    On Error Resume Next
    Set output = writer.Create(registry): caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    AT_Assert caught And Application.Workbooks.Count = bookCount, "blocked registry cannot create partial result workbook"
    ws.Range("A15").ClearContents: ws.Range("B17").ClearContents
    Set registry = planner.Run(reader.Read(ws, "EMPTY", "2026-10-16", "UTC"), Nothing)
    Set output = writer.Create(registry)
    AT_Assert Len(output.Worksheets(1).Cells(6, 1).Value2) = 0 And output.Worksheets.Count = 2, "empty address plan creates valid header-only result"
    output.Close False: Set output = Nothing
    AT_TestViewLimits writer
Cleanup:
    If Not output Is Nothing Then output.Close False
    sourceBook.Close False: Exit Sub
Failed:
    savedNumber = Err.Number: savedDescription = Err.Description
    On Error Resume Next
    If Not output Is Nothing Then output.Close False
    If Not sourceBook Is Nothing Then sourceBook.Close False
    On Error GoTo 0
    Err.Raise savedNumber, "AT_TestCoordinatorView", savedDescription
End Sub
Private Sub AT_TestViewLimits(ByVal writer As ATCoordinatorWorkbookWriter)
    Dim rawRows As New Collection, row As ATRuntimeRow, raw As New ATRuntimeValue, entry As New ATNetworkRegistryEntry
    Dim parsed As ATEvalResult, budget As New ATEvalBudget, limits As New ATEvalLimits, registry As ATNetworkRegistryResult
    Dim entries As New Collection, report As New ATCheckReport, issues As New Collection, checks As New Collection
    Dim output As Workbook, incomplete As Boolean, i As Long, caught As Boolean, code As String, bookCount As Long
    Set parsed = AT_ParseDomain("IP", "192.0.2.1", Nothing)
    limits.MaxListItems = 20000: budget.Initialize limits
    For i = 1 To 10001
        Set row = New ATRuntimeRow: row.Initialize "SUPERSET_RAW", CStr(i + 1): rawRows.Add row
    Next i
    raw.InitializeRows "SUPERSET_RAW", rawRows
    entry.Initialize "192.0.2.1", "bounded", parsed.Value.Domain
    entry.AddReference "ADDRESS_IP", "'Лист1'!B17", budget
    entry.Finish "1", raw, True, incomplete, budget: entries.Add entry
    report.Initialize atEvalOK, "NETWORK_REGISTRY", issues, checks
    Set registry = New ATNetworkRegistryResult: registry.Initialize atEvalOK, entries, report
    bookCount = Application.Workbooks.Count
    On Error Resume Next
    Set output = writer.Create(registry): caught = Err.Number <> 0: code = Err.Description: Err.Clear
    On Error GoTo 0
    AT_Assert caught And code = "VIEW_ROW_LIMIT" And Application.Workbooks.Count = bookCount, "over10000 unique source rows explicitly blocked without truncated output"
    Set registry = AT_TestViewOneRow(AT_TestRegistryRaw(Array(Array("host", "192.0.2.1", "", "", "IS", String$(32768, "x"), "", ""))).Value)
    On Error Resume Next
    Set output = writer.Create(registry): caught = Err.Number <> 0: code = Err.Description: Err.Clear
    On Error GoTo 0
    AT_Assert caught And code = "VIEW_TEXT_LIMIT" And Application.Workbooks.Count = bookCount, "overlong source cell cannot truncate silently"
    Set registry = AT_TestViewOneRow(AT_TestRegistryRaw(Array(Array("host", "192.0.2.1", "", "", "IS", "'literal", "", ""))).Value)
    Set output = writer.Create(registry)
    AT_Assert output.Worksheets(2).Cells(6, 7).Value2 = "'literal" And output.Worksheets(1).Cells(6, 8).Value2 = "'literal", "leading apostrophe preserved through worksheet enrichment roundtrip"
    output.Close False
End Sub
Private Function AT_TestViewOneRow(ByVal raw As ATRuntimeValue) As ATNetworkRegistryResult
    Dim entry As New ATNetworkRegistryEntry, parsed As ATEvalResult, entries As New Collection
    Dim report As New ATCheckReport, issues As New Collection, checks As New Collection, incomplete As Boolean
    Dim registry As New ATNetworkRegistryResult, budget As ATEvalBudget
    Set parsed = AT_ParseDomain("IP", "192.0.2.1", Nothing): Set budget = AT_TestInitiatorBudget()
    entry.Initialize "192.0.2.1", "test", parsed.Value.Domain
    entry.AddReference "ADDRESS_IP", "'Лист1'!B17", budget
    entry.Finish "1", raw, True, incomplete, budget: entries.Add entry
    report.Initialize atEvalOK, "NETWORK_REGISTRY", issues, checks
    registry.Initialize atEvalOK, entries, report: Set AT_TestViewOneRow = registry
End Function
