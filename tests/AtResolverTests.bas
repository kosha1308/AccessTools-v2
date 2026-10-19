Attribute VB_Name = "AtResolverTests"
Option Explicit

Private passed As Long, failed As Long, transcript As String
Private firstBook As Object, secondBook As Object

Public Function AtRunResolverTests(ByVal repo As String, ByVal scratch As String) As String
    On Error GoTo Broken
    passed = 0: failed = 0: transcript = ""
    Dim resolver As Object, p As CAtPackage
    Set p = AtNewPackage("RESOLVER_TEST", "Resolver test")
    p.Definitions.Add AtNewDefinition("Template", "SIMPLE", "Simple")
    ' Late-bound bootstrap makes a missing implementation an observable RED.
    Set resolver = Application.Run("'" & ThisWorkbook.Name & "'!AtNewExcelResolver", p, "SIMPLE", AtDictionary())
    Set firstBook = Application.Workbooks.Add
    firstBook.Worksheets(1).Name = "Request"
    Set secondBook = Application.Workbooks.Add
    secondBook.Worksheets(1).Name = "Return"
    TestScalars
    TestNames
    TestLabels
    TestTables
    TestSnapshotBudgets
    TestIntegration scratch
    TestSafety scratch
    Cleanup
    Dim file As Object
    Set file = CreateObject("Scripting.FileSystemObject").CreateTextFile(scratch & "\resolver-assertions.txt", True, True)
    file.Write transcript: file.Close
    AtRunResolverTests = IIf(failed = 0, "PASS|Resolver|", "FAIL|Resolver|") & passed & "|" & failed
    If failed > 0 Then AtRunResolverTests = AtRunResolverTests & "|" & transcript
    Exit Function
Broken:
    AtRunResolverTests = "FAIL|Resolver|" & passed & "|" & Err.Number & "|" & Err.Description & "|" & transcript
    Cleanup
End Function

Private Sub Cleanup()
    On Error Resume Next
    If Not firstBook Is Nothing Then firstBook.Close False
    If Not secondBook Is Nothing Then secondBook.Close False
    Set firstBook = Nothing: Set secondBook = Nothing
End Sub

Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    If ok Then passed = passed + 1 Else failed = failed + 1
    transcript = transcript & IIf(ok, "PASS ", "FAIL ") & label & vbCrLf
End Sub

Private Function Def(ByVal p As CAtPackage, ByVal id As String) As CAtDefinition
    Dim d As CAtDefinition
    For Each d In p.Definitions
        If d.Id = id Then Set Def = d: Exit Function
    Next d
    Err.Raise 5, , "Missing fixture definition: " & id
End Function

Private Function Model(Optional ByVal equipment As Boolean = False) As CAtPackage
    Dim p As CAtPackage, d As CAtDefinition, n As CAtNode, template As String, sheet As String
    template = IIf(equipment, "EQUIPMENT_RETURN", "SIMPLE"): sheet = IIf(equipment, "Return", "Request")
    Set p = AtNewPackage(template & "_PACKAGE", template): p.SchemaVersion = "2.0.0"
    p.Definitions.Add AtNewDefinition("Template", template, template)
    AddField p, template, sheet, "FIELD_TEXT", "Text", "B2"
    AddField p, template, sheet, "FIELD_FLAG", "Bool", "B3"
    Set d = AtNewDefinition("Expression", "EXPR_READY", "Ready")
    d.Values.Add "ReturnType", "Bool": d.Values.Add "Deferred", "false"
    d.ScopeTemplateId = template: d.RootNodeId = "ROOT"
    Set n = AtNewNode("ROOT", "FieldRef", "Bool")
    n.Values.Add "FieldId", "FIELD_FLAG": n.Values.Add "FieldVersion", "1.0.0"
    d.Nodes.Add n: p.Definitions.Add d
    Set d = AtNewDefinition("Rule", "RULE_READY", "Synthetic ready rule")
    AtAddReference d, "OwnerTemplate", "Template", template, "1.0.0"
    AtAddReference d, "Expression", "Expression", "EXPR_READY", "1.0.0", "Bool"
    p.Definitions.Add d
    If equipment Then
        Set d = AtNewDefinition("Locator", "LOC_ITEMS", "Items bounds")
        d.Values.Add "Strategy", "A1": d.Values.Add "Sheet", sheet: d.Values.Add "Address", "A5:C8"
        AtAddReference d, "OwnerTemplate", "Template", template, "1.0.0": p.Definitions.Add d
        Set d = AtNewDefinition("Table", "ITEMS", "Items")
        AtAddReference d, "OwnerTemplate", "Template", template, "1.0.0"
        AtAddReference d, "Locator", "Locator", "LOC_ITEMS", "1.0.0": p.Definitions.Add d
        AddColumn p, "ITEM_CODE", "Text": AddColumn p, "ITEM_COUNT", "Int64": AddColumn p, "ITEM_READY", "Bool"
    End If
    Set Model = p
End Function

Private Sub AddField(ByVal p As CAtPackage, ByVal template As String, ByVal sheet As String, ByVal id As String, ByVal typeName As String, ByVal address As String)
    Dim d As CAtDefinition
    Set d = AtNewDefinition("Locator", "LOC_" & id, id & " location")
    d.Values.Add "Strategy", "A1": d.Values.Add "Sheet", sheet: d.Values.Add "Address", address
    AtAddReference d, "OwnerTemplate", "Template", template, "1.0.0": p.Definitions.Add d
    Set d = AtNewDefinition("Field", id, id): d.Values.Add "DataType", typeName
    AtAddReference d, "OwnerTemplate", "Template", template, "1.0.0"
    AtAddReference d, "Locator", "Locator", "LOC_" & id, "1.0.0": p.Definitions.Add d
End Sub

Private Sub AddColumn(ByVal p As CAtPackage, ByVal id As String, ByVal typeName As String)
    Dim d As CAtDefinition: Set d = AtNewDefinition("Column", id, id)
    d.Values.Add "DataType", typeName
    AtAddReference d, "OwnerTable", "Table", "ITEMS", "1.0.0": p.Definitions.Add d
End Sub

Private Function Bindings() As Object
    Dim b As Object: Set b = AtDictionary()
    b.Add "ITEM_CODE", 1: b.Add "ITEM_COUNT", 2: b.Add "ITEM_READY", 3
    Set Bindings = b
End Function

Private Function Reader(ByVal p As CAtPackage, Optional ByVal equipment As Boolean = False, Optional ByVal bindings As Object = Nothing) As CAtExcelResolver
    If bindings Is Nothing Then Set bindings = AtDictionary()
    Set Reader = AtNewExcelResolver(p, IIf(equipment, "EQUIPMENT_RETURN", "SIMPLE"), bindings)
End Function

Private Function ReadText(ByVal p As CAtPackage) As CAtEvalResult
    Dim resolver As CAtExcelResolver: Set resolver = Reader(p)
    Set ReadText = resolver.ReadField(firstBook, "FIELD_TEXT")
End Function

Private Sub Strategy(ByVal p As CAtPackage, ByVal strategyName As String)
    Dim d As CAtDefinition: Set d = Def(p, "LOC_FIELD_TEXT")
    Set d.Values = AtDictionary(): d.Values.Add "Strategy", strategyName: d.Values.Add "Sheet", "Request"
End Sub

Private Sub TestScalars()
    Dim p As CAtPackage, r As CAtEvalResult, resolver As CAtExcelResolver, s As Object, value As Variant, t As Variant
    Set p = Model(): Set s = firstBook.Worksheets("Request")
    s.Range("B2").Value2 = "  " & ChrW(&H418) & ChrW(&H43C) & ChrW(&H44F) & "  "
    Set r = ReadText(p)
    Check r.State = "OK" And r.Value = s.Range("B2").Value2, "Text preserves Unicode and whitespace"
    Check r.Provenance.Count >= 4 And r.Dependency = "FIELD_TEXT", "field provenance records workbook locator requested and canonical cell"
    s.Range("B2").ClearContents: Check ReadText(p).State = "EMPTY", "existing empty cell is EMPTY"
    s.Range("B2").Formula = "=""""": s.Range("B2").Calculate
    Check ReadText(p).State = "EMPTY", "formula empty string is EMPTY"
    s.Range("B2").Formula = "=""cached""": s.Range("B2").Calculate
    Set r = ReadText(p): Check r.State = "OK" And r.Value = "cached", "formula result is read without formula execution by Resolver"
    Check r.Provenance.Count = 6, "formula result provenance is explicit"
    s.Range("B2").Formula = "=1/0": s.Range("B2").Calculate
    Set r = ReadText(p): Check r.State = "INVALID" And r.Diagnostics(1)("Code") = "R_EXCEL_ERROR", "formula error is INVALID with diagnostic"
    s.Range("B2").Value2 = 12
    Check ReadText(p).State = "INVALID", "numeric cell is not silently coerced to Text"
    Def(p, "FIELD_TEXT").Values("DataType") = "Int64"
    Set r = ReadText(p): Check r.State = "OK" And r.Value = "12", "exact numeric Int64 codec"
    s.Range("B2").Value2 = 1.5: Check ReadText(p).State = "INVALID", "fractional numeric Int64 rejected"
    s.Range("B2").Value2 = 9007199254740992#: Check ReadText(p).State = "INVALID", "unsafe numeric Int64 precision rejected"
    s.Range("B2").NumberFormat = "@": s.Range("B2").Value2 = "9223372036854775807"
    Set r = ReadText(p): Check r.State = "OK" And r.Value = "9223372036854775807", "canonical Int64 text preserves full precision"
    s.Range("B2").Value2 = "9223372036854775808": Check ReadText(p).State = "INVALID", "Int64 overflow text rejected"
    Def(p, "FIELD_TEXT").Values("DataType") = "Bool": s.Range("B2").NumberFormat = "General": s.Range("B2").Value2 = True
    Set r = ReadText(p): Check r.State = "OK" And r.Value = "true", "native Boolean maps explicitly to canonical Bool"
    s.Range("B2").Value2 = 1: Check ReadText(p).State = "INVALID", "numeric one is not Bool"
    s.Range("B2").NumberFormat = "@": s.Range("B2").Value2 = "TRUE": Check ReadText(p).State = "INVALID", "noncanonical Boolean text rejected"
    s.Range("B2").Value2 = "true": Check ReadText(p).Value = "true", "canonical Boolean text accepted"
    For Each t In Array("Decimal", "Date", "DateTime")
        Def(p, "FIELD_TEXT").Values("DataType") = CStr(t)
        s.Range("B2").NumberFormat = "@"
        Select Case t
            Case "Decimal": value = "12.5"
            Case "Date": value = "2026-10-19"
            Case "DateTime": value = "2026-10-19T10:20:30Z"
        End Select
        s.Range("B2").Value2 = value: Set r = ReadText(p)
        Check r.State = "OK" And r.Value = value, "canonical text codec " & t
        s.Range("B2").Value2 = "not-a-value": Check ReadText(p).State = "INVALID", "invalid canonical text " & t
        s.Range("B2").NumberFormat = "General": s.Range("B2").Value2 = 45000
        Check ReadText(p).State = "INVALID", "numeric/serial value not guessed for " & t
    Next t
    For Each t In Array("Email", "Hostname", "IpSet", "PortSet")
        Def(p, "FIELD_TEXT").Values("DataType") = CStr(t): s.Range("B2").Value2 = "example"
        Set r = ReadText(p)
        Check r.State = "INVALID" And r.Diagnostics(1)("Code") = "R_CODEC_UNSUPPORTED", "unsupported codec remains typed " & t
        Check r.ResultType = t And r.Value = "", "unsupported codec does not substitute Text " & t
    Next t
    Set p = Model(): s.Range("B2").Value2 = "anchor": s.Range("B2:C3").Merge
    Def(p, "LOC_FIELD_TEXT").Values("Address") = "C3"
    Check ReadText(p).Value = "anchor", "merged non-anchor scalar reads top left"
    Def(p, "LOC_FIELD_TEXT").Values("Address") = "B2:C3"
    Check ReadText(p).Value = "anchor", "whole MergeArea is one scalar"
    Def(p, "LOC_FIELD_TEXT").Values("Address") = "C2:C3"
    Check ReadText(p).Value = "anchor", "subrectangle inside one MergeArea reads top left"
    Def(p, "LOC_FIELD_TEXT").Values("Address") = "A2:C3"
    Check ReadText(p).State = "CONFLICT", "scalar rectangle spanning other cells is ambiguous"
    s.Range("B2:C3").UnMerge
    Def(p, "LOC_FIELD_TEXT").Values("Address") = "B2:C2"
    Check ReadText(p).State = "CONFLICT", "unmerged multicell scalar is CONFLICT"
    Def(p, "LOC_FIELD_TEXT").Values("Sheet") = "Absent"
    Check ReadText(p).State = "MISSING", "missing worksheet is MISSING"
    Set resolver = Reader(Model())
    Check resolver.ReadField(firstBook, "ABSENT").State = "MISSING", "missing FieldId is MISSING"
    Check resolver.ReadField(Nothing, "FIELD_TEXT").State = "UNAVAILABLE", "missing workbook resource is UNAVAILABLE"
End Sub

Private Sub TestNames()
    Dim p As CAtPackage, s As Object, n As Object, r As CAtEvalResult, before As Long
    Set p = Model(): Set s = firstBook.Worksheets("Request")
    Strategy p, "DefinedName": Def(p, "LOC_FIELD_TEXT").Values.Add "DefinedName", "InputValue"
    Check ReadText(p).State = "MISSING", "missing DefinedName is MISSING"
    s.Range("B2").Value2 = "named"
    Set n = firstBook.Names.Add("InputValue", "=Request!$B$2")
    Check ReadText(p).Value = "named", "workbook-scoped DefinedName resolves"
    s.Names.Add "Request!InputValue", "=Request!$C$2"
    Check ReadText(p).State = "CONFLICT", "local/global DefinedName ambiguity is CONFLICT"
    ' Excel short-name lookup can bind the sheet name even through Workbook.Names.
    Do While firstBook.Names.Count > 0: firstBook.Names(1).Delete: Loop
    s.Names.Add "Request!InputValue", "=Request!$C$2"
    s.Range("C2").Value2 = "local"
    Check ReadText(p).Value = "local", "worksheet-scoped DefinedName resolves"
    firstBook.Names(1).Delete
    Set n = firstBook.Names.Add("InputValue", "=1+1")
    Check ReadText(p).State = "INVALID", "name expression is never evaluated"
    n.RefersTo = "=Request!$B$2,Request!$C$2"
    Check ReadText(p).State = "INVALID", "union DefinedName unsupported explicitly"
    n.RefersTo = "=#REF!": Check ReadText(p).State = "MISSING", "broken name reference is MISSING"
    n.RefersTo = "=Request!B2"
    Check ReadText(p).State = "INVALID", "relative DefinedName rejected instead of depending on active cell"
    s.Activate: s.Range("J10").Select
    Check ReadText(p).State = "INVALID", "relative DefinedName remains rejected after active-cell change"
    n.RefersTo = "=Request!$B2"
    Check ReadText(p).State = "INVALID", "mixed relative DefinedName rejected"
    n.RefersTo = "='C:\nonexistent\[do-not-open.xlsx]Data'!$A$1"
    before = Application.Workbooks.Count: Set r = ReadText(p)
    Check r.State = "UNAVAILABLE" And r.Diagnostics(1)("Code") = "R_EXTERNAL_NAME", "external name is rejected before dereferencing"
    Check Application.Workbooks.Count = before, "external name does not open a workbook"
    n.Delete
    Set n = firstBook.Names.Add("InputValue", "=Request!$B$2:$C$2")
    Check ReadText(p).State = "CONFLICT", "multicell DefinedName does not silently pick first value"
    n.Delete
End Sub

Private Sub TestLabels()
    Dim p As CAtPackage, s As Object, r As CAtEvalResult
    Set p = Model(): Set s = firstBook.Worksheets("Request")
    Strategy p, "LabelOffset"
    Def(p, "LOC_FIELD_TEXT").Values.Add "Label", "Exact label"
    Def(p, "LOC_FIELD_TEXT").Values.Add "RowOffset", "1"
    Def(p, "LOC_FIELD_TEXT").Values.Add "ColumnOffset", "2"
    Check ReadText(p).State = "MISSING", "absent label is MISSING"
    s.Range("E5:F5").Merge: s.Range("E5").Value2 = "Exact label": s.Range("G6").Value2 = "offset"
    Check ReadText(p).Value = "offset", "merged label uses upper-left anchor for offset"
    s.Range("H5").Value2 = "Exact label"
    Check ReadText(p).State = "CONFLICT", "duplicate exact labels are CONFLICT"
    s.Range("H5").ClearContents: s.Range("E5").Value2 = "exact label"
    Check ReadText(p).State = "MISSING", "label match is exact binary and case sensitive"
    s.Range("E5").Value2 = "Exact label"
    Def(p, "LOC_FIELD_TEXT").Values("RowOffset") = "-100"
    Check ReadText(p).State = "MISSING", "label offset outside worksheet is MISSING"
    Def(p, "LOC_FIELD_TEXT").Values("RowOffset") = "1"
    s.Range("Z5000").Value2 = "extent"
    Check ReadText(p).State = "LIMIT", "oversized UsedRange is bounded before scan"
    ' This is fixture maintenance, not Resolver behavior.
    s.Rows("100:5000").Delete
    s.Range("E5:F5").UnMerge: s.Range("E5").ClearContents
End Sub

Private Sub TestTables()
    Dim p As CAtPackage, resolver As CAtExcelResolver, table As CAtTableSnapshot, s As Object, b As Object
    Dim r As CAtEvalResult, ids As Collection, d As CAtDefinition, context As CAtEvalContext, issues As Collection, evaluator As New CAtEvaluator
    Set p = Model(True): Set s = secondBook.Worksheets("Return"): Set b = Bindings()
    Set resolver = Reader(p, True, b): Set table = resolver.ReadTable(secondBook, "ITEMS")
    Check table.Status.State = "OK" And table.RowCount = 4 And table.StartRow = 5, "empty bounded table retains all four physical rows"
    Check table.ColumnIds.Count = 3 And table.IsEmptyRow(1) And table.IsEmptyRow(4), "empty table exposes declared columns and empty rows"
    s.Range("A5").Value2 = "asset": s.Range("B5").Value2 = 2: s.Range("C5").Value2 = True
    s.Range("A7").Value2 = "partial"
    Set table = resolver.ReadTable(secondBook, "ITEMS")
    Check table.ReadCell(1, "ITEM_CODE").Value = "asset" And table.ReadCell(1, "ITEM_COUNT").Value = "2", "multiple typed table columns read generically"
    Check table.ReadCell(1, "ITEM_READY").Value = "true", "table Boolean codec matches scalar codec"
    Check table.IsEmptyRow(2) And Not table.IsEmptyRow(3) And table.IsEmptyRow(4), "blank interior row does not terminate table"
    Check table.ReadCell(3, "ITEM_COUNT").State = "EMPTY", "partially filled row retains typed EMPTY cells"
    s.Range("A5").Value2 = "changed"
    Check table.ReadCell(1, "ITEM_CODE").Value = "asset", "table snapshot is independent of later workbook edits"
    Set r = table.ReadCell(1, "ITEM_CODE"): r.SetData "Text", "OK", "mutated"
    Set ids = table.ColumnIds: ids.Remove 1
    Check table.ReadCell(1, "ITEM_CODE").Value = "asset" And table.ColumnIds.Count = 3, "table getters return defensive copies"
    s.Range("B7").Value2 = "bad"
    Set table = resolver.ReadTable(secondBook, "ITEMS")
    Check table.Status.State = "INVALID" And table.ReadCell(3, "ITEM_COUNT").State = "INVALID", "invalid table cell affects table status and remains inspectable"
    Check table.Status.Diagnostics(1)("ObjectId") = "ITEM_COUNT/7", "table diagnostic identifies logical column and physical row"
    s.Range("B7").ClearContents
    s.Range("A7:A8").Merge: Set table = resolver.ReadTable(secondBook, "ITEMS")
    Check table.Status.State = "CONFLICT" And table.ReadCell(3, "ITEM_CODE").State = "CONFLICT" And table.ReadCell(4, "ITEM_CODE").State = "CONFLICT", "vertical table merge conflicts without copying values to rows"
    s.Range("A7:A8").UnMerge: s.Range("A8:C8").Merge
    Set table = resolver.ReadTable(secondBook, "ITEMS")
    Check table.ReadCell(4, "ITEM_COUNT").State = "CONFLICT", "horizontal table merge is also explicit CONFLICT"
    s.Range("A8:C8").UnMerge
    Def(p, "LOC_ITEMS").Values("Address") = "A5:D8": s.Range("D5:D6").Merge
    Set resolver = Reader(p, True, b): Set table = resolver.ReadTable(secondBook, "ITEMS")
    Check table.Status.State = "CONFLICT", "merge in an unmapped column still makes table bounds conflicting"
    s.Range("D5:D6").UnMerge: Def(p, "LOC_ITEMS").Values("Address") = "A5:C8"
    Set resolver = Reader(p, True): Set table = resolver.ReadTable(secondBook, "ITEMS")
    Check table.Status.State = "MISSING" And table.RowCount = 0, "missing column binding returns no partial rows"
    b("ITEM_COUNT") = 1: Set resolver = Reader(p, True, b)
    Check resolver.ReadTable(secondBook, "ITEMS").Status.State = "CONFLICT", "duplicate physical column bindings conflict"
    b("ITEM_COUNT") = 4: Set resolver = Reader(p, True, b)
    Check resolver.ReadTable(secondBook, "ITEMS").Status.State = "MISSING", "column binding outside bounds is MISSING"
    Set b = Bindings(): Set resolver = Reader(p, True, b): b("ITEM_CODE") = 3
    Check resolver.ReadTable(secondBook, "ITEMS").ReadCell(1, "ITEM_CODE").Value = "changed", "Resolver freezes column binding configuration"
    Check resolver.ReadTable(secondBook, "ABSENT").Status.State = "MISSING", "missing TableId is MISSING"
    Def(p, "LOC_ITEMS").Values("Address") = "A5:C1005": Set resolver = Reader(p, True, Bindings())
    Check resolver.ReadTable(secondBook, "ITEMS").Status.State = "LIMIT", "table row cap checked before data reads"
    Def(p, "LOC_ITEMS").Values("Address") = "A5:C8"
    Strategy p, "TableColumn": Def(p, "LOC_FIELD_TEXT").Values("Sheet") = "Return"
    Set d = Def(p, "LOC_FIELD_TEXT")
    AtAddReference d, "Table", "Table", "ITEMS", "1.0.0"
    AtAddReference d, "Column", "Column", "ITEM_CODE", "1.0.0", "Text"
    Set resolver = Reader(p, True, Bindings())
    Check resolver.ReadField(secondBook, "FIELD_TEXT").State = "CONFLICT", "TableColumn with several rows cannot become scalar implicitly"
    Def(p, "LOC_ITEMS").Values("Address") = "A5:C5": Set resolver = Reader(p, True, Bindings())
    Check resolver.ReadField(secondBook, "FIELD_TEXT").Value = "changed", "TableColumn scalar works for exactly one declared row"
    s.Range("A5").ClearContents
    Set r = resolver.ReadField(secondBook, "FIELD_TEXT")
    Check r.State = "EMPTY" And r.Dependency = "FIELD_TEXT", "TableColumn retains logical FieldId as dependency"
    Def(p, "LOC_FIELD_TEXT").Values("Sheet") = "Other"
    Set resolver = Reader(p, True, Bindings())
    Check resolver.ReadField(secondBook, "FIELD_TEXT").State = "INVALID", "TableColumn cannot silently ignore inconsistent declared sheet"
    Def(p, "LOC_FIELD_TEXT").Values("Sheet") = "Return"
    Set d = Def(p, "EXPR_READY")
    d.Nodes(1).Values("FieldId") = "FIELD_TEXT": d.Nodes(1).ResultType = "Text"
    d.Values("ReturnType") = "Text"
    ' Use the existing Bool expression as StateIs over the empty TableColumn.
    d.Nodes(1).Id = "INPUT": d.RootNodeId = "ROOT": d.Values("ReturnType") = "Bool"
    Dim stateNode As CAtNode
    Set stateNode = AtNewNode("ROOT", "StateIs", "Bool")
    stateNode.Values.Add "States", "EMPTY": stateNode.Children.Add "INPUT": d.Nodes.Add stateNode
    Set resolver = Reader(p, True, Bindings()): Set context = resolver.Resolve(secondBook, issues)
    Set r = evaluator.Evaluate(p, "EXPR_READY", context)
    Check r.Value = "true" And Not r.IsComplete And r.Requirements(1)("Dependency") = "FIELD_TEXT", "EMPTY TableColumn StateIs keeps unresolved field requirement"
End Sub

Private Sub TestSnapshotBudgets()
    Dim p As CAtPackage, resolver As CAtExcelResolver, table As CAtTableSnapshot, context As CAtEvalContext, issues As Collection, i As Long
    Set p = Model(True)
    Def(p, "LOC_ITEMS").Values("Address") = "A20:C59"
    secondBook.Worksheets("Return").Range("A20:A59").Value2 = String$(32767, "x")
    Set resolver = Reader(p, True, Bindings()): Set table = resolver.ReadTable(secondBook, "ITEMS")
    Check table.Status.State = "LIMIT" And table.RowCount = 0, "table aggregate text/metadata cap returns no partial rows"
    secondBook.Worksheets("Return").Range("A20:A59").ClearContents
    Set p = Model()
    firstBook.Worksheets("Request").Range("M1").Value2 = String$(32767, "x")
    For i = 1 To 34: AddField p, "SIMPLE", "Request", "FIELD_LARGE_" & i, "Text", "M1": Next i
    Set resolver = Reader(p): Set context = resolver.Resolve(firstBook, issues)
    Check context Is Nothing And issues.Count = 1 And issues(1)("Code") = "R_CONTEXT_LIMIT", "context aggregate cap returns Nothing plus LIMIT diagnostic rather than partial success"
    firstBook.Worksheets("Request").Range("M1").ClearContents
End Sub

Private Sub TestIntegration(ByVal scratch As String)
    Dim equipment As Variant, p As CAtPackage, resolver As CAtExcelResolver, evaluator As New CAtEvaluator
    Dim context As CAtEvalContext, issues As Collection, workbook As Object, sheet As Object, r As CAtEvalResult
    For Each equipment In Array(False, True)
        Set p = Model(CBool(equipment)): Set resolver = Reader(p, CBool(equipment), IIfBindings(CBool(equipment)))
        If equipment Then Set workbook = secondBook Else Set workbook = firstBook
        Set sheet = workbook.Worksheets(1): sheet.Range("B3").Value2 = True
        Set context = resolver.Resolve(workbook, issues)
        Check Not context Is Nothing And issues.Count = 0, "Excel to immutable context " & p.Id
        Set r = evaluator.Evaluate(p, "EXPR_READY", context)
        Check r.IsComplete And r.Value = "true", "real Excel to Evaluator true " & p.Id
        sheet.Range("B3").Value2 = False
        Check evaluator.Evaluate(p, "EXPR_READY", context).Value = "true", "previous context is immutable " & p.Id
        Set context = resolver.Resolve(workbook, issues)
        Check evaluator.Evaluate(p, "EXPR_READY", context).Value = "false", "fresh snapshot changes evaluation " & p.Id
        Def(p, "LOC_FIELD_FLAG").Values("Address") = "D3": sheet.Range("D3").Value2 = True
        Check evaluator.Evaluate(p, "EXPR_READY", resolver.Resolve(workbook, issues)).Value = "false", "Resolver freezes locator definitions " & p.Id
        Set resolver = Reader(p, CBool(equipment), IIfBindings(CBool(equipment)))
        Set context = resolver.Resolve(workbook, issues)
        Check evaluator.Evaluate(p, "EXPR_READY", context).Value = "true" And Def(p, "RULE_READY").Id = "RULE_READY", "moved coordinate keeps FieldId and RuleId " & p.Id
        sheet.Range("D3").Formula = "=1/0": sheet.Range("D3").Calculate
        Set r = evaluator.Evaluate(p, "EXPR_READY", resolver.Resolve(workbook, issues))
        Check r.State = "INVALID" And Not r.IsComplete And r.Requirements.Count > 0, "Excel error cannot become successful evaluation " & p.Id
        sheet.Range("D3").Value2 = True
        If equipment Then workbook.SaveAs scratch & "\equipment-return-resolver.xlsx", 51 Else workbook.SaveAs scratch & "\simple-resolver.xlsx", 51
    Next equipment
End Sub

Private Function IIfBindings(ByVal equipment As Boolean) As Object
    If equipment Then Set IIfBindings = Bindings() Else Set IIfBindings = AtDictionary()
End Function

Private Sub TestSafety(ByVal scratch As String)
    Dim p As CAtPackage, resolver As CAtExcelResolver, issues As Collection, context As CAtEvalContext, r As CAtEvalResult
    Dim s As Object, cell As Object, active As Object, count As Long, names As Long, sheets As Long, used As String
    Dim events As Boolean, screen As Boolean, alerts As Boolean, links As Boolean, calculation As Long, security As Long
    Set p = Model(): Set resolver = Reader(p): Set s = firstBook.Worksheets("Request"): Set cell = s.Range("B2")
    cell.Formula = "=""user formula""": cell.NumberFormat = "@": cell.Font.Bold = True
    cell.Interior.Color = RGB(13, 24, 35): cell.AddComment "User comment: " & ChrW(&H416)
    s.Range("J10:K11").Merge: s.Range("J10").Value2 = "user merge"
    firstBook.Save
    Set active = Application.ActiveWorkbook: count = Application.Workbooks.Count: names = firstBook.Names.Count
    sheets = firstBook.Worksheets.Count: used = s.UsedRange.Address
    events = Application.EnableEvents: screen = Application.ScreenUpdating: calculation = Application.Calculation
    alerts = Application.DisplayAlerts: links = Application.AskToUpdateLinks: security = Application.AutomationSecurity
    Set context = resolver.Resolve(firstBook, issues): Set r = resolver.ReadField(firstBook, "FIELD_TEXT")
    Check r.Value = "user formula" And firstBook.Saved, "read-only run preserves a saved workbook's Saved=True"
    Set context = resolver.Resolve(firstBook, issues)
    Check cell.Formula = "=""user formula""" And cell.NumberFormat = "@" And cell.Font.Bold And cell.Interior.Color = RGB(13, 24, 35), "repeated reads preserve formula and user formatting"
    Check cell.Comment.Text = "User comment: " & ChrW(&H416) And s.Range("J10").MergeArea.Address = "$J$10:$K$11", "repeated reads preserve user comments and merges"
    Check firstBook.Worksheets.Count = sheets And firstBook.Names.Count = names And s.UsedRange.Address = used, "repeated reads preserve workbook structure"
    Check Application.Workbooks.Count = count And Application.ActiveWorkbook Is active, "explicit non-active workbook read neither opens nor activates books"
    Check Application.EnableEvents = events And Application.ScreenUpdating = screen And Application.Calculation = calculation And Application.DisplayAlerts = alerts And Application.AskToUpdateLinks = links And Application.AutomationSecurity = security, "Resolver preserves all captured application flags"
    Application.Calculation = -4135: Application.EnableEvents = Not events: Application.ScreenUpdating = Not screen
    s.Range("A20").Value2 = "before": cell.NumberFormat = "General": cell.Formula = "=A20": cell.Calculate
    s.Range("A20").Value2 = "after"
    Check cell.HasFormula And cell.Value2 = "before", "manual fixture starts with a real stale formula"
    Set r = resolver.ReadField(firstBook, "FIELD_TEXT")
    Check r.Value = "before" And cell.Value2 = "before", "manual calculation reads cached result without recalculating dirty formula"
    Check Application.Calculation = -4135 And Application.EnableEvents = Not events And Application.ScreenUpdating = Not screen, "nondefault application flags also survive resolution"
    Application.Calculation = calculation: Application.EnableEvents = events: Application.ScreenUpdating = screen
    cell.Value2 = "unsaved edit"
    Check Not firstBook.Saved, "fixture starts dirty for Saved=False check"
    Set context = resolver.Resolve(firstBook, issues)
    Check Not firstBook.Saved And cell.Value2 = "unsaved edit", "read-only run preserves dirty workbook and unsaved value"
    Def(p, "LOC_FIELD_TEXT").Values("Sheet") = "Absent": Set resolver = Reader(p)
    Set r = resolver.ReadField(firstBook, "FIELD_TEXT")
    Check r.State = "MISSING" And Not firstBook.Saved And Application.Workbooks.Count = count, "failed resolution also leaves workbook untouched"
    Set p = Model(): Strategy p, "LabelOffset"
    Def(p, "LOC_FIELD_TEXT").Values.Add "Label", "Metadata label"
    Def(p, "LOC_FIELD_TEXT").Values.Add "RowOffset", "0"
    Def(p, "LOC_FIELD_TEXT").Values.Add "ColumnOffset", "2"
    s.Range("F15:G15").Merge: s.Range("F15").Value2 = "Metadata label": s.Range("H15").Value2 = "metadata value"
    firstBook.Save: Set resolver = Reader(p): Set r = resolver.ReadField(firstBook, "FIELD_TEXT")
    Check r.Value = "metadata value" And firstBook.Saved, "LabelOffset UsedRange and merged-anchor reads preserve Saved=True"
    Set p = Model(True): Set resolver = Reader(p, True, Bindings())
    Dim table As CAtTableSnapshot: Set table = resolver.ReadTable(secondBook, "ITEMS")
    Check table.Status.State = "OK" And secondBook.Saved, "table snapshot read preserves Saved=True"
End Sub
