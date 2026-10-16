Attribute VB_Name = "modATDocumentFieldTests"
Option Explicit

Public Sub AT_TestDocumentFields()
    Dim book As Workbook, ws As Worksheet, spec As ATExcelFieldSpec, reader As New ATExcelFieldReader
    Dim field As ATExcelFieldValue, context As ATEvaluationContext, copy As ATEvaluationContext
    Dim schema As New ATCompileContext, compiler As New ATCompiler, evaluator As New ATEvaluator
    Dim result As ATEvalResult, savedNumber As Long, savedText As String
    Dim trace As ATTraceEntry, located As Boolean, rows As New Collection, row As New ATRuntimeRow
    Dim value As New ATRuntimeValue, limits As New ATEvalLimits
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = book.Worksheets(1): ws.Name = "Request"
    ws.Range("A1:B1").Merge: ws.Range("A1").Value2 = "Request type"
    ws.Range("C1:D1").Merge: ws.Range("C1").Value2 = "external"
    Set spec = AT_TestExcelFieldSpec("TYPE", "Request type", 1, 1, 1, 4, 0, 1)
    Set field = reader.Read(ws, spec)
    Set context = AT_TestDocumentContext()
    context.AddField "TYPE", field.Result, field.CellRef
    Set copy = context.Snapshot
    AT_Assert copy.CellRef("TYPE") = "'Request'!C1", "document snapshot retains cell metadata"
    AT_Assert copy.CellRef("ABSENT") = "", "missing document field has no fabricated cell"
    schema.AddField "TYPE", "TEXT"
    book.Close False: Set book = Nothing
    Set result = evaluator.Evaluate(compiler.Compile("@TYPE = ""external""", schema), copy)
    AT_Assert result.State = atEvalOK, "document extracted field evaluates after closing source"
    If result.State = atEvalOK Then AT_Assert result.Value.Scalar, "document extracted comparison correct"
    AT_Assert AT_TestTraceCell(result, "TYPE", "'Request'!C1", True), "document field success trace has source cell"
    row.Initialize "T", "12"
    row.AddField "N", AT_RuntimeOK(AT_RuntimeScalar("NUMBER", 1#)), "'Request'!A12"
    rows.Add row: value.InitializeRows "T", rows
    context.AddField "ROWS", AT_RuntimeOK(value)
    schema.AddField "ROWS", "COLLECTION<ROW<T>>": schema.AddField "N", "NUMBER", "T"
    Set result = evaluator.Evaluate(compiler.Compile("ALLROWS(@ROWS AS t, @t.N > 0 AND @TYPE = ""external"")", schema), context)
    AT_Assert result.State = atEvalOK, "document field inside row predicate evaluates"
    AT_Assert AT_TestTraceCell(result, "TYPE", "'Request'!C1", True), "document reference never uses active row cell"
    AT_Assert AT_TestTraceCell(result, "N", "'Request'!A12", False), "row field keeps own source reference"
    Set context = AT_TestDocumentContext()
    context.AddField "TYPE", AT_RuntimeFailure(atEvalError, "EXCEL_CELL_ERROR", Nothing), "'Request'!C1"
    Set result = evaluator.Evaluate(compiler.Compile("@TYPE = ""external""", schema), context)
    AT_Assert result.State = atEvalError, "document input error propagates"
    AT_Assert AT_TestTraceCell(result, "TYPE", "'Request'!C1", True), "document error trace retains actual field cell"
    Set result = evaluator.Evaluate(compiler.Compile("TRUE OR @TYPE = ""external""", schema), context)
    AT_Assert result.State = atEvalOK, "skipped document error stays lazy"
    Set context = AT_TestDocumentContext()
    context.AddField "TYPE", AT_RuntimeOK(AT_RuntimeState("TEXT", atValueEmpty)), "'Request'!C1"
    Set result = evaluator.Evaluate(compiler.Compile("@TYPE = ""external""", schema), context)
    AT_Assert AT_TestTraceCell(result, "TYPE", "'Request'!C1", True), "empty document field still traceable"
    Set context = AT_TestDocumentContext()
    context.AddField "TYPE", AT_RuntimeOK(AT_RuntimeScalar("TEXT", "external")), String$(21, "x")
    limits.MaxTextLength = 20
    Set result = evaluator.Evaluate(compiler.Compile("@TYPE = ""external""", schema), context, limits)
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "RUNTIME_LIMIT_TEXT", "document cell metadata obeys trace text budget"
    Exit Sub
CleanupError:
    savedNumber = Err.Number: savedText = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise savedNumber, , savedText
End Sub

Private Function AT_TestDocumentContext() As ATEvaluationContext
    Dim context As New ATEvaluationContext
    context.Initialize "DOCUMENT_TEST", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    Set AT_TestDocumentContext = context
End Function

Private Function AT_TestTraceCell(ByVal result As ATEvalResult, ByVal fieldId As String, ByVal cellRef As String, ByVal document As Boolean) As Boolean
    Dim entry As ATTraceEntry
    For Each entry In result.Trace
        If entry.FieldId = fieldId And entry.CellRef = cellRef Then
            If document Then
                AT_TestTraceCell = Len(entry.RowId) = 0 And Len(entry.TableId) = 0
            Else
                AT_TestTraceCell = entry.RowId = "12" And entry.TableId = "T"
            End If
            If AT_TestTraceCell Then Exit Function
        End If
    Next entry
End Function
