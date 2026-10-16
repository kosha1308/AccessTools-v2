Attribute VB_Name = "modATRowEvalTests"
Option Explicit

Public Function AT_TestRowsValue(ByVal numbers As Variant, Optional ByVal failure As ATEvalResult = Nothing) As ATRuntimeValue
    Dim rows As New Collection, row As ATRuntimeRow, number As Variant, i As Long, value As New ATRuntimeValue
    For Each number In numbers
        i = i + 1
        Set row = New ATRuntimeRow
        row.Initialize "T", "R" & CStr(i)
        row.AddField "N", AT_RuntimeOK(AT_RuntimeScalar("NUMBER", CDbl(number))), "Sheet1!C" & CStr(i + 10)
        rows.Add row
    Next number
    If Not failure Is Nothing Then
        i = i + 1
        Set row = New ATRuntimeRow
        row.Initialize "T", "R" & CStr(i)
        row.AddField "N", failure, "Sheet1!C" & CStr(i + 10)
        rows.Add row
    End If
    value.InitializeRows "T", rows
    Set AT_TestRowsValue = value
End Function

Public Function AT_RowFixture(ByVal expression As String, ByVal rows As ATRuntimeValue, Optional ByVal requiredType As String = "BOOL", Optional ByVal limits As ATEvalLimits = Nothing) As ATEvalResult
    Dim schema As New ATCompileContext, compiler As New ATCompiler, context As New ATEvaluationContext, evaluator As New ATEvaluator
    schema.AddField "ROWS", "COLLECTION<ROW<T>>": schema.AddField "N", "NUMBER", "T": schema.AddField "DOC", "NUMBER"
    context.Initialize "ROW_TEST", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    context.AddField "ROWS", AT_RuntimeOK(rows): context.AddField "DOC", AT_RuntimeOK(AT_RuntimeScalar("NUMBER", 0#))
    Set AT_RowFixture = evaluator.Evaluate(compiler.Compile(expression, schema, , requiredType), context, limits)
End Function

Public Sub AT_TestRowEval()
    Dim expression As Variant, state As Variant, op As Variant, result As ATEvalResult, rows As ATRuntimeValue
    Dim trace As ATTraceEntry, skipped As Boolean, located As Boolean, limits As New ATEvalLimits
    Set rows = AT_TestRowsValue(Array(1#, 2#))
    For Each expression In Array("ALLROWS(@ROWS AS t, @t.N > @DOC)", "ALLROWS(@ROWS AS t, @N > 0)", "ANYROW(@ROWS AS t, @t.N = 2)", "NOROWS(@ROWS AS t, @t.N < 0)", "COUNT(@ROWS) = 2", "EXISTS(@ROWS)", "COUNT(FILTER(@ROWS AS t, @t.N > 1)) = 1", "ALLROWS(@ROWS AS t, ANYROW(@ROWS AS u, @u.N = @t.N))")
        Set result = AT_RowFixture(CStr(expression), rows)
        AT_Assert result.State = atEvalOK, "row evaluation state: " & CStr(expression) & AT_RuntimeTestDetails(result)
        If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "row evaluation value"
    Next expression
    Set result = AT_RowFixture("FILTER(@ROWS AS t, @t.N > 1)", rows, "COLLECTION<ROW<T>>")
    AT_Assert result.State = atEvalOK, "FILTER yields typed collection"
    If result.State = atEvalOK Then AT_Assert result.Value.Row(1).RowId = "R2" And result.Value.Count = 1, "FILTER preserves row identity/order"
    Set rows = AT_TestRowsValue(Array())
    For Each expression In Array("ALLROWS(@ROWS AS t, FALSE)", "NOROWS(@ROWS AS t, TRUE)", "NOT ANYROW(@ROWS AS t, TRUE)", "NOT EXISTS(@ROWS)", "COUNT(FILTER(@ROWS AS t, TRUE)) = 0")
        Set result = AT_RowFixture(CStr(expression), rows)
        AT_Assert result.State = atEvalOK, "empty row collection state"
        If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "empty row collection laws"
    Next expression
    For Each state In Array(atEvalError, atEvalSkipped, atEvalManualReview)
        Set rows = AT_TestRowsValue(Array(), AT_RuntimeFailure(CLng(state), "ROW_INPUT_FAILURE", Nothing))
        For Each op In Array("ALLROWS", "ANYROW", "NOROWS", "FILTER")
            Set result = AT_RowFixture(CStr(op) & "(@ROWS AS t, @t.N > 0)", rows, IIf(op = "FILTER", "COLLECTION<ROW<T>>", "BOOL"))
            AT_Assert result.State = state, "row predicate preserves non-OK state"
            AT_Assert result.Diagnostics(1).Code = "ROW_INPUT_FAILURE", "row predicate preserves primary diagnostic"
        Next op
        Set result = AT_RowFixture("COUNT(@ROWS) = 1 AND EXISTS(@ROWS)", rows)
        AT_Assert result.State = atEvalOK, "COUNT/EXISTS do not evaluate row fields"
    Next state
    Set rows = AT_TestRowsValue(Array(1#), AT_RuntimeFailure(atEvalError, "DO_NOT_READ", Nothing))
    Set result = AT_RowFixture("ANYROW(@ROWS AS t, @t.N = 1)", rows)
    AT_Assert result.State = atEvalOK, "row short circuit skips bad following field"
    If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "row short circuit result"
    For Each trace In result.Trace
        If trace.Status = "NOT_EVALUATED" And trace.RowId = "R2" Then skipped = True
        If trace.FieldId = "N" And trace.RowId = "R1" And trace.CellRef = "Sheet1!C11" Then located = True
    Next trace
    AT_Assert skipped, "skipped row identity explained"
    AT_Assert located, "evaluated row field cell metadata explained"
    limits.MaxListItems = 1
    Set result = AT_RowFixture("COUNT(@ROWS) = 2", rows, , limits)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_LIMIT_LIST", "row count bounded before returned clone"
    limits.MaxListItems = 10000: limits.MaxTraceEntries = 1
    Set rows = AT_TestRowsValue(Array(), AT_RuntimeFailure(atEvalManualReview, "NEEDS_ROW_REVIEW", Nothing))
    Set result = AT_RowFixture("ALLROWS(@ROWS AS t, @t.N > 0)", rows, , limits)
    AT_Assert result.State = atEvalManualReview, "row trace overflow retains primary state"
    AT_Assert result.Diagnostics(1).Code = "NEEDS_ROW_REVIEW", "row trace overflow retains primary code"
    limits.MaxTraceEntries = 20000: limits.MaxDepth = 2
    Set rows = AT_TestRowsValue(Array(1#, 2#))
    Set result = AT_RowFixture("ALLROWS(@ROWS AS t, @t.N > 0)", rows, , limits)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_LIMIT_DEPTH", "row predicate depth bounded"
    limits.MaxDepth = 3
    Set result = AT_RowFixture("ALLROWS(@ROWS AS t, @t.N > 0)", rows, , limits)
    AT_Assert result.State = atEvalOK, "row predicate exact depth boundary"
    limits.MaxSteps = 1
    Set result = AT_RowFixture("ALLROWS(@ROWS AS t, TRUE)", rows, , limits)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_LIMIT_STEPS", "row work bounded"
    Set rows = AT_TestRowsValue(Array(), AT_RuntimeOK(AT_RuntimeScalar("TEXT", "wrong type")))
    Set result = AT_RowFixture("ALLROWS(@ROWS AS t, @t.N > 0)", rows)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_TYPE_MISMATCH", "row field type mismatch"
    Set rows = AT_TestRowsValue(Array(), AT_RuntimeOK(AT_RuntimeState("NUMBER", atValueInvalid)))
    Set result = AT_RowFixture("ALLROWS(@ROWS AS t, @t.N > 0)", rows)
    AT_Assert result.State = atEvalError, "invalid row field cannot pass predicate"
    Set result = AT_RowFixture("@ROWS = @ROWS", AT_TestRowsValue(Array(1#)))
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_UNSUPPORTED_PRIMITIVE", "collection equality remains explicit unsupported"
    AT_TestNamedRows
    AT_TestCorrelatedIPRows
    AT_TestUnusedRowBudgets
End Sub

Private Sub AT_TestNamedRows()
    Dim schema As New ATCompileContext, context As New ATEvaluationContext, compiler As New ATCompiler, evaluator As New ATEvaluator
    Dim named As ATNamedExpression, result As ATEvalResult, expression As Variant
    schema.AddField "ROWS", "COLLECTION<ROW<T>>": schema.AddField "N", "NUMBER", "T"
    context.Initialize "ROW_NAMES", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    context.AddField "ROWS", AT_RuntimeOK(AT_TestRowsValue(Array(1#, 2#)))
    Set named = New ATNamedExpression
    named.Initialize "INNER_DOC", "BOOL", "ANYROW(@ROWS AS t, @t.N = 2)"
    schema.AddNamed named
    Set named = New ATNamedExpression
    named.Initialize "CAPTURE", "BOOL", "$INNER_DOC AND @N = $X"
    named.Scope = "ROW": named.AddParameter "X", "NUMBER", "@N"
    schema.AddNamed named
    Set named = New ATNamedExpression
    named.Initialize "DOC_DEFAULT", "BOOL", "$X"
    named.AddParameter "X", "BOOL", "ANYROW(@ROWS AS t, @t.N = 2)"
    schema.AddNamed named
    Set named = New ATNamedExpression
    named.Initialize "CAPTURE_DEFAULT", "BOOL", "$DOC_DEFAULT AND @N = $X"
    named.Scope = "ROW": named.AddParameter "X", "NUMBER", "@N"
    schema.AddNamed named
    For Each expression In Array("ALLROWS(@ROWS AS t, $INNER_DOC)", "ALLROWS(@ROWS AS t, $CAPTURE)", "ALLROWS(@ROWS AS t, $CAPTURE(@t.N))", "ALLROWS(@ROWS AS t, $DOC_DEFAULT)", "ALLROWS(@ROWS AS t, $CAPTURE_DEFAULT)")
        Set result = evaluator.Evaluate(compiler.Compile(CStr(expression), schema), context)
        AT_Assert result.State = atEvalOK, "named row/document scope state" & AT_RuntimeTestDetails(result)
        If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "named row/document scopes restore aliases"
    Next expression
    Set result = evaluator.Evaluate(compiler.Compile("ALLROWS(@ROWS AS t, ANYROW(@ROWS AS u, 1 / (@u.N - 1) > 0))", schema), context)
    AT_Assert result.State = atEvalError, "nested predicate error fixture"
    Set result = evaluator.Evaluate(compiler.Compile("ALLROWS(@ROWS AS t, $CAPTURE_DEFAULT)", schema), context)
    AT_Assert result.State = atEvalOK, "row evaluator reused after nested error"
    If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "row scopes fresh after nested error"
    Set result = evaluator.Evaluate(compiler.Compile("ANYROW(@ROWS AS t, @t.N = 1)", schema), context)
    Set result = evaluator.Evaluate(compiler.Compile("ALLROWS(@ROWS AS t, $CAPTURE_DEFAULT)", schema), context)
    AT_Assert result.State = atEvalOK, "row evaluator reused after decisive skip"
    If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "row scopes fresh after decisive skip"
End Sub

Private Sub AT_TestUnusedRowBudgets()
    Dim schema As New ATCompileContext, context As New ATEvaluationContext, compiler As New ATCompiler, evaluator As New ATEvaluator
    Dim row As ATRuntimeRow, rows As New Collection, values As New Collection, list As New ATRuntimeValue, collection As New ATRuntimeValue
    Dim limits As New ATEvalLimits, result As ATEvalResult, i As Long, expression As Variant
    schema.AddField "ROWS", "COLLECTION<ROW<T>>": schema.AddField "N", "NUMBER", "T"
    schema.AddField "BIG", "TEXT", "T": schema.AddField "BIGLIST", "LIST<NUMBER>", "T"
    For i = 1 To 20
        values.Add AT_RuntimeScalar("NUMBER", CDbl(i))
    Next i
    list.InitializeList "NUMBER", values
    For i = 1 To 2
        Set row = New ATRuntimeRow: row.Initialize "T", "R" & CStr(i)
        row.AddField "N", AT_RuntimeOK(AT_RuntimeScalar("NUMBER", CDbl(i)))
        row.AddField "BIG", AT_RuntimeOK(AT_RuntimeScalar("TEXT", String$(256, "x")))
        row.AddField "BIGLIST", AT_RuntimeOK(list)
        rows.Add row
    Next i
    collection.InitializeRows "T", rows
    context.Initialize "ROW_LIMITS", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    context.AddField "ROWS", AT_RuntimeOK(collection)
    limits.MaxTextLength = 64: limits.MaxListItems = 2
    For Each expression In Array("COUNT(@ROWS) = 2", "EXISTS(@ROWS)", "ANYROW(@ROWS AS t, @t.N = 1)")
        Set result = evaluator.Evaluate(compiler.Compile(CStr(expression), schema), context, limits)
        AT_Assert result.State = atEvalOK, "unused oversize row fields are not validated"
        If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "unused oversize row fields do not affect result"
    Next expression
    Set result = evaluator.Evaluate(compiler.Compile("ALLROWS(@ROWS AS t, LEN(@t.BIG) > 0)", schema), context, limits)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_LIMIT_TEXT", "used oversize row text rejected before clone"
    Set result = evaluator.Evaluate(compiler.Compile("ALLROWS(@ROWS AS t, SUM(@t.BIGLIST) > 0)", schema), context, limits)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_LIMIT_LIST", "used oversize row list rejected before clone"
    Set rows = New Collection: Set row = New ATRuntimeRow
    row.Initialize "T", "MISSING": rows.Add row
    Set collection = New ATRuntimeValue: collection.InitializeRows "T", rows
    expression = "ALLROWS(@ROWS AS t, @t.N > 0)"
    Set result = AT_RowFixture(CStr(expression), collection)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_ROW_FIELD_MISSING", "missing row field stable diagnostic"
    AT_Assert result.Diagnostics(1).SourceId = "RULE" And result.Diagnostics(1).Start = InStr(1, CStr(expression), "@t.N"), "missing row field original rule position"
End Sub

Private Sub AT_TestCorrelatedIPRows()
    Dim schema As New ATCompileContext, context As New ATEvaluationContext, compiler As New ATCompiler, evaluator As New ATEvaluator
    Dim rows As Collection, row As ATRuntimeRow, value As ATRuntimeValue, ip As Variant, i As Long, result As ATEvalResult
    schema.AddField "TECH", "COLLECTION<ROW<TECHNICAL>>": schema.AddField "LOWER", "COLLECTION<ROW<ADDRESS>>"
    schema.AddField "DEST", "IP_EXPRESSION", "TECHNICAL": schema.AddField "IP", "IP_EXPRESSION", "ADDRESS"
    context.Initialize "IP_ROWS", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    Set rows = New Collection
    For Each ip In Array("192.0.2.0/24", "198.51.100.0/24")
        i = i + 1: Set row = New ATRuntimeRow
        row.Initialize "TECHNICAL", "T" & CStr(i)
        row.AddField "DEST", AT_ParseDomain("IP_EXPRESSION", CStr(ip), Nothing), "Request!B" & CStr(i + 10)
        rows.Add row
    Next ip
    Set value = New ATRuntimeValue: value.InitializeRows "TECHNICAL", rows
    context.AddField "TECH", AT_RuntimeOK(value)
    Set rows = New Collection: i = 0
    For Each ip In Array("192.0.2.10", "198.51.100.20", "203.0.113.1")
        i = i + 1: Set row = New ATRuntimeRow
        row.Initialize "ADDRESS", "A" & CStr(i)
        row.AddField "IP", AT_ParseDomain("IP_EXPRESSION", CStr(ip), Nothing), "Request!C" & CStr(i + 20)
        rows.Add row
    Next ip
    Set value = New ATRuntimeValue: value.InitializeRows "ADDRESS", rows
    context.AddField "LOWER", AT_RuntimeOK(value)
    Set result = evaluator.Evaluate(compiler.Compile("ALLROWS(@TECH AS t, ANYROW(@LOWER AS a, @a.IP INTERSECTS @t.DEST))", schema), context)
    AT_Assert result.State = atEvalOK, "correlated IP tables result state"
    If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "each upper network has lower matching IP"
    Set result = evaluator.Evaluate(compiler.Compile("ALLROWS(@LOWER AS a, ANYROW(@TECH AS t, @a.IP COVER @t.DEST))", schema), context)
    AT_Assert result.State = atEvalOK, "uncovered lower IP result state"
    If result.State = atEvalOK Then AT_Assert result.Value.Scalar = False, "uncovered lower IP not hidden by valid rows"
    Set result = evaluator.Evaluate(compiler.Compile("COUNT(FILTER(@LOWER AS a, ANYROW(@TECH AS t, @a.IP IN @t.DEST))) = 2", schema), context)
    AT_Assert result.State = atEvalOK, "correlated FILTER result state"
    If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "correlated FILTER counts matched rows"
End Sub
