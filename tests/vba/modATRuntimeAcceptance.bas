Attribute VB_Name = "modATRuntimeAcceptance"
Option Explicit

Public Sub AT_TestRuntimeAcceptance()
    Dim schema As New ATCompileContext, context As New ATEvaluationContext, tomorrow As New ATEvaluationContext
    Dim evaluator As New ATEvaluator, compiler As New ATCompiler, serializer As New ATSerializer
    Dim compiled As ATCompileResult, result As ATEvalResult, repeated As ATEvalResult
    Dim emptyItems As New Collection, items As New Collection, list As New ATRuntimeValue
    Dim limits As New ATEvalLimits, item As Variant, before As String, trace As Collection, entry As ATTraceEntry
    context.Initialize "ACCEPT", "2026-10-16", "UTC", "INVARIANT", "ASCII_TRIM_CI"
    tomorrow.Initialize "NEXT", "2026-10-17", "UTC", "INVARIANT", "ASCII_TRIM_CI"
    schema.AddField "NAME", "TEXT"
    schema.AddField "VALUES", "LIST<NUMBER>"
    schema.AddField "IP", "IP_EXPRESSION"
    schema.AddField "ROWS", "COLLECTION<ROW<T>>"
    schema.AddSource "S", "NUMBER", "NUMBER"
    schema.AddSourceColumn "S", "N", "NUMBER"
    context.AddField "NAME", AT_RuntimeOK(AT_RuntimeScalar("TEXT", " secret "))
    list.InitializeList "NUMBER", emptyItems
    context.AddField "VALUES", AT_RuntimeOK(list)
    For Each item In Array("SUM(@VALUES) = 0", "@VALUES EXACT @VALUES", "@VALUES COVER [1]", "[1] CONTAINS @VALUES", "@VALUES NONE [1]", "@NAME = ""SECRET""", """SECRET"" = @NAME")
        Set compiled = compiler.Compile(CStr(item), schema)
        Set result = evaluator.Evaluate(compiled, context)
        AT_Assert result.State = atEvalOK, "acceptance result state"
        If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "acceptance result value"
    Next item
    For Each item In Array("MIN(@VALUES) = 0", "MAX(@VALUES) = 0", "AVG(@VALUES) = 0")
        Set result = evaluator.Evaluate(compiler.Compile(CStr(item), schema), context)
        AT_Assert result.Diagnostics(1).Code = "AGGREGATE_EMPTY", "empty aggregate explicit error"
    Next item
    For Each item In Array("@IP COVER @IP", "@NAME MATCHES "".*""", "@NAME LIKE ""*""", "SOURCE(""S"") = 1", "LOOKUP(""S"",1,""N"") = 1", "COUNT(@ROWS) = 0", "EXISTS(@ROWS)", "ALLROWS(@ROWS AS t, TRUE)", "ANYROW(@ROWS AS t, TRUE)", "NOROWS(@ROWS AS t, FALSE)", "COUNT(FILTER(@ROWS AS t, TRUE)) = 0", "DISTINCT_COUNT(@VALUES) = 0")
        Set compiled = compiler.Compile(CStr(item), schema)
        AT_Assert Not compiled.Root Is Nothing, "deferred fixture compiles"
        Set result = evaluator.Evaluate(compiled, context)
        AT_Assert result.Diagnostics(1).Code = "RUNTIME_UNSUPPORTED_PRIMITIVE", "deferred signature explicit unsupported before missing input"
    Next item
    Set compiled = compiler.Compile("@NAME = ""SECRET""", schema)
    before = serializer.Serialize(compiled.Root)
    Set result = evaluator.Evaluate(compiled, context)
    Set repeated = evaluator.Evaluate(compiled, context)
    AT_Assert serializer.Serialize(compiled.Root) = before, "evaluation never mutates AST"
    AT_Assert result.Trace.Count = repeated.Trace.Count, "trace deterministic size"
    Set trace = result.Trace
    For Each entry In trace
        AT_Assert InStr(entry.Reason, "secret") = 0, "trace omits source payload"
        AT_Assert Len(entry.SourceId) > 0 And entry.Start > 0, "trace original provenance"
    Next entry
    Set compiled = compiler.Compile("TODAY() = DATE(""2026-10-16"")", schema)
    Set result = evaluator.Evaluate(compiled, context)
    Set repeated = evaluator.Evaluate(compiled, tomorrow)
    AT_Assert result.Value.Scalar = True And repeated.Value.Scalar = False, "context date changes without cache leak"
    Set result = evaluator.Evaluate(compiler.Compile("1 / 0 = 1", schema), context)
    Set result = evaluator.Evaluate(compiler.Compile("TRUE", schema), context)
    AT_Assert result.Value.Scalar = True, "evaluator resets after error"
    limits.MaxSteps = 1
    Set result = evaluator.Evaluate(compiler.Compile("NOT TRUE", schema), context, limits)
    Set result = evaluator.Evaluate(compiler.Compile("TRUE", schema), context)
    AT_Assert result.Value.Scalar = True, "evaluator resets after budget exhaustion"
    Set result = evaluator.Evaluate(compiler.Compile("@MISSING = 1", schema), context)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_COMPILE_REQUIRED", "invalid compiled input rejected"
    Set result = evaluator.Evaluate(Nothing, context)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_COMPILE_REQUIRED", "missing compiled input rejected"
    Set result = evaluator.Evaluate(compiler.Compile("TRUE", schema), Nothing)
    AT_Assert result.Diagnostics(1).Code = "CONFIG_CONTEXT_REQUIRED", "missing evaluation context rejected"
    AT_TestRuntimeInputLimits
End Sub

Private Sub AT_TestRuntimeInputLimits()
    Dim schema As New ATCompileContext, context As New ATEvaluationContext, compiler As New ATCompiler, evaluator As New ATEvaluator
    Dim list As New ATRuntimeValue, items As New Collection, limits As New ATEvalLimits, result As ATEvalResult
    schema.AddField "VALUES", "LIST<NUMBER>"
    schema.AddField "UNUSED", "LIST<NUMBER>"
    items.Add AT_RuntimeScalar("NUMBER", 1#)
    items.Add AT_RuntimeScalar("NUMBER", 2#)
    list.InitializeList "NUMBER", items
    context.Initialize "BOUNDS", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    context.AddField "VALUES", AT_RuntimeOK(list)
    context.AddField "UNUSED", AT_RuntimeOK(list)
    limits.MaxListItems = 1
    Set result = evaluator.Evaluate(compiler.Compile("SUM(@VALUES) = 3", schema), context, limits)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_LIMIT_LIST", "field list bounded before returned clone"
    Set result = evaluator.Evaluate(compiler.Compile("TRUE", schema), context, limits)
    AT_Assert result.State = atEvalOK, "unused fields not copied or evaluated"
    limits.MaxListItems = 2
    Set result = evaluator.Evaluate(compiler.Compile("SUM(@VALUES) = 3", schema), context, limits)
    AT_Assert result.Value.Scalar = True, "field list exact limit allowed"
End Sub
