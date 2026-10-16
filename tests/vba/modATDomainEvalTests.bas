Attribute VB_Name = "modATDomainEvalTests"
Option Explicit

Public Sub AT_TestDomainEval()
    Dim schema As New ATCompileContext, context As New ATEvaluationContext, compiler As New ATCompiler, evaluator As New ATEvaluator
    Dim named As New ATNamedExpression, result As ATEvalResult, expression As Variant, parts As Variant
    Dim numbers As New ATRuntimeValue, items As New Collection
    schema.AddField "A", "IP_EXPRESSION": schema.AddField "R", "IP_EXPRESSION"
    schema.AddField "P", "PORT_SET": schema.AddField "N", "NUMBER": schema.AddField "NUMS", "LIST<NUMBER>"
    context.Initialize "DOMAINS", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    context.AddField "A", AT_ParseDomain("IP_EXPRESSION", "192.0.2.1-10", Nothing)
    context.AddField "R", AT_ParseDomain("IP_EXPRESSION", "192.0.2.0/24", Nothing)
    context.AddField "P", AT_ParseDomain("PORT_SET", "443,8443", Nothing)
    context.AddField "N", AT_RuntimeOK(AT_RuntimeScalar("NUMBER", 443#))
    items.Add AT_RuntimeScalar("NUMBER", 443#)
    items.Add AT_RuntimeScalar("NUMBER", 8443#)
    numbers.InitializeList "NUMBER", items
    context.AddField "NUMS", AT_RuntimeOK(numbers)
    named.Initialize "HAS_PORT", "BOOL", "@P CONTAINS $VALUE"
    named.AddParameter "VALUE", "PORT_SET", "443"
    schema.AddNamed named
    Set named = New ATNamedExpression
    named.Initialize "HAS_COMPUTED", "BOOL", "@P CONTAINS $VALUE"
    named.AddParameter "VALUE", "PORT_SET", "SUM([400,43])"
    schema.AddNamed named
    For Each expression In Array("@A COVER @R", "@A IN @R", "@A CONTAINED_BY @R", "@A INTERSECTS @R", "@A ANY @R", "@R CONTAINS @A", "NOT (@A EXACT @R)", "NOT (@A NONE @R)", "@P COVER:443,8443", "@P CONTAINS:443", "@P ANY:443", "@P CONTAINS @N", "@P CONTAINS (400 + 43)", "$HAS_PORT(443)", "$HAS_PORT")
        Set result = evaluator.Evaluate(compiler.Compile(CStr(expression), schema), context)
        AT_Assert result.State = atEvalOK, "domain evaluation state: " & CStr(expression) & AT_RuntimeTestDetails(result)
        If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "domain evaluation result: " & CStr(expression)
    Next expression
    For Each expression In Array("@P CONTAINS ABS(-443)", "@P CONTAINS SUM([400,43])", "@P EXACT @NUMS", "$HAS_COMPUTED")
        Set result = evaluator.Evaluate(compiler.Compile(CStr(expression), schema), context)
        AT_Assert result.State = atEvalOK, "function/list computed port conversion state" & AT_RuntimeTestDetails(result)
        If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "function/list computed port conversion value"
    Next expression
    For Each expression In Array("@R COVER @A", "@P COVER:443", "@P EXACT:443", "@P NONE:8443")
        Set result = evaluator.Evaluate(compiler.Compile(CStr(expression), schema), context)
        AT_Assert result.State = atEvalOK, "partial coverage yields typed Boolean"
        If result.State = atEvalOK Then AT_Assert result.Value.Scalar = False, "partial coverage never PASS"
    Next expression
    For Each expression In Array("@P COVER (1 / 2)", "@P ANY (65535 + 1)", "$HAS_PORT(1 / 2)", "@P ANY [443,(1 / 2)]")
        Set result = evaluator.Evaluate(compiler.Compile(CStr(expression), schema), context)
        AT_Assert result.State = atEvalError, "computed invalid port conversion rejected"
    Next expression
    Set result = evaluator.Evaluate(compiler.Compile("TRUE OR (@P COVER (1 / 2))", schema), context)
    AT_Assert result.State = atEvalOK, "invalid port conversion skipped branch state"
    If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "invalid port conversion in skipped branch remains skipped"
    Set result = AT_NormalizeScalar("IP_EXPRESSION", "192.0.2.x", Nothing)
    AT_Assert result.State = atEvalOK, "domain normalizer integrated"
    If result.State = atEvalOK Then AT_Assert result.Value.ValueState = atValuePresent, "domain normalizer yields usable domain"
End Sub
