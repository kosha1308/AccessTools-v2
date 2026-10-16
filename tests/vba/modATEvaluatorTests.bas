Attribute VB_Name = "modATEvaluatorTests"
Option Explicit
Public Function AT_EvalFixture(ByVal expression As String, Optional ByVal value As ATEvalResult = Nothing, Optional ByVal requiredType As String = "BOOL", Optional ByVal limits As ATEvalLimits = Nothing) As ATEvalResult
    Dim schema As New ATCompileContext, compiler As New ATCompiler, context As New ATEvaluationContext, evaluator As New ATEvaluator
    schema.AddField "X", "BOOL"
    context.Initialize "TEST", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    If Not value Is Nothing Then context.AddField "X", value
    Set AT_EvalFixture = evaluator.Evaluate(compiler.Compile(expression, schema, , requiredType), context, limits)
End Function
Public Sub AT_TestEvaluator()
    Dim result As ATEvalResult, incoming As ATEvalResult, states As Variant, state As Variant
    Dim left As Variant, op As Variant, expression As String, expected As Variant, shorted As Boolean
    Dim schema As New ATCompileContext, context As New ATEvaluationContext, compiler As New ATCompiler, evaluator As New ATEvaluator
    Dim compiled As ATCompileResult, named As New ATNamedExpression, trace As Collection
    For Each op In Array("AND", "OR")
        For Each left In Array(False, True)
            shorted = (op = "AND" And Not left) Or (op = "OR" And left)
            For Each state In Array(atEvalError, atEvalSkipped, atEvalManualReview)
                Set incoming = AT_RuntimeFailure(CLng(state), "TEST_INCOMING", Nothing)
                expression = IIf(left, "TRUE", "FALSE") & " " & op & " @X"
                Set result = AT_EvalFixture(expression, incoming)
                If shorted Then
                    AT_Assert result.State = atEvalOK, "logic short circuit state"
                    If result.State = atEvalOK Then AT_Assert result.Value.Scalar = left, "logic short circuit value"
                    Set trace = result.Trace
                    AT_Assert trace(2).Status = "NOT_EVALUATED", "short circuit explanation"
                Else
                    AT_Assert result.State = state, "logic propagates right non-OK"
                End If
                Set result = AT_EvalFixture("@X " & op & " TRUE", incoming)
                AT_Assert result.State = state, "logic propagates left non-OK"
            Next state
            For Each expected In Array(False, True)
                expression = IIf(left, "TRUE", "FALSE") & " " & op & " " & IIf(expected, "TRUE", "FALSE")
                Set result = AT_EvalFixture(expression)
                If op = "AND" Then
                    AT_Assert result.Value.Scalar = (left And expected), "AND truth table"
                Else
                    AT_Assert result.Value.Scalar = (left Or expected), "OR truth table"
                End If
            Next expected
        Next left
    Next op
    Set result = AT_EvalFixture("FALSE AND (1 / 0 = 1)")
    AT_Assert result.State = atEvalOK, "division in skipped branch not evaluated"
    For Each op In Array("EMPTY", "NOT_EMPTY", "IS NULL", "IS NOT NULL", "VALID", "INVALID")
        For state = atValuePresent To atValueNull
            If state = atValuePresent Then
                Set incoming = AT_RuntimeOK(AT_RuntimeScalar("BOOL", True))
            Else
                Set incoming = AT_RuntimeOK(AT_RuntimeState("BOOL", CLng(state)))
            End If
            Set result = AT_EvalFixture("@X " & op, incoming)
            If state = atValueMissing Or (state = atValueInvalid And op <> "VALID" And op <> "INVALID") Then
                AT_Assert result.State = atEvalError, "inspector rejects unusable value"
            Else
                Select Case op
                    Case "EMPTY": expected = (state = atValueEmpty)
                    Case "NOT_EMPTY", "VALID": expected = (state = atValuePresent)
                    Case "IS NULL": expected = (state = atValueNull)
                    Case "IS NOT NULL": expected = (state <> atValueNull)
                    Case "INVALID": expected = (state = atValueInvalid)
                End Select
                AT_Assert result.State = atEvalOK, "inspector state"
                If result.State = atEvalOK Then AT_Assert result.Value.Scalar = expected, "inspector matrix value"
            End If
        Next state
        For Each state In Array(atEvalError, atEvalSkipped, atEvalManualReview)
            Set result = AT_EvalFixture("@X " & op, AT_RuntimeFailure(CLng(state), "TEST_INCOMING", Nothing))
            AT_Assert result.State = state, "inspector never hides non-OK"
        Next state
    Next op
    Set result = AT_EvalFixture("COALESCE(@X,FALSE)", AT_RuntimeOK(AT_RuntimeState("BOOL", atValueNull)))
    AT_Assert result.Value.Scalar = False, "coalesce explicit null fallback"
    Set result = AT_EvalFixture("COALESCE(@X,1 / 0 = 1)", AT_RuntimeOK(AT_RuntimeScalar("BOOL", True)))
    AT_Assert result.State = atEvalOK, "coalesce fallback lazy"
    Set result = AT_EvalFixture("@X", AT_RuntimeOK(AT_RuntimeScalar("TEXT", "wrong")))
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_TYPE_MISMATCH", "field type mismatch"
    context.Initialize "NAMES", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    named.Initialize "PLUS", "NUMBER", "$A + $B"
    named.AddParameter "A", "NUMBER"
    named.AddParameter "B", "NUMBER", "$A + 1"
    schema.AddNamed named
    Set compiled = compiler.Compile("$PLUS(2) = 5 AND $PLUS(3) = 7", schema)
    Set result = evaluator.Evaluate(compiled, context)
    AT_Assert result.State = atEvalOK, "named defaults run"
    If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "named defaults and frame isolation"
    Set named = New ATNamedExpression
    named.Initialize "INNER", "NUMBER", "$X"
    named.AddParameter "X", "NUMBER"
    schema.AddNamed named
    Set named = New ATNamedExpression
    named.Initialize "OUTER", "NUMBER", "$B"
    named.AddParameter "A", "NUMBER"
    named.AddParameter "B", "NUMBER", "$INNER($A)"
    schema.AddNamed named
    Set result = evaluator.Evaluate(compiler.Compile("$OUTER(2) = 2", schema), context)
    AT_Assert result.State = atEvalOK, "supplied argument inside default uses caller frame"
    If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "nested default supplied argument value"
    AT_Assert Not AT_RuntimeVersionSupported("2.0.0", "1.0.0"), "runtime rejects unknown DSL version"
    AT_Assert Not AT_RuntimeVersionSupported("1.0.0", "2.0.0"), "runtime rejects unknown schema version"
End Sub
