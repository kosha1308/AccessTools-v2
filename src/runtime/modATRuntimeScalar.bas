Attribute VB_Name = "modATRuntimeScalar"
Option Explicit

Public Function AT_ApplyScalar(ByVal node As ATAstNode, ByVal args As Collection, ByVal context As ATEvaluationContext, ByVal budget As ATEvalBudget) As ATEvalResult
    Dim a As ATRuntimeValue, b As ATRuntimeValue, c As ATRuntimeValue, result As ATEvalResult
    Dim value As ATRuntimeValue, item As ATRuntimeValue, text As String, other As String
    Dim n As Double, x As Double, y As Double, precision As Double, cmp As Long, answer As Boolean, i As Long
    On Error GoTo Failed
    Set a = args(1)
    If args.Count > 1 Then Set b = args(2)
    If args.Count > 2 Then Set c = args(3)
    If AT_RuntimeDomainType(a.DataType) Then
        Set AT_ApplyScalar = AT_ApplyDomain(node, args, budget)
        Exit Function
    End If
    Select Case node.PrimitiveId
        Case "ANY", "NONE", "COVER", "EXACT", "OVERLAPS", "IN"
            Set AT_ApplyScalar = AT_ApplySet(node, args, context, budget)
            Exit Function
        Case "CONTAINS"
            If a.DataType <> "TEXT" Then
                Set AT_ApplyScalar = AT_ApplySet(node, args, context, budget)
                Exit Function
            End If
    End Select
    Select Case node.PrimitiveId
        Case "OP_EQUALS", "OP_NOT_EQUALS"
            answer = AT_ValuesEqual(a, b, context, budget)
            If node.PrimitiveId = "OP_NOT_EQUALS" Then answer = Not answer
            Set value = AT_RuntimeScalar("BOOL", answer)
        Case "GREATER", "LESS", "GREATER_EQUAL", "LESS_EQUAL"
            cmp = Compare(a, b, context)
            Select Case node.PrimitiveId
                Case "GREATER": answer = cmp > 0
                Case "LESS": answer = cmp < 0
                Case "GREATER_EQUAL": answer = cmp >= 0
                Case "LESS_EQUAL": answer = cmp <= 0
            End Select
            Set value = AT_RuntimeScalar("BOOL", answer)
        Case "BETWEEN"
            answer = Compare(a, b, context) >= 0 And Compare(a, c, context) <= 0
            Set value = AT_RuntimeScalar("BOOL", answer)
        Case "ADD", "SUBTRACT", "MULTIPLY", "DIVIDE", "REMAINDER", "UNARY_PLUS", "UNARY_MINUS", "ABS", "ROUND"
            x = a.Scalar
            If Not b Is Nothing Then y = b.Scalar
            Select Case node.PrimitiveId
                Case "ADD": n = x + y
                Case "SUBTRACT": n = x - y
                Case "MULTIPLY"
                    n = x * y
                    If x <> 0# And y <> 0# And n = 0# Then GoTo NumberRange
                Case "DIVIDE", "REMAINDER"
                    If y = 0# Then
                        Set AT_ApplyScalar = AT_RuntimeFailure(atEvalError, "RUNTIME_DIVISION_BY_ZERO", node)
                        Exit Function
                    End If
                    n = x / y
                    If x <> 0# And n = 0# Then GoTo NumberRange
                    If node.PrimitiveId = "REMAINDER" Then n = x - Fix(n) * y
                Case "UNARY_PLUS": n = x
                Case "UNARY_MINUS": n = -x
                Case "ABS": n = Abs(x)
                Case "ROUND"
                    If y <> Fix(y) Or y < 0 Or y > 28 Then
                        Set AT_ApplyScalar = AT_RuntimeFailure(atEvalError, "RUNTIME_ARGUMENT_RANGE", node)
                        Exit Function
                    End If
                    n = RoundEven(x, CLng(y))
            End Select
            Set value = AT_RuntimeScalar("NUMBER", n)
        Case "LOWER", "UPPER", "TRIM", "LEN", "STARTS_WITH", "ENDS_WITH", "CONTAINS"
            text = a.Scalar
            Select Case node.PrimitiveId
                Case "LOWER": text = AT_AsciiCase(text, False)
                Case "UPPER": text = AT_AsciiCase(text, True)
                Case "TRIM": text = Trim$(text)
                Case "LEN"
                    Set value = AT_RuntimeScalar("NUMBER", CDbl(Len(text)))
                Case Else
                    text = AT_TextKey(text, context.ComparisonMode)
                    other = AT_TextKey(b.Scalar, context.ComparisonMode)
                    Select Case node.PrimitiveId
                        Case "STARTS_WITH": answer = StrComp(Left$(text, Len(other)), other, vbBinaryCompare) = 0
                        Case "ENDS_WITH": answer = StrComp(Right$(text, Len(other)), other, vbBinaryCompare) = 0
                        Case "CONTAINS": answer = InStr(1, text, other, vbBinaryCompare) > 0
                    End Select
                    Set value = AT_RuntimeScalar("BOOL", answer)
            End Select
            If value Is Nothing Then
                budget.CheckText Len(text)
                Set value = AT_RuntimeScalar("TEXT", text)
            End If
        Case "MIN", "MAX", "SUM", "AVG"
            If Left$(a.DataType, 5) = "LIST<" Then
                If a.Count = 0 And node.PrimitiveId <> "SUM" Then
                    Set AT_ApplyScalar = AT_RuntimeFailure(atEvalError, "AGGREGATE_EMPTY", node)
                    Exit Function
                End If
                For i = 1 To a.Count
                    budget.Step
                    Set item = a.Item(i)
                    Set result = AT_Usable(item, node)
                    If result.State <> atEvalOK Then
                        Set AT_ApplyScalar = result
                        Exit Function
                    End If
                    If node.PrimitiveId = "SUM" Or node.PrimitiveId = "AVG" Then
                        n = n + item.Scalar
                    ElseIf value Is Nothing Then
                        Set value = item
                    Else
                        cmp = Compare(item, value, context)
                        If (node.PrimitiveId = "MIN" And cmp < 0) Or (node.PrimitiveId = "MAX" And cmp > 0) Then Set value = item
                    End If
                Next i
                If node.PrimitiveId = "AVG" Then
                    x = n
                    n = n / a.Count
                    If x <> 0# And n = 0# Then GoTo NumberRange
                End If
                If value Is Nothing Then Set value = AT_RuntimeScalar("NUMBER", n)
            Else
                cmp = Compare(a, b, context)
                Set value = a
                If (node.PrimitiveId = "MIN" And cmp > 0) Or (node.PrimitiveId = "MAX" And cmp < 0) Then Set value = b
            End If
        Case "DATE_LITERAL"
            Set AT_ApplyScalar = AT_NormalizeScalar("DATE", a.Scalar, node)
            Exit Function
        Case "DATE_DIFF"
            Set value = AT_RuntimeScalar("NUMBER", CDbl(b.Scalar) - CDbl(a.Scalar))
        Case Else
            Set AT_ApplyScalar = AT_RuntimeFailure(atEvalError, "RUNTIME_UNSUPPORTED_PRIMITIVE", node)
            Exit Function
    End Select
    Set AT_ApplyScalar = AT_RuntimeOK(value)
    Exit Function
NumberRange:
    Set AT_ApplyScalar = AT_RuntimeFailure(atEvalError, "RUNTIME_NUMBER_RANGE", node)
    Exit Function
Failed:
    If Err.Number = 6 Then
        Set AT_ApplyScalar = AT_RuntimeFailure(atEvalError, "RUNTIME_NUMBER_RANGE", node)
    Else
        Set AT_ApplyScalar = AT_RuntimeFailure(atEvalError, AT_StableErrorCode(Err.Description, "RUNTIME_INTERNAL_ERROR"), node)
    End If
End Function

Private Function Compare(ByVal a As ATRuntimeValue, ByVal b As ATRuntimeValue, ByVal context As ATEvaluationContext) As Long
    If a.DataType <> b.DataType Then Err.Raise vbObjectError + 2530, , "RUNTIME_TYPE_MISMATCH"
    Select Case a.DataType
        Case "TEXT": Compare = StrComp(AT_TextKey(a.Scalar, context.ComparisonMode), AT_TextKey(b.Scalar, context.ComparisonMode), vbBinaryCompare)
        Case "NUMBER", "DATE"
            If a.Scalar < b.Scalar Then Compare = -1
            If a.Scalar > b.Scalar Then Compare = 1
        Case Else: Err.Raise vbObjectError + 2530, , "RUNTIME_TYPE_MISMATCH"
    End Select
End Function

Private Function RoundEven(ByVal value As Double, ByVal precision As Long) As Double
    Dim factor As Double, scaled As Double, whole As Double, fraction As Double
    If Abs(value) >= 9007199254740992# Then
        RoundEven = value
        Exit Function
    End If
    factor = 10# ^ precision
    scaled = Abs(value) * factor
    If scaled >= 9007199254740992# Then
        RoundEven = value
        Exit Function
    End If
    whole = Fix(scaled)
    fraction = scaled - whole
    If fraction > 0.5 Or (fraction = 0.5 And whole - 2# * Fix(whole / 2#) <> 0#) Then whole = whole + 1#
    RoundEven = Sgn(value) * (whole / factor)
End Function
