Attribute VB_Name = "modATRuntimeSets"
Option Explicit

Public Function AT_ValuesEqual(ByVal a As ATRuntimeValue, ByVal b As ATRuntimeValue, ByVal context As ATEvaluationContext, ByVal budget As ATEvalBudget) As Boolean
    Dim i As Long
    budget.Step
    If a.DataType <> b.DataType Then Err.Raise vbObjectError + 2530, , "RUNTIME_TYPE_MISMATCH"
    If a.ValueState <> atValuePresent Or b.ValueState <> atValuePresent Then Err.Raise vbObjectError + 2531, , "RUNTIME_VALUE_UNUSABLE"
    If Left$(a.DataType, 5) = "LIST<" Then
        If a.Count <> b.Count Then Exit Function
        For i = 1 To a.Count
            If Not AT_ValuesEqual(a.Item(i), b.Item(i), context, budget) Then Exit Function
        Next i
        AT_ValuesEqual = True
    ElseIf a.DataType = "TEXT" Then
        AT_ValuesEqual = (StrComp(AT_TextKey(a.Scalar, context.ComparisonMode), AT_TextKey(b.Scalar, context.ComparisonMode), vbBinaryCompare) = 0)
    Else
        AT_ValuesEqual = (a.Scalar = b.Scalar)
    End If
End Function

Private Function ContainsValue(ByVal list As ATRuntimeValue, ByVal value As ATRuntimeValue, ByVal context As ATEvaluationContext, ByVal budget As ATEvalBudget) As Boolean
    Dim i As Long
    For i = 1 To list.Count
        If AT_ValuesEqual(list.Item(i), value, context, budget) Then
            ContainsValue = True
            Exit Function
        End If
    Next i
End Function

Private Function IsSubset(ByVal a As ATRuntimeValue, ByVal b As ATRuntimeValue, ByVal context As ATEvaluationContext, ByVal budget As ATEvalBudget) As Boolean
    Dim i As Long
    For i = 1 To a.Count
        budget.Step
        If Not ContainsValue(b, a.Item(i), context, budget) Then Exit Function
    Next i
    IsSubset = True
End Function

Public Function AT_ApplySet(ByVal node As ATAstNode, ByVal args As Collection, ByVal context As ATEvaluationContext, ByVal budget As ATEvalBudget) As ATEvalResult
    Dim a As ATRuntimeValue, b As ATRuntimeValue, answer As Boolean, i As Long
    Set a = args(1): Set b = args(2)
    Select Case node.PrimitiveId
        Case "IN": answer = ContainsValue(b, a, context, budget)
        Case "COVER": answer = IsSubset(a, b, context, budget)
        Case "CONTAINS": answer = IsSubset(b, a, context, budget)
        Case "EXACT"
            answer = IsSubset(a, b, context, budget)
            If answer Then answer = IsSubset(b, a, context, budget)
        Case "ANY", "OVERLAPS", "NONE"
            For i = 1 To a.Count
                budget.Step
                If ContainsValue(b, a.Item(i), context, budget) Then
                    answer = True
                    Exit For
                End If
            Next i
            If node.PrimitiveId = "NONE" Then answer = Not answer
        Case Else
            Set AT_ApplySet = AT_RuntimeFailure(atEvalError, "RUNTIME_UNSUPPORTED_PRIMITIVE", node)
            Exit Function
    End Select
    Set AT_ApplySet = AT_RuntimeOK(AT_RuntimeScalar("BOOL", answer))
End Function
