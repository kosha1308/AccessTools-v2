Attribute VB_Name = "modATDomainEval"
Option Explicit

Public Function AT_ConvertPortValue(ByVal value As ATRuntimeValue, ByVal node As ATAstNode, ByVal budget As ATEvalBudget) As ATEvalResult
    Dim intervals As New ATIntervalSet, keys As New Collection, item As ATRuntimeValue
    Dim result As ATEvalResult, converted As New ATRuntimeValue, count As Long, i As Long, number As Double
    On Error GoTo Failed
    Set result = AT_Usable(value, node, budget)
    If result.State <> atEvalOK Then
        Set AT_ConvertPortValue = result
        Exit Function
    End If
    If value.DataType <> "NUMBER" And value.DataType <> "LIST<NUMBER>" Then Err.Raise vbObjectError + 2620, , "RUNTIME_TYPE_MISMATCH"
    intervals.Initialize 65535#
    count = 1
    If value.DataType = "LIST<NUMBER>" Then count = value.Count
    budget.CheckList count
    For i = 1 To count
        budget.Step
        Set item = value
        If value.DataType = "LIST<NUMBER>" Then Set item = value.Item(i)
        number = item.Scalar
        If number < 0# Or number > 65535# Or number <> Fix(number) Then Err.Raise vbObjectError + 2621, , "RUNTIME_INVALID_PORT"
        AT_AddDomainRange intervals, keys, number, number, budget
    Next i
    intervals.Freeze
    converted.InitializeDomain "PORT_SET", intervals, keys
    Set AT_ConvertPortValue = AT_RuntimeOK(converted)
    Exit Function
Failed:
    Set AT_ConvertPortValue = AT_RuntimeFailure(atEvalError, AT_StableErrorCode(Err.Description, "RUNTIME_INTERNAL_ERROR"), node)
End Function

Public Function AT_ApplyDomain(ByVal node As ATAstNode, ByVal args As Collection, ByVal budget As ATEvalBudget) As ATEvalResult
    Dim a As ATRuntimeValue, b As ATRuntimeValue, answer As Boolean
    Set a = args(1): Set b = args(2)
    If a.DataType <> b.DataType Then
        Set AT_ApplyDomain = AT_RuntimeFailure(atEvalError, "RUNTIME_TYPE_MISMATCH", node)
        Exit Function
    End If
    Select Case node.PrimitiveId
        Case "ANY", "OVERLAPS", "INTERSECTS": answer = a.Domain.Overlaps(b.Domain, budget)
        Case "NONE": answer = Not a.Domain.Overlaps(b.Domain, budget)
        Case "COVER", "CONTAINED_BY", "IN": answer = a.Domain.IsSubsetOf(b.Domain, budget)
        Case "CONTAINS": answer = b.Domain.IsSubsetOf(a.Domain, budget)
        Case "EXACT", "OP_EQUALS", "OP_NOT_EQUALS"
            answer = a.Domain.IsSubsetOf(b.Domain, budget)
            If answer Then answer = b.Domain.IsSubsetOf(a.Domain, budget)
            If node.PrimitiveId = "OP_NOT_EQUALS" Then answer = Not answer
        Case Else
            Set AT_ApplyDomain = AT_RuntimeFailure(atEvalError, "RUNTIME_UNSUPPORTED_PRIMITIVE", node)
            Exit Function
    End Select
    Set AT_ApplyDomain = AT_RuntimeOK(AT_RuntimeScalar("BOOL", answer))
End Function
