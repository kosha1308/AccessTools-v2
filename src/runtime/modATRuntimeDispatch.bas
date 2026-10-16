Attribute VB_Name = "modATRuntimeDispatch"
Option Explicit

Public Function AT_RuntimeVersionSupported(ByVal dslVersion As String, ByVal schemaVersion As String) As Boolean
    AT_RuntimeVersionSupported = (dslVersion = "1.0.0" And schemaVersion = "1.0.0")
End Function
Public Function AT_RuntimeSupports(ByVal node As ATAstNode) As Boolean
    Dim child As ATAstNode
    If Len(node.ConversionNormalizerId) > 0 Then Exit Function
    Select Case node.PrimitiveId
        Case "EMPTY", "NOT_EMPTY", "VALID", "INVALID", "IS_NULL", "IS_NOT_NULL"
            AT_RuntimeSupports = (node.Children.Count = 1)
            Exit Function
    End Select
    If Not AT_RuntimeTypeSupported(node.DataType) Then Exit Function
    For Each child In node.Children
        If Not AT_RuntimeTypeSupported(child.DataType) Then Exit Function
        If Len(child.ConversionNormalizerId) > 0 Then Exit Function
    Next child
    Select Case node.PrimitiveId
        Case "AND", "OR", "NOT", "COALESCE", "ADD", "SUBTRACT", "MULTIPLY", "DIVIDE", "REMAINDER", "UNARY_PLUS", "UNARY_MINUS", "OP_EQUALS", "OP_NOT_EQUALS", "GREATER", "LESS", "GREATER_EQUAL", "LESS_EQUAL", "BETWEEN"
            AT_RuntimeSupports = True
        Case "LOWER", "UPPER", "TRIM", "LEN", "STARTS_WITH", "ENDS_WITH", "CONTAINS", "ABS", "ROUND", "MIN", "MAX", "SUM", "AVG", "DATE_LITERAL", "TODAY", "DATE_DIFF", "IN", "ANY", "NONE", "COVER", "EXACT", "OVERLAPS"
            AT_RuntimeSupports = True
    End Select
End Function
Public Function AT_RuntimeTypeSupported(ByVal typeName As String) As Boolean
    If AT_RuntimeScalarType(typeName) Then
        AT_RuntimeTypeSupported = True
    ElseIf Left$(typeName, 5) = "LIST<" And Right$(typeName, 1) = ">" Then
        AT_RuntimeTypeSupported = AT_RuntimeScalarType(Mid$(typeName, 6, Len(typeName) - 6))
    End If
End Function
Public Function AT_Usable(ByVal value As ATRuntimeValue, ByVal node As ATAstNode, Optional ByVal budget As ATEvalBudget = Nothing) As ATEvalResult
    Dim i As Long, result As ATEvalResult
    Select Case value.ValueState
        Case atValuePresent
            If Left$(value.DataType, 5) = "LIST<" Then
                For i = 1 To value.Count
                    If Not budget Is Nothing Then budget.Step
                    Set result = AT_Usable(value.Item(i), node, budget)
                    If result.State <> atEvalOK Then
                        Set AT_Usable = result
                        Exit Function
                    End If
                Next i
            End If
            Set AT_Usable = AT_RuntimeOK(value)
        Case atValueMissing: Set AT_Usable = AT_RuntimeFailure(atEvalError, "RUNTIME_FIELD_MISSING", node)
        Case atValueInvalid: Set AT_Usable = AT_RuntimeFailure(atEvalError, "RUNTIME_INVALID_VALUE", node)
        Case Else: Set AT_Usable = AT_RuntimeFailure(atEvalError, "RUNTIME_VALUE_UNUSABLE", node)
    End Select
End Function
Public Function AT_EvalStateName(ByVal state As ATEvalState) As String
    Select Case state
        Case atEvalOK: AT_EvalStateName = "OK"
        Case atEvalError: AT_EvalStateName = "ERROR"
        Case atEvalSkipped: AT_EvalStateName = "SKIPPED"
        Case atEvalManualReview: AT_EvalStateName = "MANUAL_REVIEW"
    End Select
End Function
