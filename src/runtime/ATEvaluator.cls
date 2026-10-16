VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATEvaluator"
Option Explicit
Private mContext As ATEvaluationContext
Private mBudget As ATEvalBudget

Public Function Evaluate(ByVal compiled As ATCompileResult, ByVal context As ATEvaluationContext, Optional ByVal limits As ATEvalLimits = Nothing) As ATEvalResult
    Dim result As ATEvalResult, final As New ATEvalResult, frame As New Collection
    Dim code As String, value As ATRuntimeValue
    Set mBudget = Nothing
    Set mContext = Nothing
    On Error GoTo Failed
    If compiled Is Nothing Then GoTo NoCompile
    If compiled.HasErrors Or compiled.Root Is Nothing Then GoTo NoCompile
    If Not AT_RuntimeVersionSupported(compiled.DSLVersion, compiled.SchemaVersion) Then
        Set Evaluate = AT_RuntimeFailure(atEvalError, "RUNTIME_VERSION_UNSUPPORTED", Nothing)
        Exit Function
    End If
    If context Is Nothing Then Err.Raise vbObjectError + 2504, , "CONFIG_CONTEXT_REQUIRED"
    context.Validate
    If limits Is Nothing Then Set limits = New ATEvalLimits
    Set mBudget = New ATEvalBudget
    mBudget.Initialize limits
    Set mContext = context.Snapshot
    Set result = Visit(compiled.Root, frame)
    If result.State = atEvalOK Then Set value = result.Value
    final.Initialize result.State, value, result.Diagnostics, mBudget.Trace
    Set Evaluate = final
    Exit Function
NoCompile:
    Set Evaluate = AT_RuntimeFailure(atEvalError, "RUNTIME_COMPILE_REQUIRED", Nothing)
    Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "RUNTIME_INTERNAL_ERROR")
    Set Evaluate = AT_RuntimeFailure(atEvalError, code, Nothing)
End Function

Private Function Visit(ByVal node As ATAstNode, ByVal frame As Collection) As ATEvalResult
    Dim result As ATEvalResult, entered As Boolean, code As String
    On Error GoTo Failed
    mBudget.Step
    mBudget.EnterNode
    entered = True
    Set result = EvaluateNode(node, frame)
    LogNode node, AT_EvalStateName(result.State), result
    mBudget.LeaveNode
    Set Visit = result
    Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "RUNTIME_INTERNAL_ERROR")
    If entered Then mBudget.LeaveNode
    If Not result Is Nothing Then
        If result.State <> atEvalOK Then
            Set Visit = WithSecondary(result, code, node)
            Exit Function
        End If
    End If
    Set Visit = AT_RuntimeFailure(atEvalError, code, node)
End Function

Private Function WithSecondary(ByVal result As ATEvalResult, ByVal code As String, ByVal node As ATAstNode) As ATEvalResult
    Dim diagnostics As Collection, entry As ATDiagnostic, final As New ATEvalResult
    Set diagnostics = result.Diagnostics
    For Each entry In diagnostics
        If entry.Code = code Then
            Set WithSecondary = result
            Exit Function
        End If
    Next entry
    diagnostics.Add AT_RuntimeDiagnostic(code, node)
    final.Initialize result.State, Nothing, diagnostics, Nothing
    Set WithSecondary = final
End Function

Private Sub LogNode(ByVal node As ATAstNode, ByVal status As String, Optional ByVal result As ATEvalResult = Nothing)
    Dim entry As New ATTraceEntry, valueState As String, reason As String
    If Not result Is Nothing Then
        If result.State = atEvalOK Then valueState = CStr(result.Value.ValueState)
        If result.Diagnostics.Count > 0 Then reason = result.Diagnostics(1).Code
    End If
    entry.Initialize node.SourceId, node.Start, node.Length, node.PrimitiveId, status, valueState, reason, mBudget.TraceCount + 1
    mBudget.AddTrace entry
End Sub

Private Function EvaluateNode(ByVal node As ATAstNode, ByVal frame As Collection) As ATEvalResult
    Dim result As ATEvalResult, right As ATEvalResult, value As ATRuntimeValue, items As New Collection
    Dim child As ATAstNode, args As New Collection, localFrame As Collection, binding As ATArgumentBinding
    Dim i As Long, state As ATValueState, answer As Boolean, param As ATEvalResult
    Select Case node.Kind
        Case "NUMBER"
            mBudget.CheckText Len(node.Text)
            Set EvaluateNode = AT_ParseDecimal(node.Text, node)
            Exit Function
        Case "STRING"
            mBudget.CheckText Len(node.Text)
            Set EvaluateNode = AT_RuntimeOK(AT_RuntimeScalar("TEXT", node.Text))
            Exit Function
        Case "BOOL"
            Set EvaluateNode = AT_RuntimeOK(AT_RuntimeScalar("BOOL", node.Text = "TRUE"))
            Exit Function
        Case "FIELD"
            Set result = mContext.ResolveField(node.Text, mBudget)
            If result.State = atEvalError Then
                If result.Diagnostics.Count > 0 Then
                    If result.Diagnostics(1).Code = "RUNTIME_FIELD_MISSING" Then Set result = AT_RuntimeFailure(atEvalError, "RUNTIME_FIELD_MISSING", node)
                End If
            End If
            If result.State = atEvalOK Then
                Set value = result.Value
                If value.DataType <> node.DataType Then
                    Set EvaluateNode = AT_RuntimeFailure(atEvalError, "RUNTIME_TYPE_MISMATCH", node)
                    Exit Function
                End If
            End If
            Set EvaluateNode = result
            Exit Function
        Case "LIST"
            mBudget.CheckList node.Children.Count
            For Each child In node.Children
                mBudget.Step
                Set result = Visit(child, frame)
                If result.State <> atEvalOK Then
                    Set EvaluateNode = result
                    Exit Function
                End If
                items.Add result.Value
            Next child
            Set value = New ATRuntimeValue
            value.InitializeList Mid$(node.DataType, 6, Len(node.DataType) - 6), items
            Set EvaluateNode = AT_RuntimeOK(value)
            Exit Function
        Case "NAMED"
            If node.Expansion Is Nothing Then
                On Error Resume Next
                Set param = frame(node.Text)
                On Error GoTo 0
                If param Is Nothing Then
                    Set EvaluateNode = AT_RuntimeFailure(atEvalError, "RUNTIME_PARAMETER_MISSING", node)
                Else
                    Set EvaluateNode = param.Clone
                End If
                Exit Function
            End If
            Set localFrame = New Collection
            For Each binding In node.ParameterBindings
                If Len(binding.NormalizerId) > 0 Then
                    Set EvaluateNode = AT_RuntimeFailure(atEvalError, "RUNTIME_UNSUPPORTED_NORMALIZER", binding.Expression)
                    Exit Function
                End If
                If binding.IsDefault Then
                    Set result = Visit(binding.Expression, localFrame)
                Else
                    Set result = Visit(binding.Expression, frame)
                End If
                If result.State <> atEvalOK Then
                    Set EvaluateNode = result
                    Exit Function
                End If
                localFrame.Add result, binding.Id
            Next binding
            Set EvaluateNode = Visit(node.Expansion, localFrame)
            Exit Function
        Case "CALL", "UNARY", "BINARY", "RELATION"
        Case Else
            Set EvaluateNode = AT_RuntimeFailure(atEvalError, "RUNTIME_UNSUPPORTED_NODE", node)
            Exit Function
    End Select
    If Not AT_RuntimeSupports(node) Then
        Set EvaluateNode = AT_RuntimeFailure(atEvalError, "RUNTIME_UNSUPPORTED_PRIMITIVE", node)
        Exit Function
    End If
    If node.PrimitiveId = "TODAY" Then
        Set EvaluateNode = AT_NormalizeScalar("DATE", mContext.RunDate, node)
        Exit Function
    End If
    Set result = Visit(node.Children(1), frame)
    If result.State <> atEvalOK Then
        Set EvaluateNode = SkipRemaining(node, 2, result)
        Exit Function
    End If
    Set value = result.Value
    Select Case node.PrimitiveId
        Case "EMPTY", "NOT_EMPTY", "VALID", "INVALID", "IS_NULL", "IS_NOT_NULL"
            state = value.ValueState
            If state = atValueMissing Then
                Set EvaluateNode = AT_RuntimeFailure(atEvalError, "RUNTIME_FIELD_MISSING", node)
                Exit Function
            End If
            If state = atValueInvalid And node.PrimitiveId <> "VALID" And node.PrimitiveId <> "INVALID" Then
                Set EvaluateNode = AT_RuntimeFailure(atEvalError, "RUNTIME_INVALID_VALUE", node)
                Exit Function
            End If
            Select Case node.PrimitiveId
                Case "EMPTY": answer = (state = atValueEmpty)
                Case "NOT_EMPTY", "VALID": answer = (state = atValuePresent)
                Case "INVALID": answer = (state = atValueInvalid)
                Case "IS_NULL": answer = (state = atValueNull)
                Case "IS_NOT_NULL": answer = (state <> atValueNull)
            End Select
            Set EvaluateNode = AT_RuntimeOK(AT_RuntimeScalar("BOOL", answer))
            Exit Function
        Case "COALESCE"
            If value.ValueState = atValueNull Then
                Set EvaluateNode = Visit(node.Children(2), frame)
            Else
                LogNode node.Children(2), "NOT_EVALUATED"
                Set EvaluateNode = result
            End If
            Exit Function
        Case "AND", "OR", "NOT"
            Set result = AT_Usable(value, node, mBudget)
            If result.State <> atEvalOK Then
                Set EvaluateNode = SkipRemaining(node, 2, result)
                Exit Function
            End If
            If node.PrimitiveId = "NOT" Then
                Set EvaluateNode = AT_RuntimeOK(AT_RuntimeScalar("BOOL", Not CBool(value.Scalar)))
                Exit Function
            End If
            If (node.PrimitiveId = "AND" And Not value.Scalar) Or (node.PrimitiveId = "OR" And value.Scalar) Then
                LogNode node.Children(2), "NOT_EVALUATED"
                Set EvaluateNode = result
                Exit Function
            End If
            Set right = Visit(node.Children(2), frame)
            If right.State = atEvalOK Then Set right = AT_Usable(right.Value, node, mBudget)
            Set EvaluateNode = right
            Exit Function
    End Select
    Set result = AT_Usable(value, node, mBudget)
    If result.State <> atEvalOK Then
        Set EvaluateNode = SkipRemaining(node, 2, result)
        Exit Function
    End If
    args.Add value
    For i = 2 To node.Children.Count
        Set result = Visit(node.Children(i), frame)
        If result.State <> atEvalOK Then
            Set EvaluateNode = SkipRemaining(node, i + 1, result)
            Exit Function
        End If
        Set result = AT_Usable(result.Value, node, mBudget)
        If result.State <> atEvalOK Then
            Set EvaluateNode = SkipRemaining(node, i + 1, result)
            Exit Function
        End If
        args.Add result.Value
    Next i
    Set EvaluateNode = AT_ApplyScalar(node, args, mContext, mBudget)
End Function

Private Function SkipRemaining(ByVal node As ATAstNode, ByVal first As Long, ByVal primary As ATEvalResult) As ATEvalResult
    Dim i As Long, code As String
    On Error GoTo Failed
    For i = first To node.Children.Count
        LogNode node.Children(i), "NOT_EVALUATED"
    Next i
    Set SkipRemaining = primary
    Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "RUNTIME_INTERNAL_ERROR")
    Set SkipRemaining = WithSecondary(primary, code, node)
End Function

