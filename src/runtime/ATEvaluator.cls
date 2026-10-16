VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATEvaluator"
Option Explicit
Private mContext As ATEvaluationContext
Private mBudget As ATEvalBudget
Private mRows As Collection
Private mActiveRow As ATRuntimeRow

Public Function Evaluate(ByVal compiled As ATCompileResult, ByVal context As ATEvaluationContext, Optional ByVal limits As ATEvalLimits = Nothing) As ATEvalResult
    Dim result As ATEvalResult, final As New ATEvalResult, frame As New Collection
    Dim code As String, value As ATRuntimeValue
    Set mBudget = Nothing
    Set mContext = Nothing
    Set mRows = New Collection
    Set mActiveRow = Nothing
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
    If result.State = atEvalOK And Len(node.ConversionNormalizerId) > 0 Then
        If node.ConversionNormalizerId = "PORT_SET_V1" Then
            Set result = AT_ConvertPortValue(result.Value, node, mBudget)
        Else
            Set result = AT_RuntimeFailure(atEvalError, "RUNTIME_UNSUPPORTED_NORMALIZER", node)
        End If
    End If
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

Private Sub LogNode(ByVal node As ATAstNode, ByVal status As String, Optional ByVal result As ATEvalResult = Nothing, Optional ByVal rowContext As ATRuntimeRow = Nothing)
    Dim entry As New ATTraceEntry, valueState As String, reason As String
    Dim rowId As String, tableId As String, cellRef As String
    If Not result Is Nothing Then
        If result.State = atEvalOK Then valueState = CStr(result.Value.ValueState)
        If result.Diagnostics.Count > 0 Then reason = result.Diagnostics(1).Code
    End If
    If rowContext Is Nothing Then
        If Len(node.RowAlias) > 0 Then
            Set rowContext = FindRow(node.RowAlias)
        ElseIf node.Kind <> "FIELD" Then
            Set rowContext = mActiveRow
        End If
    End If
    If Not rowContext Is Nothing Then
        rowId = rowContext.RowId: tableId = rowContext.TableId
        If Len(node.RowAlias) > 0 Then cellRef = rowContext.CellRef(node.FieldId)
        mBudget.CheckText Len(rowId): mBudget.CheckText Len(tableId): mBudget.CheckText Len(cellRef)
    End If
    entry.Initialize node.SourceId, node.Start, node.Length, node.PrimitiveId, status, valueState, reason, mBudget.TraceCount + 1, rowId, tableId, node.FieldId, cellRef
    mBudget.AddTrace entry
End Sub

Private Function EvaluateNode(ByVal node As ATAstNode, ByVal frame As Collection) As ATEvalResult
    Dim result As ATEvalResult, right As ATEvalResult, value As ATRuntimeValue, items As New Collection
    Dim child As ATAstNode, args As New Collection
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
            Set result = ResolveBoundField(node)
            If result.State = atEvalError Then
                If result.Diagnostics.Count > 0 Then
                    If result.Diagnostics(1).Code = "RUNTIME_FIELD_MISSING" Or result.Diagnostics(1).Code = "RUNTIME_ROW_FIELD_MISSING" Then Set result = AT_RuntimeFailure(atEvalError, result.Diagnostics(1).Code, node)
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
            Set EvaluateNode = EvaluateNamed(node, frame)
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
    Select Case node.PrimitiveId
        Case "ALLROWS", "ANYROW", "NOROWS", "FILTER"
            Set EvaluateNode = EvaluateRows(node, frame)
            Exit Function
    End Select
    Set result = Visit(node.Children(1), frame)
    If result.State <> atEvalOK Then
        Set EvaluateNode = SkipRemaining(node, 2, result)
        Exit Function
    End If
    Set value = result.Value
    Select Case node.PrimitiveId
        Case "COUNT", "EXISTS"
            Set result = AT_Usable(value, node, mBudget)
            If result.State <> atEvalOK Then
                Set EvaluateNode = result
            ElseIf node.PrimitiveId = "COUNT" Then
                Set EvaluateNode = AT_RuntimeOK(AT_RuntimeScalar("NUMBER", CDbl(value.Count)))
            Else
                Set EvaluateNode = AT_RuntimeOK(AT_RuntimeScalar("BOOL", value.Count > 0))
            End If
            Exit Function
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

Private Function FindRow(ByVal aliasId As String) As ATRuntimeRow
    On Error Resume Next
    Set FindRow = mRows(aliasId)
    On Error GoTo 0
End Function

Private Function ResolveBoundField(ByVal node As ATAstNode) As ATEvalResult
    Dim row As ATRuntimeRow, id As String
    id = node.FieldId
    If Len(id) = 0 Then id = node.Text
    If Len(node.RowAlias) = 0 Then
        Set ResolveBoundField = mContext.ResolveField(id, mBudget)
    Else
        Set row = FindRow(node.RowAlias)
        If row Is Nothing Then
            Set ResolveBoundField = AT_RuntimeFailure(atEvalError, "RUNTIME_ROW_ALIAS_REQUIRED", node)
        ElseIf row.TableId <> node.TableId Then
            Set ResolveBoundField = AT_RuntimeFailure(atEvalError, "RUNTIME_ROW_TABLE_MISMATCH", node)
        Else
            Set ResolveBoundField = row.ResolveField(id, mBudget)
        End If
    End If
End Function

Private Function EvaluateNamed(ByVal node As ATAstNode, ByVal frame As Collection) As ATEvalResult
    Dim callerRows As Collection, definitionRows As Collection, callerActive As ATRuntimeRow, definitionActive As ATRuntimeRow
    Dim localFrame As New Collection, binding As ATArgumentBinding, result As ATEvalResult
    Dim code As Long, description As String
    Set callerRows = mRows: Set callerActive = mActiveRow
    On Error GoTo Failed
    Set definitionRows = New Collection
    If node.NamedScope = "ROW" Then
        Set definitionRows = callerRows: Set definitionActive = callerActive
    End If
    For Each binding In node.ParameterBindings
        If binding.IsDefault Then
            Set mRows = definitionRows: Set mActiveRow = definitionActive
            Set result = Visit(binding.Expression, localFrame)
        Else
            Set mRows = callerRows: Set mActiveRow = callerActive
            Set result = Visit(binding.Expression, frame)
        End If
        Set mRows = callerRows: Set mActiveRow = callerActive
        If result.State = atEvalOK And Len(binding.NormalizerId) > 0 Then
            If binding.NormalizerId = "PORT_SET_V1" Then
                Set result = AT_ConvertPortValue(result.Value, binding.Expression, mBudget)
            Else
                Set result = AT_RuntimeFailure(atEvalError, "RUNTIME_UNSUPPORTED_NORMALIZER", binding.Expression)
            End If
        End If
        If result.State <> atEvalOK Then
            Set EvaluateNamed = result
            Exit Function
        End If
        localFrame.Add result, binding.Id
    Next binding
    Set mRows = definitionRows: Set mActiveRow = definitionActive
    Set result = Visit(node.Expansion, localFrame)
    Set mRows = callerRows: Set mActiveRow = callerActive
    Set EvaluateNamed = result
    Exit Function
Failed:
    code = Err.Number: description = Err.Description
    Set mRows = callerRows: Set mActiveRow = callerActive
    Err.Raise code, "ATEvaluator", description
End Function

Private Function EvaluateRows(ByVal node As ATAstNode, ByVal frame As Collection) As ATEvalResult
    Dim binding As ATAstNode, result As ATEvalResult, collection As ATRuntimeValue, row As ATRuntimeRow
    Dim selected As New Collection, filtered As New ATRuntimeValue, i As Long, answer As Boolean, decisive As Boolean
    Dim previousActive As ATRuntimeRow, added As Boolean, code As Long, description As String
    Set previousActive = mActiveRow
    Set binding = node.Children(1)
    On Error GoTo Failed
    Set result = Visit(binding.Children(1), frame)
    If result.State <> atEvalOK Then
        Set EvaluateRows = SkipRemaining(node, 2, result)
        Exit Function
    End If
    Set result = AT_Usable(result.Value, node, mBudget)
    If result.State <> atEvalOK Then
        Set EvaluateRows = SkipRemaining(node, 2, result)
        Exit Function
    End If
    Set collection = result.Value
    mBudget.CheckList collection.Count
    If Not FindRow(binding.Text) Is Nothing Then Err.Raise vbObjectError + 2702, , "RUNTIME_ROW_ALIAS_SHADOWING"
    For i = 1 To collection.Count
        mBudget.Step
        Set row = collection.Row(i)
        mRows.Add row, binding.Text: added = True
        Set mActiveRow = row
        Set result = Visit(node.Children(2), frame)
        mRows.Remove binding.Text: added = False
        Set mActiveRow = previousActive
        If result.State = atEvalOK Then Set result = AT_Usable(result.Value, node.Children(2), mBudget)
        If result.State <> atEvalOK Then
            Set EvaluateRows = SkipRows(node.Children(2), collection, i + 1, result)
            Exit Function
        End If
        answer = CBool(result.Value.Scalar)
        If node.PrimitiveId = "FILTER" Then
            If answer Then
                mBudget.CheckList selected.Count + 1
                selected.Add row
            End If
        Else
            decisive = (node.PrimitiveId = "ALLROWS" And Not answer) Or (node.PrimitiveId <> "ALLROWS" And answer)
            If decisive Then
                answer = (node.PrimitiveId = "ANYROW")
                Set result = AT_RuntimeOK(AT_RuntimeScalar("BOOL", answer))
                Set EvaluateRows = SkipRows(node.Children(2), collection, i + 1, result)
                Exit Function
            End If
        End If
    Next i
    If node.PrimitiveId = "FILTER" Then
        filtered.InitializeRows collection.RowTableId, selected
        Set EvaluateRows = AT_RuntimeOK(filtered)
    Else
        Set EvaluateRows = AT_RuntimeOK(AT_RuntimeScalar("BOOL", node.PrimitiveId <> "ANYROW"))
    End If
    Exit Function
Failed:
    code = Err.Number: description = Err.Description
    If added Then mRows.Remove binding.Text
    Set mActiveRow = previousActive
    Err.Raise code, "ATEvaluator", description
End Function

Private Function SkipRows(ByVal predicate As ATAstNode, ByVal rows As ATRuntimeValue, ByVal first As Long, ByVal primary As ATEvalResult) As ATEvalResult
    Dim i As Long, code As String
    On Error GoTo Failed
    For i = first To rows.Count
        mBudget.Step
        LogNode predicate, "NOT_EVALUATED", , rows.Row(i)
    Next i
    Set SkipRows = primary
    Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "RUNTIME_INTERNAL_ERROR")
    If primary.State = atEvalOK Then
        Set SkipRows = AT_RuntimeFailure(atEvalError, code, predicate)
    Else
        Set SkipRows = WithSecondary(primary, code, predicate)
    End If
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

