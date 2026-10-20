VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtEvaluator"
Option Explicit

Private mNodes As Object, mContext As CAtEvalContext, mRegistry As CAtOperationRegistry
Private mTrace As Collection, mSeenDiagnostics As Object
Private mObjectId As String, mLimitNode As String, mLimitReason As String
Private mSteps As Long, mMaxSteps As Long, mMaxDepth As Long, mMaxDiagnostics As Long, mMaxTrace As Long, mMaxText As Long
Private mHalted As Boolean

Public Function Evaluate(ByVal package As CAtPackage, ByVal expressionId As String, ByVal context As CAtEvalContext, Optional ByVal maxSteps As Long = 1024, Optional ByVal maxDepth As Long = 64, Optional ByVal maxDiagnostics As Long = 128, Optional ByVal maxTrace As Long = 2048, Optional ByVal maxText As Long = 65536) As CAtEvalResult
    Dim d As CAtDefinition, expression As CAtDefinition, n As CAtNode, r As CAtEvalResult, env As Object, issues As Collection, issue As Variant
    On Error GoTo Broken
    mObjectId = expressionId: mSteps = 0: mHalted = False: mLimitNode = "": mLimitReason = ""
    Set mTrace = New Collection: Set mSeenDiagnostics = AtDictionary(): Set mRegistry = New CAtOperationRegistry
    mMaxSteps = maxSteps: mMaxDepth = maxDepth: mMaxDiagnostics = maxDiagnostics: mMaxTrace = maxTrace: mMaxText = maxText
    If maxSteps < 1 Or maxSteps > 1024 Or maxDepth < 1 Or maxDepth > 64 Or maxDiagnostics < 1 Or maxDiagnostics > 128 Or maxTrace < 1 Or maxTrace > 2048 Or maxText < 1 Or maxText > 65536 Then
        Set r = Failure("Text", "ERROR", "", "E_BUDGET", "Invalid host budget; limits cannot exceed hard caps."): GoTo Done
    End If
    If package Is Nothing Or context Is Nothing Then GoTo Broken
    For Each d In package.Definitions
        If d.Kind = "Expression" And d.Id = expressionId Then Set expression = d
    Next d
    If expression Is Nothing Then Set r = Failure("Text", "ERROR", "", "E_EXPRESSION", "Expression does not exist."): GoTo Done
    If package.SchemaVersion <> "2.0.0" And package.SchemaVersion <> "3.0.0" And package.SchemaVersion <> "4.0.0" Then Set r = Failure(AtValue(expression, "ReturnType"), "ERROR", expression.RootNodeId, "E_SCHEMA", "Only validated AST schemas may execute."): GoTo Done
    If context.PackageId <> package.Id Or context.PackageVersion <> package.Version Or context.TemplateId <> expression.ScopeTemplateId Then
        Set r = Failure(AtValue(expression, "ReturnType"), "ERROR", expression.RootNodeId, "E_CONTEXT", "Snapshot package/version/template mismatch."): GoTo Done
    End If
    Set issues = AtValidate(package, mRegistry.Catalog)
    If AtHasErrors(issues) Then
        Set r = Failure(AtValue(expression, "ReturnType"), "ERROR", expression.RootNodeId, "E_SCHEMA", "Static validation failed; no AST executed.")
        For Each issue In issues
            If issue("Severity") = "Error" Then
                If r.Diagnostics.Count >= maxDiagnostics Then Halt expression.RootNodeId, "Diagnostic budget exhausted during validation.": Exit For
                r.AddDiagnostic CStr(issue("Code")), CStr(issue("ObjectId")), CStr(issue("NodeId")), CStr(issue("Message"))
            End If
        Next issue
        GoTo Done
    End If
    Set mContext = context: Set mNodes = AtDictionary(): Set env = AtDictionary()
    For Each n In expression.Nodes: mNodes.Add n.Id, n: Next n
    Set n = mNodes(expression.RootNodeId)
    Set r = EvalNode(n, env, 1)
Done:
    If mHalted Then
        r.SetData r.ResultType, "LIMIT", ""
        r.AddDiagnostic "E_LIMIT", mObjectId, mLimitNode, mLimitReason
        r.AddRequirement mObjectId, mLimitNode, "Budget", "LIMIT"
    End If
    If mMaxDiagnostics >= 1 Then r.BoundDiagnostics mMaxDiagnostics, mObjectId, mLimitNode
    r.SetTrace mTrace
    Set Evaluate = r: Exit Function
Broken:
    Set r = Failure("Text", "ERROR", "", "E_INTERNAL", "Invalid context/model or internal computation failure; no success inferred.")
    Resume Done
End Function

Private Function EvalNode(ByVal n As CAtNode, ByVal env As Object, ByVal depth As Long) As CAtEvalResult
    Dim r As CAtEvalResult, child As CAtEvalResult, effects As CAtEvalResult, source As CAtEvalResult
    Dim i As Long, chosen As Long, scopeEnv As Object, args As New Collection, firstState As String, id As Variant
    If mHalted Or depth > mMaxDepth Then
        If Not mHalted Then Halt n.Id, "Depth budget exhausted."
        SkipTree n, "LIMIT": Set EvalNode = Failure(n.ResultType, "LIMIT", n.Id, "E_LIMIT", mLimitReason): Exit Function
    End If
    If Not Spend(n.Id) Then SkipTree n, "LIMIT": Set EvalNode = Failure(n.ResultType, "LIMIT", n.Id, "E_LIMIT", mLimitReason): Exit Function
    AddTrace n, "EXECUTED", "Reached in ordered evaluation"
    If mHalted Then Set EvalNode = Failure(n.ResultType, "LIMIT", n.Id, "E_LIMIT", mLimitReason): Exit Function
    Set effects = AtEvalValue("Text", "OK", "")
    Select Case n.Kind
        Case "TypedLiteral"
            Set r = AtEvalValue(n.ResultType, n.Values("State"), n.Values("Value"))
        Case "FieldRef"
            Set source = mContext.ReadField(CStr(n.Values("FieldId")))
            If source Is Nothing Then
                Set r = Failure(n.ResultType, "MISSING", n.Id, "E_FIELD_MISSING", "Field is absent from the immutable snapshot.", "Field/" & n.Values("FieldId"))
            ElseIf source.ResultType <> n.ResultType Then
                Set r = Failure(n.ResultType, "ERROR", n.Id, "E_FIELD_TYPE", "Snapshot type differs from the declared field type.", "Field/" & n.Values("FieldId"))
                r.MergeEffects source
            Else
                Set r = source.Snapshot()
            End If
            r.AddProvenance "Field", CStr(n.Values("FieldId"))
        Case "ValueRef": Set r = env(CStr(n.Values("BindingId"))).Snapshot()
        Case "Let"
            Set scopeEnv = AtDictionary()
            For Each id In env.Keys: scopeEnv.Add CStr(id), env(id): Next id
            For i = 1 To n.Children.Count
                If mHalted Then
                    SkipTree mNodes(CStr(n.Children(i))), "LIMIT"
                Else
                    Set child = EvalNode(mNodes(CStr(n.Children(i))), scopeEnv, depth + 1)
                    effects.MergeEffects child
                    If i <= n.Bindings.Count Then scopeEnv.Add CStr(n.Bindings(i)), child.Snapshot()
                    Set r = child.Snapshot()
                End If
            Next i
            r.MergeEffects effects
        Case "StateIs"
            Set child = EvalNode(mNodes(CStr(n.Children(1))), env, depth + 1)
            Set r = AtEvalValue("Bool", "OK", IIf(InStr(1, "|" & n.Values("States") & "|", "|" & child.State & "|", vbBinaryCompare) > 0, "true", "false"))
            r.MergeEffects child
        Case "Not"
            Set child = EvalNode(mNodes(CStr(n.Children(1))), env, depth + 1)
            Set r = Reframe(child, n.ResultType)
            If child.State = "OK" Then r.SetData "Bool", "OK", IIf(child.Value = "true", "false", "true")
        Case "When"
            Set child = EvalNode(mNodes(CStr(n.Children(1))), env, depth + 1)
            If child.State <> "OK" Or mHalted Then
                Set r = Reframe(child, n.ResultType)
                SkipTree mNodes(CStr(n.Children(2))), "Non-OK gate: " & child.State
                SkipTree mNodes(CStr(n.Children(3))), "Non-OK gate: " & child.State
            Else
                chosen = IIf(child.Value = "true", 2, 3)
                SkipTree mNodes(CStr(n.Children(IIf(chosen = 2, 3, 2)))), "When branch not selected"
                Set r = EvalNode(mNodes(CStr(n.Children(chosen))), env, depth + 1)
                r.MergeEffects child
            End If
        Case "AllOf", "AnyOf"
            Set r = AtEvalValue("Bool", "OK", IIf(n.Kind = "AllOf", "true", "false"))
            For i = 1 To n.Children.Count
                Set child = EvalNode(mNodes(CStr(n.Children(i))), env, depth + 1)
                r.MergeEffects child
                If child.State <> "OK" Or mHalted Then
                    r.SetData "Bool", child.State, "": Exit For
                End If
                If (n.Kind = "AllOf" And child.Value = "false") Or (n.Kind = "AnyOf" And child.Value = "true") Then
                    r.SetData "Bool", "OK", child.Value: Exit For
                End If
            Next i
            For i = i + 1 To n.Children.Count: SkipTree mNodes(CStr(n.Children(i))), "Short-circuit " & n.Kind & ": " & r.State & " " & r.Value: Next i
        Case "Call"
            ' Calls eagerly visit every argument in order; only control nodes short-circuit.
            For i = 1 To n.Children.Count
                If mHalted Then
                    SkipTree mNodes(CStr(n.Children(i))), "LIMIT"
                Else
                    Set child = EvalNode(mNodes(CStr(n.Children(i))), env, depth + 1)
                    args.Add child
                    If child.State = "EMPTY" And mRegistry.HandlesEmpty(CStr(n.Values("OperationId")), CStr(n.Values("OperationVersion"))) Then
                        effects.MergeEffectsConsumingEmpty child
                    Else
                        effects.MergeEffects child
                        If firstState = "" And child.State <> "OK" Then firstState = child.State
                    End If
                End If
            Next i
            If mHalted Then
                Set r = Failure(n.ResultType, "LIMIT", n.Id, "E_LIMIT", mLimitReason)
            ElseIf firstState <> "" Then
                Set r = AtEvalValue(n.ResultType, firstState, "")
            Else
                For Each id In n.Parameters.Keys
                    If Len(CStr(n.Parameters(id)("Value"))) > mMaxText Then Halt n.Id, "Parameter payload budget exhausted."
                Next id
                If Not mHalted And Spend(n.Id) Then
                    AddTrace n, "OPERATION", "Exact registered implementation"
                    If Not mHalted Then Set r = mRegistry.ExecuteOperation(CStr(n.Values("OperationId")), CStr(n.Values("OperationVersion")), args, n.Parameters)
                End If
                If r Is Nothing Then Set r = Failure(n.ResultType, "LIMIT", n.Id, "E_LIMIT", mLimitReason)
            End If
            r.MergeEffects effects
        Case Else: Set r = Failure(n.ResultType, "ERROR", n.Id, "E_NODE", "Unsupported AST node.")
    End Select
    If Len(r.Value) > mMaxText Then
        Halt n.Id, "Scalar payload budget exhausted.": r.SetData n.ResultType, "LIMIT", ""
    End If
    ' Create obligations at observation/operation origins, never at propagation nodes.
    ' Explicit EMPTY constants remain constants through Let/ValueRef and strict Calls.
    If r.State <> "OK" And ((n.Kind = "FieldRef" And r.Requirements.Count = 0) Or (n.Kind = "Call" And firstState = "")) Then
        r.AddRequirement mObjectId, n.Id, IIf(r.Dependency = "", "AST/" & n.Id, r.Dependency), r.State
        r.AddDiagnostic "E_" & r.State, mObjectId, n.Id, "Non-OK typed input/result; no Boolean coercion."
    End If
    r.AddProvenance "Node", mObjectId & "/" & n.Id
    ObserveDiagnostics r, n.Id
    If mHalted Then
        r.SetData n.ResultType, "LIMIT", ""
        r.AddRequirement mObjectId, mLimitNode, "Budget", "LIMIT"
    End If
    Set EvalNode = r
End Function

Private Function Reframe(ByVal source As CAtEvalResult, ByVal typeName As String) As CAtEvalResult
    Dim r As CAtEvalResult
    Set r = AtEvalValue(typeName, source.State, IIf(source.State = "OK", source.Value, ""))
    r.MergeEffects source: Set Reframe = r
End Function

Private Function Failure(ByVal typeName As String, ByVal state As String, ByVal nodeId As String, ByVal code As String, ByVal message As String, Optional ByVal dependency As String = "Computation") As CAtEvalResult
    Dim r As New CAtEvalResult
    If Not AtAstType(typeName) Then typeName = "Text"
    r.SetData typeName, state, "", dependency
    r.AddDiagnostic code, mObjectId, nodeId, message
    r.AddRequirement mObjectId, nodeId, dependency, state
    Set Failure = r
End Function

Private Function Spend(ByVal nodeId As String) As Boolean
    If mSteps >= mMaxSteps Then Halt nodeId, "Evaluation step budget exhausted.": Exit Function
    mSteps = mSteps + 1: Spend = True
End Function

Private Sub Halt(ByVal nodeId As String, ByVal reason As String)
    If mHalted Then Exit Sub
    mHalted = True: mLimitNode = nodeId: mLimitReason = reason
End Sub

Private Sub AddTrace(ByVal n As CAtNode, ByVal eventName As String, ByVal reason As String)
    If mTrace.Count >= mMaxTrace Then Halt n.Id, "Trace budget exhausted; remaining trace omitted.": Exit Sub
    Dim item As Object: Set item = AtDictionary()
    item.Add "Sequence", CStr(mTrace.Count + 1): item.Add "ObjectId", mObjectId: item.Add "NodeId", n.Id
    item.Add "Kind", n.Kind: item.Add "Event", eventName: item.Add "Reason", reason
    mTrace.Add item
End Sub

Private Sub SkipTree(ByVal n As CAtNode, ByVal reason As String)
    Dim id As Variant
    AddTrace n, "NOT_EVALUATED", reason
    If mTrace.Count >= mMaxTrace Then
        If n.Children.Count > 0 Then Halt n.Id, "Trace budget exhausted; skipped subtree omitted."
        Exit Sub
    End If
    For Each id In n.Children: SkipTree mNodes(CStr(id)), reason: Next id
End Sub

Private Sub ObserveDiagnostics(ByVal r As CAtEvalResult, ByVal nodeId As String)
    Dim item As Variant
    For Each item In r.Diagnostics
        If Not mSeenDiagnostics.Exists(item("DiagnosticId")) Then
            If mSeenDiagnostics.Count >= mMaxDiagnostics Then Halt nodeId, "Diagnostic budget exhausted.": Exit Sub
            mSeenDiagnostics.Add item("DiagnosticId"), True
        End If
    Next item
End Sub
