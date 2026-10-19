Attribute VB_Name = "modAtAst"
Option Explicit

Private Const TYPES As String = "|Text|Bool|Int64|Decimal|Date|DateTime|Email|Hostname|IpSet|PortSet|"
Private Const STATES As String = "|OK|EMPTY|MISSING|INVALID|UNAVAILABLE|ERROR|CONFLICT|LIMIT|NOT_APPLICABLE|"

Public Function AtAstType(ByVal text As String) As Boolean
    AtAstType = Member(text, TYPES)
End Function

Public Function AtAstState(ByVal text As String) As Boolean
    AtAstState = Member(text, STATES)
End Function

Private Function Member(ByVal value As String, ByVal choices As String) As Boolean
    If value <> "" And InStr(value, "|") = 0 Then Member = InStr(1, choices, "|" & value & "|", vbBinaryCompare) > 0
End Function

Public Sub AtAstProblem(ByVal issues As Collection, ByVal code As String, ByVal objectId As String, ByVal nodeId As String, ByVal reason As String)
    AtProblem issues, code, objectId, "AST", reason, "Correct this condition node: " & reason
    issues(issues.Count)("NodeId") = nodeId
End Sub

Public Sub AtValidateAst(ByVal p As CAtPackage, ByVal definitions As Object, ByVal catalog As Object, ByVal issues As Collection, Optional ByVal maxNodes As Long = 512, Optional ByVal maxDepth As Long = 64)
    Dim d As CAtDefinition, total As Long, r As Variant, expression As CAtDefinition
    On Error GoTo Broken
    If maxNodes < 1 Or maxNodes > 512 Or maxDepth < 1 Or maxDepth > 64 Then
        AtAstProblem issues, "AST_BUDGET", p.Id, "", "Host limits must be within 1..512 nodes and 1..64 depth.": Exit Sub
    End If
    For Each d In p.Definitions
        total = total + d.Nodes.Count
        If total > 4096 Then AtAstProblem issues, "AST_LIMIT", d.Id, d.RootNodeId, "Package exceeds 4096 nodes.": Exit Sub
        If d.Kind = "Expression" Then
            ValidateTree d, definitions, catalog, issues, maxNodes, maxDepth
        ElseIf d.Nodes.Count > 0 Or d.RootNodeId <> "" Or d.ScopeTemplateId <> "" Then
            AtAstProblem issues, "AST_OWNER", d.Id, d.RootNodeId, "Only Expression definitions may contain AST."
        End If
    Next d
    For Each d In p.Definitions
        If d.Kind = "Rule" Then
            For Each r In d.References
                If r("Role") = "Expression" And definitions.Exists("Expression/" & r("Id")) Then
                    Set expression = definitions("Expression/" & r("Id"))
                    If expression.ScopeTemplateId <> Owner(d) Then AtAstProblem issues, "AST_SCOPE", d.Id, expression.RootNodeId, "Rule and expression must have the same template context."
                End If
            Next r
        End If
    Next d
    Exit Sub
Broken:
    If d Is Nothing Then
        AtAstProblem issues, "AST_MODEL", p.Id, "", "Malformed AST metadata."
    Else
        AtAstProblem issues, "AST_MODEL", d.Id, d.RootNodeId, "Malformed AST metadata: " & Err.Number
    End If
End Sub

Private Sub ValidateTree(ByVal d As CAtDefinition, ByVal definitions As Object, ByVal catalog As Object, ByVal issues As Collection, ByVal maxNodes As Long, ByVal maxDepth As Long)
    Dim nodes As Object, parents As Object, colors As Object, bindings As Object
    Dim n As CAtNode, id As Variant, prop As Variant, key As Variant, startErrors As Long
    Dim child As CAtNode, env As Object
    If d.Nodes.Count = 0 Or d.Nodes.Count > maxNodes Then AtAstProblem issues, "AST_LIMIT", d.Id, d.RootNodeId, "Expression must contain 1.." & maxNodes & " nodes.": Exit Sub
    If Not definitions.Exists("Template/" & d.ScopeTemplateId) Then AtAstProblem issues, "AST_SCOPE", d.Id, d.RootNodeId, "Template context does not exist."
    Set nodes = AtDictionary(): Set parents = AtDictionary(): Set colors = AtDictionary(): Set bindings = AtDictionary()
    startErrors = issues.Count
    For Each n In d.Nodes
        If Not AtValidId(n.Id) Then AtAstProblem issues, "AST_ID", d.Id, n.Id, "Invalid stable NodeId."
        If nodes.Exists(n.Id) Then
            AtAstProblem issues, "AST_DUPLICATE", d.Id, n.Id, "Duplicate NodeId."
        Else
            nodes.Add n.Id, n: parents.Add n.Id, 0
        End If
        If Not AtAstType(n.ResultType) Then AtAstProblem issues, "AST_TYPE", d.Id, n.Id, "Unknown declared result type."
        For Each prop In n.Values.Keys
            If Not Member(CStr(prop), NodeProperties(n.Kind)) Or VarType(n.Values(prop)) <> vbString Then AtAstProblem issues, "AST_PROPERTY", d.Id, n.Id, "Unknown or non-string node property: " & prop
        Next prop
        If n.Kind <> "Call" And n.Parameters.Count > 0 Then AtAstProblem issues, "AST_PARAMETER", d.Id, n.Id, "Only Call accepts operation parameters."
        If n.Kind <> "Let" And n.Bindings.Count > 0 Then AtAstProblem issues, "AST_BINDING", d.Id, n.Id, "Only Let declares lexical bindings."
        For Each id In n.Bindings
            If Not AtValidId(CStr(id)) Or bindings.Exists(CStr(id)) Then AtAstProblem issues, "AST_BINDING", d.Id, n.Id, "Invalid or duplicate BindingId."
            bindings(CStr(id)) = True
        Next id
    Next n
    If Not nodes.Exists(d.RootNodeId) Then AtAstProblem issues, "AST_ROOT", d.Id, d.RootNodeId, "Root NodeId does not exist.": Exit Sub
    For Each n In d.Nodes
        For Each id In n.Children
            If Not nodes.Exists(CStr(id)) Then
                AtAstProblem issues, "AST_REFERENCE", d.Id, n.Id, "Child NodeId does not exist: " & CStr(id)
            Else
                parents(CStr(id)) = parents(CStr(id)) + 1
            End If
        Next id
    Next n
    For Each key In parents.Keys
        If parents(key) <> IIf(key = d.RootNodeId, 0, 1) Then AtAstProblem issues, "AST_TREE", d.Id, CStr(key), "Root needs zero parents; each other node exactly one."
    Next key
    For Each n In d.Nodes
        GraphVisit d.Id, n, nodes, colors, issues, 1, maxDepth
    Next n
    If issues.Count > startErrors Then Exit Sub
    Set env = AtDictionary(): Set colors = AtDictionary()
    Set n = nodes(d.RootNodeId)
    Walk d, n, nodes, definitions, catalog, env, colors, issues, 1, maxDepth
    For Each key In nodes.Keys
        If Not colors.Exists(key) Then AtAstProblem issues, "AST_UNREACHABLE", d.Id, CStr(key), "Node is not reachable from the root."
    Next key
    If n.ResultType <> AtValue(d, "ReturnType") Then AtAstProblem issues, "AST_TYPE", d.Id, n.Id, "Root type differs from Expression ReturnType."
End Sub

Private Sub GraphVisit(ByVal objectId As String, ByVal n As CAtNode, ByVal nodes As Object, ByVal colors As Object, ByVal issues As Collection, ByVal depth As Long, ByVal maxDepth As Long)
    Dim id As Variant, child As CAtNode
    If colors.Exists(n.Id) Then
        If colors(n.Id) = 1 Then AtAstProblem issues, "AST_CYCLE", objectId, n.Id, "Circular child reference."
        Exit Sub
    End If
    If depth > maxDepth Then AtAstProblem issues, "AST_LIMIT", objectId, n.Id, "Tree depth exceeds host budget.": Exit Sub
    colors.Add n.Id, 1
    For Each id In n.Children
        If nodes.Exists(CStr(id)) Then
            Set child = nodes(CStr(id)): GraphVisit objectId, child, nodes, colors, issues, depth + 1, maxDepth
        End If
    Next id
    colors(n.Id) = 2
End Sub

Private Sub Walk(ByVal d As CAtDefinition, ByVal n As CAtNode, ByVal nodes As Object, ByVal definitions As Object, ByVal catalog As Object, ByVal env As Object, ByVal visited As Object, ByVal issues As Collection, ByVal depth As Long, ByVal maxDepth As Long)
    Dim scopeEnv As Object, id As Variant, child As CAtNode, i As Long, expected As String, a As CAtNode, b As CAtNode
    If depth > maxDepth Then AtAstProblem issues, "AST_LIMIT", d.Id, n.Id, "Tree depth exceeds host budget.": Exit Sub
    If visited.Exists(n.Id) Then Exit Sub
    visited.Add n.Id, True
    Set scopeEnv = CopyEnv(env)
    For i = 1 To n.Children.Count
        Set child = nodes(CStr(n.Children(i)))
        Walk d, child, nodes, definitions, catalog, scopeEnv, visited, issues, depth + 1, maxDepth
        If n.Kind = "Let" And i <= n.Bindings.Count Then scopeEnv(CStr(n.Bindings(i))) = child.ResultType
    Next i
    Select Case n.Kind
        Case "TypedLiteral"
            Arity d, n, 0, 0, issues
            If Not n.Values.Exists("Value") Or Not Member(Value(n, "State"), "|OK|EMPTY|") Then
                AtAstProblem issues, "AST_LITERAL", d.Id, n.Id, "Literal requires Value and State OK or EMPTY."
            ElseIf Value(n, "State") = "EMPTY" Then
                If Value(n, "Value") <> "" Then AtAstProblem issues, "AST_LITERAL", d.Id, n.Id, "EMPTY literal has no payload."
            ElseIf Not AtLiteralValid(n.ResultType, Value(n, "Value")) Then
                AtAstProblem issues, "AST_LITERAL", d.Id, n.Id, "Invalid canonical payload for declared literal type."
            End If
        Case "FieldRef"
            Arity d, n, 0, 0, issues
            Dim field As CAtDefinition
            If Not AtValidId(Value(n, "FieldId")) Or Not definitions.Exists("Field/" & Value(n, "FieldId")) Then
                AtAstProblem issues, "AST_FIELD", d.Id, n.Id, "Referenced field does not exist."
            Else
                Set field = definitions("Field/" & Value(n, "FieldId"))
                If field.Version <> Value(n, "FieldVersion") Then AtAstProblem issues, "AST_VERSION", d.Id, n.Id, "Field exact version mismatch."
                If Owner(field) <> d.ScopeTemplateId Then AtAstProblem issues, "AST_SCOPE", d.Id, n.Id, "Field belongs to another template."
                MatchType d, n, AtValue(field, "DataType"), issues
            End If
        Case "ValueRef"
            Arity d, n, 0, 0, issues
            If Not env.Exists(Value(n, "BindingId")) Then
                AtAstProblem issues, "AST_SCOPE", d.Id, n.Id, "Binding is absent, forward, self-referencing or outside lexical scope."
            Else
                MatchType d, n, CStr(env(Value(n, "BindingId"))), issues
            End If
        Case "Call": CheckCall d, n, nodes, catalog, issues
        Case "AllOf", "AnyOf", "Not"
            If n.Kind = "Not" Then Arity d, n, 1, 1, issues Else Arity d, n, 1, 512, issues
            MatchType d, n, "Bool", issues
            For Each id In n.Children
                Set child = nodes(CStr(id)): MatchType d, child, "Bool", issues
            Next id
        Case "When"
            Arity d, n, 3, 3, issues
            If n.Children.Count = 3 Then
                Set child = nodes(CStr(n.Children(1))): MatchType d, child, "Bool", issues
                Set a = nodes(CStr(n.Children(2))): Set b = nodes(CStr(n.Children(3)))
                MatchType d, a, n.ResultType, issues: MatchType d, b, n.ResultType, issues
            End If
        Case "StateIs"
            Arity d, n, 1, 1, issues: MatchType d, n, "Bool", issues
            Dim seen As Object: Set seen = AtDictionary()
            For Each id In Split(Value(n, "States"), "|")
                If Not AtAstState(CStr(id)) Or seen.Exists(CStr(id)) Then AtAstProblem issues, "AST_STATE", d.Id, n.Id, "Unknown or repeated EvalState."
                seen(CStr(id)) = True
            Next id
        Case "Let"
            If n.Bindings.Count = 0 Then AtAstProblem issues, "AST_BINDING", d.Id, n.Id, "Let requires at least one binding."
            Arity d, n, n.Bindings.Count + 1, n.Bindings.Count + 1, issues
            If n.Children.Count > 0 Then
                Set child = nodes(CStr(n.Children(n.Children.Count))): MatchType d, n, child.ResultType, issues
            End If
        Case Else: AtAstProblem issues, "AST_KIND", d.Id, n.Id, "Unsupported NodeKind."
    End Select
End Sub

Private Sub CheckCall(ByVal d As CAtDefinition, ByVal n As CAtNode, ByVal nodes As Object, ByVal catalog As Object, ByVal issues As Collection)
    Dim key As String, s As Object, types As Collection, parameters As Object, states As Collection
    Dim i As Long, child As CAtNode, prop As Variant, parameter As Object, seen As Object
    On Error GoTo BadSignature
    key = Value(n, "OperationId") & "@" & Value(n, "OperationVersion")
    If Not AtValidId(Value(n, "OperationId")) Or Not AtValidVersion(Value(n, "OperationVersion")) Then AtAstProblem issues, "AST_VERSION", d.Id, n.Id, "Operation needs a valid ID and exact version.": Exit Sub
    If catalog Is Nothing Then GoTo UnknownOperation
    If Not catalog.Exists(key) Then GoTo UnknownOperation
    If Not IsObject(catalog(key)) Then GoTo BadSignature
    Set s = catalog(key)
    If s.Count <> 5 Then GoTo BadSignature
    If Not AtAstType(CStr(s("ReturnType"))) Or s("Propagation") <> "Strict" Then GoTo BadSignature
    Set types = s("Arguments"): Set parameters = s("Parameters"): Set states = s("States")
    If types.Count > 512 Or states.Count < 1 Or states.Count > 9 Then GoTo BadSignature
    For Each prop In types
        If Not AtAstType(CStr(prop)) Then GoTo BadSignature
    Next prop
    Set seen = AtDictionary()
    For Each prop In states
        If Not AtAstState(CStr(prop)) Or seen.Exists(CStr(prop)) Then GoTo BadSignature
        seen.Add CStr(prop), True
    Next prop
    If Not seen.Exists("OK") Then GoTo BadSignature
    For Each prop In parameters.Keys
        If Not AtValidId(CStr(prop)) Or Not AtAstType(CStr(parameters(prop))) Then GoTo BadSignature
    Next prop
    MatchType d, n, CStr(s("ReturnType")), issues
    Arity d, n, types.Count, types.Count, issues
    If n.Children.Count = types.Count Then
        For i = 1 To types.Count
            Set child = nodes(CStr(n.Children(i))): MatchType d, child, CStr(types(i)), issues
        Next i
    End If
    For Each prop In parameters.Keys
        If Not n.Parameters.Exists(prop) Then AtAstProblem issues, "AST_PARAMETER", d.Id, n.Id, "Missing operation parameter: " & prop
    Next prop
    For Each prop In n.Parameters.Keys
        If Not parameters.Exists(prop) Then
            AtAstProblem issues, "AST_PARAMETER", d.Id, n.Id, "Unknown operation parameter: " & prop
        Else
            Set parameter = n.Parameters(prop)
            If parameter.Count <> 2 Or parameter("Type") <> parameters(prop) Then
                AtAstProblem issues, "AST_PARAMETER", d.Id, n.Id, "Parameter type mismatch: " & prop
            ElseIf VarType(parameter("Value")) <> vbString Or Not AtLiteralValid(CStr(parameter("Type")), CStr(parameter("Value"))) Then
                AtAstProblem issues, "AST_PARAMETER", d.Id, n.Id, "Invalid canonical parameter payload: " & prop
            End If
        End If
    Next prop
    Exit Sub
UnknownOperation:
    AtAstProblem issues, "AST_OPERATION", d.Id, n.Id, "Exact operation version is not registered.": Exit Sub
BadSignature:
    AtAstProblem issues, "AST_SIGNATURE", d.Id, n.Id, "Trusted catalog entry or parameter record is malformed; no signature inferred."
End Sub

Private Sub Arity(ByVal d As CAtDefinition, ByVal n As CAtNode, ByVal low As Long, ByVal high As Long, ByVal issues As Collection)
    If n.Children.Count < low Or n.Children.Count > high Then AtAstProblem issues, "AST_ARITY", d.Id, n.Id, "Invalid number of ordered children."
End Sub

Private Sub MatchType(ByVal d As CAtDefinition, ByVal n As CAtNode, ByVal expected As String, ByVal issues As Collection)
    If n.ResultType <> expected Then AtAstProblem issues, "AST_TYPE", d.Id, n.Id, "Expected " & expected & ", declared " & n.ResultType & "; implicit conversion is forbidden."
End Sub

Private Function NodeProperties(ByVal kind As String) As String
    Select Case kind
        Case "TypedLiteral": NodeProperties = "|State|Value|"
        Case "FieldRef": NodeProperties = "|FieldId|FieldVersion|"
        Case "Call": NodeProperties = "|OperationId|OperationVersion|"
        Case "ValueRef": NodeProperties = "|BindingId|"
        Case "StateIs": NodeProperties = "|States|"
        Case Else: NodeProperties = "|"
    End Select
End Function

Private Function Value(ByVal n As CAtNode, ByVal key As String) As String
    If n.Values.Exists(key) Then Value = CStr(n.Values(key))
End Function

Private Function Owner(ByVal d As CAtDefinition) As String
    Dim r As Variant
    For Each r In d.References
        If r("Role") = "OwnerTemplate" Then Owner = r("Id"): Exit Function
    Next r
End Function

Private Function CopyEnv(ByVal env As Object) As Object
    Dim key As Variant, result As Object: Set result = AtDictionary()
    For Each key In env.Keys: result.Add key, env(key): Next key
    Set CopyEnv = result
End Function

Public Function AtLiteralValid(ByVal typeName As String, ByVal text As String) As Boolean
    On Error GoTo Invalid
    Dim unsigned As String, negative As Boolean, i As Long, c As String, parts As Variant, dateValue As Date, significant As String
    Select Case typeName
        Case "Text": AtLiteralValid = True
        Case "Bool": AtLiteralValid = text = "true" Or text = "false"
        Case "Int64", "Decimal"
            unsigned = text: negative = Left$(text, 1) = "-"
            If negative Then unsigned = Mid$(text, 2)
            parts = Split(unsigned, ".")
            If UBound(parts) > IIf(typeName = "Int64", 0, 1) Then Exit Function
            If parts(0) = "" Or (Len(parts(0)) > 1 And Left$(parts(0), 1) = "0") Then Exit Function
            If negative And Replace(unsigned, ".", "") = String$(Len(Replace(unsigned, ".", "")), "0") Then Exit Function
            For i = 1 To Len(unsigned)
                c = Mid$(unsigned, i, 1)
                If (c < "0" Or c > "9") And c <> "." Then Exit Function
            Next i
            If UBound(parts) = 1 Then
                If parts(1) = "" Or Right$(parts(1), 1) = "0" Then Exit Function
            End If
            If typeName = "Int64" Then
                If Len(unsigned) > 19 Then Exit Function
                If Len(unsigned) = 19 Then
                    If unsigned > IIf(negative, "9223372036854775808", "9223372036854775807") Then Exit Function
                End If
            Else
                If UBound(parts) = 1 Then
                    If Len(parts(1)) > 28 Then Exit Function
                End If
                significant = Replace(unsigned, ".", "")
                Do While Len(significant) > 1 And Left$(significant, 1) = "0"
                    significant = Mid$(significant, 2)
                Loop
                If Len(significant) > 28 Then Exit Function
            End If
            AtLiteralValid = True
        Case "Date", "DateTime"
            If typeName = "Date" Then
                If Not text Like "####-##-##" Then Exit Function
            Else
                If Not text Like "####-##-##T##:##:##Z" Then Exit Function
            End If
            If CLng(Left$(text, 4)) < 100 Then Exit Function
            dateValue = DateSerial(CInt(Left$(text, 4)), CInt(Mid$(text, 6, 2)), CInt(Mid$(text, 9, 2)))
            If Format$(dateValue, "yyyy-mm-dd") <> Left$(text, 10) Then Exit Function
            If typeName = "DateTime" Then
                dateValue = TimeSerial(CInt(Mid$(text, 12, 2)), CInt(Mid$(text, 15, 2)), CInt(Mid$(text, 18, 2)))
                If Format$(dateValue, "hh:nn:ss") <> Mid$(text, 12, 8) Then Exit Function
            End If
            AtLiteralValid = True
    End Select
Invalid:
End Function
