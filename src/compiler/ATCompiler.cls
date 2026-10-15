VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATCompiler"
Option Explicit

Private mContext As ATCompileContext
Private mResult As ATCompileResult
Private mResolver As ATResolver
Private mTypes As ATTypeChecker
Private mNamedStack As Collection
Private mCurrent As ATAstNode
Private mLimits As ATLimits
Private mCollectionDepth As Long
Private mSteps As Long
Private mCheckDepth As Long

Public Function Compile(ByVal text As String, ByVal context As ATCompileContext, Optional ByVal contextualFieldId As String = "", Optional ByVal requiredType As String = "BOOL", Optional ByVal limits As ATLimits = Nothing) As ATCompileResult
    Dim parser As New ATParser, parsed As ATParseResult
    Dim aliases As New Collection, parameters As New Collection
    Dim diagnostic As ATDiagnostic, typeName As String
    Set mContext = context
    Set mResult = New ATCompileResult
    Set mResolver = New ATResolver
    Set mTypes = New ATTypeChecker
    Set mNamedStack = New Collection
    Set mCurrent = Nothing
    On Error GoTo Failed
    If limits Is Nothing Then Set limits = New ATLimits
    Set mLimits = limits
    mLimits.Validate
    mCollectionDepth = 0
    mSteps = 0
    mCheckDepth = 0
    context.ValidateSchema
    Set parsed = parser.Parse(text, contextualFieldId, context.Registry, mLimits)
    For Each diagnostic In parsed.Diagnostics
        mResult.Diagnostics.Add diagnostic
    Next diagnostic
    If parsed.Diagnostics.Count > 0 Then
        Set Compile = mResult
        Exit Function
    End If
    typeName = CheckNode(parsed.Root, aliases, parameters)
    If typeName <> requiredType Then Err.Raise vbObjectError + 2330, "ATCompiler", "TYPE_RULE_ROOT: expected " & requiredType & ", got " & typeName
    Set mResult.Root = parsed.Root
    Set Compile = mResult
    Exit Function
Failed:
    Set diagnostic = New ATDiagnostic
    If mCurrent Is Nothing Then
        diagnostic.Initialize "COMPILE_CONTEXT", "SEMANTIC", "ERROR", 1, 0, Err.Description
    Else
        diagnostic.Initialize Split(Err.Description, ":")(0), "SEMANTIC", "ERROR", mCurrent.Start, mCurrent.Length, Err.Description
    End If
    mResult.Diagnostics.Add diagnostic
    Set mResult.Root = Nothing
    Set Compile = mResult
End Function

Private Function CheckNode(ByVal node As ATAstNode, ByVal aliases As Collection, ByVal parameters As Collection) As String
    Dim types As New Collection, child As ATAstNode
    Dim definition As ATFieldDefinition, primitive As ATPrimitive
    Dim typeName As String, item As Variant
    Set mCurrent = node
    mSteps = mSteps + 1
    mCheckDepth = mCheckDepth + 1
    If mSteps > mLimits.MaxCompileSteps Then Err.Raise vbObjectError + 2401, "ATCompiler", "LIMIT_COMPILE_STEPS"
    If mCheckDepth > mLimits.MaxNesting Then Err.Raise vbObjectError + 2402, "ATCompiler", "LIMIT_EXPANSION_DEPTH"
    Select Case node.Kind
        Case "NUMBER": typeName = "NUMBER"
        Case "STRING": typeName = "TEXT"
        Case "BOOL": typeName = "BOOL"
        Case "FIELD"
            Set definition = mResolver.ResolveField(node.Text, mContext, aliases)
            typeName = definition.TypeName
            mResult.Dependencies.Add "FIELD:" & definition.TableId & ":" & definition.Id, Owner()
        Case "LIST"
            If node.Children.Count = 0 Then Err.Raise vbObjectError + 2331, "ATCompiler", "TYPE_EMPTY_LIST: explicit element type required"
            For Each child In node.Children
                types.Add CheckNode(child, aliases, parameters)
            Next child
            typeName = CStr(types(1))
            For Each item In types
                If item <> typeName Then Err.Raise vbObjectError + 2332, "ATCompiler", "TYPE_MIXED_LIST: homogeneous list required"
            Next item
            typeName = "LIST<" & typeName & ">"
        Case "NAMED"
            typeName = CheckNamed(node, aliases, parameters)
        Case "CALL"
            If node.Text = "SOURCE" Then
                Set primitive = mContext.Registry.FindSyntax(node.Text, "FUNCTION")
                If node.Children.Count <> 1 Then Err.Raise vbObjectError + 2333, "ATCompiler", "TYPE_SOURCE_ARGUMENTS"
                Set child = node.Children(1)
                If child.Kind <> "STRING" Then Err.Raise vbObjectError + 2333, "ATCompiler", "TYPE_SOURCE_ID_LITERAL_REQUIRED"
                Set definition = mContext.FindSource(child.Text)
                If definition Is Nothing Then Err.Raise vbObjectError + 2334, "ATCompiler", "RESOLVE_UNKNOWN_SOURCE: " & child.Text
                typeName = definition.TypeName
                child.DataType = "TEXT"
                mResult.Dependencies.Add "SOURCE:" & definition.Id, Owner()
            ElseIf mContext.Registry.HasSyntax(node.Text, "QUANTIFIER") Then
                typeName = CheckCollection(node, aliases, parameters)
            Else
                Set primitive = mContext.Registry.FindSyntax(node.Text, "FUNCTION")
                For Each child In node.Children
                    types.Add CheckNode(child, aliases, parameters)
                Next child
                typeName = mTypes.ResolveSignature(primitive, types, mContext.Registry)
                If node.Text = "DATE" Then ValidateDate node
            End If
        Case "UNARY", "BINARY", "RELATION"
            For Each child In node.Children
                types.Add CheckNode(child, aliases, parameters)
            Next child
            If node.Kind = "UNARY" And (node.Text = "+" Or node.Text = "-") Then
                Set primitive = mContext.Registry.GetPrimitive("UNARY_" & IIf(node.Text = "+", "PLUS", "MINUS"))
            Else
                Set primitive = mContext.Registry.FindSyntax(node.Text, "OPERATOR")
            End If
            typeName = mTypes.ResolveSignature(primitive, types, mContext.Registry)
        Case Else
            Err.Raise vbObjectError + 2335, "ATCompiler", "TYPE_INVALID_AST_NODE: " & node.Kind
    End Select
    node.DataType = typeName
    mCheckDepth = mCheckDepth - 1
    CheckNode = typeName
End Function

Private Function Owner() As String
    If mNamedStack.Count = 0 Then
        Owner = "RULE"
    Else
        Owner = "NAMED:" & mNamedStack(mNamedStack.Count)
    End If
End Function

Private Function CheckCollection(ByVal node As ATAstNode, ByVal aliases As Collection, ByVal parameters As Collection) As String
    Dim binding As ATAstNode, collectionType As String, predicateType As String
    Dim nested As New Collection, entry As ATFieldDefinition, aliasValue As ATFieldDefinition
    Dim tableName As String, primitive As ATPrimitive
    Dim types As New Collection
    Set primitive = mContext.Registry.FindSyntax(node.Text, "QUANTIFIER")
    mCollectionDepth = mCollectionDepth + 1
    If mCollectionDepth > mLimits.MaxCollectionNesting Then Err.Raise vbObjectError + 2403, "ATCompiler", "LIMIT_COLLECTION_NESTING"
    If node.Children.Count <> 2 Then Err.Raise vbObjectError + 2340, "ATCompiler", "TYPE_QUANTIFIER_ARITY"
    Set binding = node.Children(1)
    If binding.Kind <> "BIND" Then Err.Raise vbObjectError + 2341, "ATCompiler", "RESOLVE_EXPLICIT_ALIAS_REQUIRED"
    collectionType = CheckNode(binding.Children(1), aliases, parameters)
    If Left$(collectionType, 15) <> "COLLECTION<ROW<" Or Right$(collectionType, 2) <> ">>" Then Err.Raise vbObjectError + 2342, "ATCompiler", "TYPE_ROW_COLLECTION_REQUIRED"
    tableName = Mid$(collectionType, 16, Len(collectionType) - 17)
    For Each entry In aliases
        If entry.Id = binding.Text Then Err.Raise vbObjectError + 2343, "ATCompiler", "RESOLVE_ALIAS_SHADOWING: unique aliases required"
        nested.Add entry
    Next entry
    Set aliasValue = New ATFieldDefinition
    aliasValue.Id = binding.Text
    aliasValue.TableId = tableName
    nested.Add aliasValue
    predicateType = CheckNode(node.Children(2), nested, parameters)
    If predicateType <> "BOOL" Then Err.Raise vbObjectError + 2344, "ATCompiler", "TYPE_BOOL_PREDICATE_REQUIRED"
    binding.DataType = collectionType
    types.Add collectionType
    types.Add predicateType
    CheckCollection = mTypes.ResolveSignature(primitive, types, mContext.Registry)
    mCollectionDepth = mCollectionDepth - 1
End Function

Private Function CheckNamed(ByVal node As ATAstNode, ByVal aliases As Collection, ByVal parameters As Collection) As String
    Dim definition As ATNamedExpression, parameter As ATFieldDefinition
    Dim nested As New Collection, child As ATAstNode, parser As New ATParser
    Dim parsed As ATParseResult, index As Long, actualType As String, present As Variant
    Dim supplied As New Collection
    Dim warning As ATDiagnostic
    For Each parameter In parameters
        If parameter.Id = node.Text Then
            If node.Children.Count <> 0 Then Err.Raise vbObjectError + 2350, "ATCompiler", "TYPE_PARAMETER_NOT_CALLABLE"
            CheckNamed = parameter.TypeName
            Exit Function
        End If
    Next parameter
    Set definition = mContext.FindNamed(node.Text)
    If definition Is Nothing Then Err.Raise vbObjectError + 2351, "ATCompiler", "RESOLVE_UNKNOWN_NAMED: " & node.Text
    If definition.Status = "DEPRECATED" Then
        Set warning = New ATDiagnostic
        warning.Initialize "NAMED_DEPRECATED", "SEMANTIC", "WARNING", node.Start, node.Length, "Deprecated named expression " & node.Text & "; replacement: " & definition.ReplacedBy
        mResult.Diagnostics.Add warning
    End If
    For Each present In mNamedStack
        If present = node.Text Then Err.Raise vbObjectError + 2352, "ATCompiler", "NAMED_CYCLE: " & node.Text
    Next present
    If node.Children.Count > definition.Parameters.Count Then Err.Raise vbObjectError + 2353, "ATCompiler", "TYPE_NAMED_ARITY"
    If mNamedStack.Count >= mLimits.MaxNamedDepth Then Err.Raise vbObjectError + 2404, "ATCompiler", "LIMIT_NAMED_DEPTH"
    mResult.Dependencies.Add "NAMED:" & node.Text, Owner()
    For index = 1 To node.Children.Count
        supplied.Add CheckNode(node.Children(index), aliases, parameters)
    Next index
    mNamedStack.Add node.Text
    For index = 1 To definition.Parameters.Count
        Set parameter = definition.Parameters(index)
        If index <= node.Children.Count Then
            actualType = supplied(index)
        Else
            If Len(parameter.DefaultExpression) = 0 Then Err.Raise vbObjectError + 2353, "ATCompiler", "TYPE_NAMED_REQUIRED_ARGUMENT"
            Set parsed = parser.Parse(parameter.DefaultExpression, "", mContext.Registry, mLimits)
            If parsed.Diagnostics.Count > 0 Then Err.Raise vbObjectError + 2354, "ATCompiler", "NAMED_INVALID_DEFAULT"
            actualType = CheckNode(parsed.Root, aliases, parameters)
        End If
        If actualType <> parameter.TypeName And Not mContext.Registry.CanConvert(actualType, parameter.TypeName) Then Err.Raise vbObjectError + 2355, "ATCompiler", "TYPE_NAMED_ARGUMENT"
        nested.Add parameter
    Next index
    Set parsed = parser.Parse(definition.Expression, "", mContext.Registry, mLimits)
    If parsed.Diagnostics.Count > 0 Then Err.Raise vbObjectError + 2356, "ATCompiler", "NAMED_INVALID_EXPRESSION: " & node.Text
    actualType = CheckNode(parsed.Root, aliases, nested)
    If actualType <> definition.ReturnType Then Err.Raise vbObjectError + 2357, "ATCompiler", "TYPE_NAMED_RETURN"
    mNamedStack.Remove mNamedStack.Count
    Set node.Expansion = parsed.Root
    CheckNamed = actualType
End Function

Private Sub ValidateDate(ByVal node As ATAstNode)
    Dim value As String, year As Long, month As Long, day As Long
    Dim parsed As Date
    If node.Children(1).Kind <> "STRING" Then Err.Raise vbObjectError + 2360, "ATCompiler", "DATE_LITERAL_REQUIRED"
    value = node.Children(1).Text
    If Len(value) <> 10 Or Mid$(value, 5, 1) <> "-" Or Mid$(value, 8, 1) <> "-" Then Err.Raise vbObjectError + 2361, "ATCompiler", "DATE_INVALID_LITERAL"
    If Not value Like "####-##-##" Then Err.Raise vbObjectError + 2361, "ATCompiler", "DATE_INVALID_LITERAL"
    year = CLng(Left$(value, 4))
    month = CLng(Mid$(value, 6, 2))
    day = CLng(Right$(value, 2))
    If year < 100 Or year > 9999 Then Err.Raise vbObjectError + 2361, "ATCompiler", "DATE_INVALID_LITERAL"
    parsed = DateSerial(year, month, day)
    If VBA.Year(parsed) <> year Or VBA.Month(parsed) <> month Or VBA.Day(parsed) <> day Then Err.Raise vbObjectError + 2361, "ATCompiler", "DATE_INVALID_LITERAL"
End Sub

