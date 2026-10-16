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
Private mSourceId As String

Public Function Compile(ByVal text As String, ByVal context As ATCompileContext, Optional ByVal contextualFieldId As String = "", Optional ByVal requiredType As String = "BOOL", Optional ByVal limits As ATLimits = Nothing) As ATCompileResult
    Dim parser As New ATParser, parsed As ATParseResult
    Dim aliases As New Collection, parameters As New Collection
    Dim diagnostic As ATDiagnostic, typeName As String
    Dim failureDescription As String
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
    mSourceId = "SCHEMA"
    context.ValidateSchema
    mSourceId = "RULE"
    Set parsed = parser.Parse(text, contextualFieldId, context.Registry, mLimits)
    For Each diagnostic In parsed.Diagnostics
        mResult.Diagnostics.Add diagnostic
    Next diagnostic
    If parsed.Diagnostics.Count > 0 Then
        Set Compile = mResult
        Exit Function
    End If
    typeName = CheckNode(parsed.Root, aliases, parameters)
    Set mCurrent = parsed.Root
    If typeName <> requiredType Then Err.Raise vbObjectError + 2330, "ATCompiler", "TYPE_RULE_ROOT: expected " & requiredType & ", got " & typeName
    Set mResult.Root = parsed.Root
    Set Compile = mResult
    Exit Function
Failed:
    failureDescription = Err.Description
    If Not mResult.HasErrors Then
        Set diagnostic = New ATDiagnostic
        If mCurrent Is Nothing Then
            diagnostic.Initialize AT_StableErrorCode(failureDescription, "COMPILE_CONTEXT"), "SEMANTIC", "ERROR", 1, 0, failureDescription
        Else
            diagnostic.Initialize AT_StableErrorCode(failureDescription, "COMPILE_INTERNAL"), "SEMANTIC", "ERROR", mCurrent.Start, mCurrent.Length, failureDescription
        End If
        diagnostic.SourceId = mSourceId
        mResult.Diagnostics.Add diagnostic
    End If
    Set mResult.Root = Nothing
    Set Compile = mResult
End Function

Private Function CheckNode(ByVal node As ATAstNode, ByVal aliases As Collection, ByVal parameters As Collection) As String
    Dim types As New Collection, child As ATAstNode
    Dim definition As ATFieldDefinition, primitive As ATPrimitive
    Dim typeName As String, item As Variant, contextualIndex As Long
    Set mCurrent = node
    node.SourceId = mSourceId
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
                node.PrimitiveId = primitive.CanonicalId
                Set node.Signature = primitive.Signatures(1)
                child.DataType = "TEXT"
                child.SourceId = mSourceId
                child.Text = definition.Id
                mResult.Dependencies.Add "SOURCE:" & definition.Id, Owner()
            ElseIf mContext.Registry.HasSyntax(node.Text, "QUANTIFIER") Then
                typeName = CheckCollection(node, aliases, parameters)
            Else
                Set primitive = mContext.Registry.FindSyntax(node.Text, "FUNCTION")
                If primitive.CanonicalId = "LOOKUP" Then
                    typeName = CheckLookup(node, aliases, parameters, primitive)
                Else
                    For Each child In node.Children
                        types.Add CheckNode(child, aliases, parameters)
                    Next child
                    Set mCurrent = node
                    typeName = mTypes.ResolveSignature(primitive, types, mContext.Registry)
                    ValidateConversions node, mTypes.LastSignature
                    If node.Text = "DATE" Then ValidateDate node
                End If
                node.Text = primitive.CanonicalSyntax
                node.PrimitiveId = primitive.CanonicalId
                Set node.Signature = mTypes.LastSignature
            End If
        Case "UNARY", "BINARY", "RELATION"
            For Each child In node.Children
                types.Add CheckNode(child, aliases, parameters)
            Next child
            Set mCurrent = node
            If node.Kind = "UNARY" And (node.Text = "+" Or node.Text = "-") Then
                Set primitive = mContext.Registry.GetPrimitive("UNARY_" & IIf(node.Text = "+", "PLUS", "MINUS"))
            Else
                Set primitive = mContext.Registry.FindSyntax(node.Text, "OPERATOR")
            End If
            If node.Children.Count = 2 Then
                If node.Children(2).ContextualScalar Then contextualIndex = 2
            End If
            typeName = mTypes.ResolveSignature(primitive, types, mContext.Registry, contextualIndex)
            node.PrimitiveId = primitive.CanonicalId
            Set node.Signature = mTypes.LastSignature
            If mTypes.LastContextWrap > 0 Then WrapContextValue node, mTypes.LastContextWrap
            ValidateConversions node, mTypes.LastSignature
        Case Else
            Err.Raise vbObjectError + 2335, "ATCompiler", "TYPE_INVALID_AST_NODE: " & node.Kind
    End Select
    node.DataType = typeName
    Set mCurrent = node
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

Private Function CheckLookup(ByVal node As ATAstNode, ByVal aliases As Collection, ByVal parameters As Collection, ByVal primitive As ATPrimitive) As String
    Dim source As ATFieldDefinition, column As ATFieldDefinition, keyType As String
    Dim argumentTypes As New Collection, ignored As String
    If node.Children.Count <> 3 Then Err.Raise vbObjectError + 2370, "ATCompiler", "TYPE_LOOKUP_ARITY"
    If node.Children(1).Kind <> "STRING" Or node.Children(3).Kind <> "STRING" Then Err.Raise vbObjectError + 2371, "ATCompiler", "LOOKUP_LITERAL_BINDING_REQUIRED"
    Set source = mContext.FindSource(node.Children(1).Text)
    If source Is Nothing Then Err.Raise vbObjectError + 2334, "ATCompiler", "RESOLVE_UNKNOWN_SOURCE: " & node.Children(1).Text
    Set column = mContext.FindSourceColumn(source.Id, node.Children(3).Text)
    If column Is Nothing Then Err.Raise vbObjectError + 2372, "ATCompiler", "RESOLVE_UNKNOWN_SOURCE_COLUMN: " & node.Children(3).Text
    If Len(source.LookupKeyType) = 0 Then Err.Raise vbObjectError + 2373, "ATCompiler", "LOOKUP_KEY_SCHEMA_REQUIRED: " & source.Id
    keyType = CheckNode(node.Children(2), aliases, parameters)
    If Not mContext.Registry.CanConvert(keyType, source.LookupKeyType) Then Err.Raise vbObjectError + 2374, "ATCompiler", "TYPE_LOOKUP_KEY: " & source.Id
    If source.LookupKeyType = "PORT_SET" And keyType <> "PORT_SET" Then AT_ValidatePortLiteral node.Children(2)
    node.Children(1).DataType = "TEXT"
    node.Children(3).DataType = "TEXT"
    node.Children(1).SourceId = node.SourceId
    node.Children(3).SourceId = node.SourceId
    node.Children(1).Text = source.Id
    node.Children(3).Text = column.Id
    argumentTypes.Add "TEXT"
    argumentTypes.Add keyType
    argumentTypes.Add "TEXT"
    ignored = mTypes.ResolveSignature(primitive, argumentTypes, mContext.Registry)
    mResult.Dependencies.Add "SOURCE:" & source.Id, Owner()
    mResult.Dependencies.Add "SOURCE_COLUMN:" & source.Id & ":" & column.Id, Owner()
    CheckLookup = column.TypeName
End Function

Private Sub ValidateConversions(ByVal node As ATAstNode, ByVal signature As ATSignature)
    Dim index As Long, child As ATAstNode, target As String
    For index = 1 To node.Children.Count
        Set child = node.Children(index)
        target = signature.ArgumentTypes(index)
        If target = "PORT_SET" And child.DataType <> target Then AT_ValidatePortLiteral child
        If target <> child.DataType Then
            If mContext.Registry.CanConvert(child.DataType, target) Then child.ConversionNormalizerId = mContext.Registry.GetPrimitive(target).NormalizerId
        End If
    Next index
End Sub

Private Sub WrapContextValue(ByVal parent As ATAstNode, ByVal index As Long)
    Dim child As ATAstNode, wrapper As New ATAstNode
    Set child = parent.Children(index)
    wrapper.Kind = "LIST"
    wrapper.Start = child.Start
    wrapper.Length = child.Length
    wrapper.SourceId = child.SourceId
    wrapper.Children.Add child
    wrapper.DataType = "LIST<" & child.DataType & ">"
    wrapper.Depth = child.Depth + 1
    If wrapper.Depth > mLimits.MaxNesting Then Err.Raise vbObjectError + 2402, "ATCompiler", "LIMIT_EXPANSION_DEPTH"
    parent.Children.Remove index
    If index <= parent.Children.Count Then parent.Children.Add wrapper, Before:=index Else parent.Children.Add wrapper
End Sub

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
    binding.SourceId = node.SourceId
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
    node.PrimitiveId = primitive.CanonicalId
    Set node.Signature = mTypes.LastSignature
    mCollectionDepth = mCollectionDepth - 1
End Function

Private Function CheckNamed(ByVal node As ATAstNode, ByVal aliases As Collection, ByVal parameters As Collection) As String
    Dim definition As ATNamedExpression, parameter As ATFieldDefinition
    Dim nested As New Collection, child As ATAstNode, parser As New ATParser
    Dim parsed As ATParseResult, index As Long, actualType As String, present As Variant
    Dim supplied As New Collection
    Dim warning As ATDiagnostic
    Dim callerSource As String
    Dim actualNode As ATAstNode, binding As ATArgumentBinding
    Dim bodyAliases As New Collection
    callerSource = mSourceId
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
        warning.SourceId = node.SourceId
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
    If definition.Scope = "ROW" Then
        If aliases.Count = 0 Then Err.Raise vbObjectError + 2358, "ATCompiler", "NAMED_ROW_CONTEXT_REQUIRED"
        Set bodyAliases = aliases
    End If
    For index = 1 To definition.Parameters.Count
        Set parameter = definition.Parameters(index)
        If index <= node.Children.Count Then
            actualType = supplied(index)
            Set actualNode = node.Children(index)
        Else
            If Len(parameter.DefaultExpression) = 0 Then Err.Raise vbObjectError + 2353, "ATCompiler", "TYPE_NAMED_REQUIRED_ARGUMENT"
            mSourceId = "DEFAULT:" & definition.Id & ":" & parameter.Id
            Set parsed = parser.Parse(parameter.DefaultExpression, "", mContext.Registry, mLimits)
            PreserveNestedDiagnostics parsed, mSourceId
            actualType = CheckNode(parsed.Root, bodyAliases, nested)
            Set actualNode = parsed.Root
        End If
        If actualType <> parameter.TypeName And Not mContext.Registry.CanConvert(actualType, parameter.TypeName) Then Err.Raise vbObjectError + 2355, "ATCompiler", "TYPE_NAMED_ARGUMENT"
        If parameter.TypeName = "PORT_SET" And actualType <> parameter.TypeName Then
            If index <= node.Children.Count Then
                AT_ValidatePortLiteral node.Children(index)
            Else
                AT_ValidatePortLiteral parsed.Root
            End If
        End If
        nested.Add parameter
        Set binding = New ATArgumentBinding
        binding.Id = parameter.Id
        binding.DataType = parameter.TypeName
        Set binding.Expression = actualNode
        binding.IsDefault = (index > node.Children.Count)
        If actualType <> parameter.TypeName Then binding.NormalizerId = mContext.Registry.GetPrimitive(parameter.TypeName).NormalizerId
        node.ParameterBindings.Add binding
        mSourceId = callerSource
    Next index
    mSourceId = "NAMED:" & definition.Id
    Set parsed = parser.Parse(definition.Expression, "", mContext.Registry, mLimits)
    PreserveNestedDiagnostics parsed, mSourceId
    actualType = CheckNode(parsed.Root, bodyAliases, nested)
    If actualType <> definition.ReturnType Then Err.Raise vbObjectError + 2357, "ATCompiler", "TYPE_NAMED_RETURN"
    mNamedStack.Remove mNamedStack.Count
    Set node.Expansion = parsed.Root
    mSourceId = callerSource
    CheckNamed = actualType
End Function

Private Sub PreserveNestedDiagnostics(ByVal parsed As ATParseResult, ByVal source As String)
    Dim diagnostic As ATDiagnostic
    For Each diagnostic In parsed.Diagnostics
        diagnostic.SourceId = source
        mResult.Diagnostics.Add diagnostic
    Next diagnostic
    If parsed.Diagnostics.Count > 0 Then Err.Raise vbObjectError + 2356, "ATCompiler", "NAMED_INVALID_EXPRESSION"
End Sub

Private Sub ValidateDate(ByVal node As ATAstNode)
    Dim value As String, year As Long, month As Long, day As Long
    Dim days As Long
    If node.Children(1).Kind <> "STRING" Then Err.Raise vbObjectError + 2360, "ATCompiler", "DATE_LITERAL_REQUIRED"
    value = node.Children(1).Text
    If Len(value) <> 10 Or Mid$(value, 5, 1) <> "-" Or Mid$(value, 8, 1) <> "-" Then Err.Raise vbObjectError + 2361, "ATCompiler", "DATE_INVALID_LITERAL"
    If Not value Like "####-##-##" Then Err.Raise vbObjectError + 2361, "ATCompiler", "DATE_INVALID_LITERAL"
    year = CLng(Left$(value, 4))
    month = CLng(Mid$(value, 6, 2))
    day = CLng(Right$(value, 2))
    If year < 100 Or year > 9999 Then Err.Raise vbObjectError + 2361, "ATCompiler", "DATE_INVALID_LITERAL"
    If month < 1 Or month > 12 Or day < 1 Then Err.Raise vbObjectError + 2361, "ATCompiler", "DATE_INVALID_LITERAL"
    Select Case month
        Case 4, 6, 9, 11: days = 30
        Case 2
            days = 28
            If year Mod 400 = 0 Or (year Mod 4 = 0 And year Mod 100 <> 0) Then days = 29
        Case Else: days = 31
    End Select
    If day > days Then Err.Raise vbObjectError + 2361, "ATCompiler", "DATE_INVALID_LITERAL"
End Sub

