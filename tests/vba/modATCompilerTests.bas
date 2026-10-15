Attribute VB_Name = "modATCompilerTests"
Option Explicit

Public Sub AT_TestCompiler()
    Dim context As New ATCompileContext
    Dim compiler As New ATCompiler
    Dim result As ATCompileResult
    Dim expression As ATNamedExpression
    Dim item As Variant
    Dim rejected As Boolean
    context.AddField "A", "NUMBER"
    context.AddField "B", "NUMBER"
    context.AddField "NAME", "TEXT"
    context.AddField "TECHNICAL_ROWS", "COLLECTION<ROW<TECHNICAL>>"
    context.AddField "ADDRESS_ROWS", "COLLECTION<ROW<ADDRESS>>"
    context.AddField "TECH_PROTOCOL", "TEXT", "TECHNICAL"
    context.AddField "TECH_DEST_PORT", "PORT_SET", "TECHNICAL"
    context.AddField "IP", "IP_EXPRESSION", "TECHNICAL"
    context.AddField "IP", "IP_EXPRESSION", "ADDRESS"
    context.AddSource "ALLOWED_NETWORKS", "IP_EXPRESSION"
    Set expression = New ATNamedExpression
    expression.Initialize "POSITIVE", "BOOL", "@A > 0"
    context.AddNamed expression
    Set expression = New ATNamedExpression
    expression.Initialize "OVER_LIMIT", "BOOL", "@A > $LIMIT"
    expression.AddParameter "LIMIT", "NUMBER", "10"
    context.AddNamed expression
    For Each item In Array("@A + @B > 3", "LOWER(@NAME) = ""test""", "$POSITIVE", "$OVER_LIMIT()", "$OVER_LIMIT(20)", "DATE(""2026-10-15"") > DATE(""2026-01-01"")", "ALLROWS(@TECHNICAL_ROWS AS t, @t.TECH_DEST_PORT COVER: 443,8443)", "ALLROWS(@TECHNICAL_ROWS AS t, ANYROW(@ADDRESS_ROWS AS a, @a.IP INTERSECTS @t.IP))", "COUNT(@TECHNICAL_ROWS) > 2", "ANYROW(@TECHNICAL_ROWS AS t, @t.IP CONTAINED_BY SOURCE(""ALLOWED_NETWORKS""))")
        Set result = compiler.Compile(CStr(item), context)
        AT_Assert result.Diagnostics.Count = 0 And Not result.Root Is Nothing, "compile valid: " & CStr(item)
    Next item
    For Each item In Array("@UNKNOWN = 1", "@A = @NAME", "LOWER(12) = ""x""", "LOWER(@NAME, @NAME) = ""x""", "MISSING_FUNCTION() = TRUE", "SOURCE(""UNKNOWN"") = TRUE", "@A + 1", "@A AND TRUE", "$UNKNOWN", "$OVER_LIMIT(""bad"")", "ANYROW(@TECHNICAL_ROWS AS t, @a.IP = @t.IP)", "ALLROWS(@TECHNICAL_ROWS AS t, ANYROW(@ADDRESS_ROWS AS a, @IP = @t.IP))", "DATE(""2026-02-30"") = DATE(""2026-03-02"")")
        Set result = compiler.Compile(CStr(item), context)
        AT_Assert result.Diagnostics.Count > 0 And result.Root Is Nothing, "compile invalid: " & CStr(item)
    Next item
    Set result = compiler.Compile("ANYROW(@TECHNICAL_ROWS AS t, @t.IP CONTAINED_BY SOURCE(""ALLOWED_NETWORKS"")) AND $POSITIVE", context)
    AT_Assert result.Dependencies.Contains("SOURCE:ALLOWED_NETWORKS"), "compile source dependency"
    AT_Assert result.Dependencies.Contains("NAMED:POSITIVE"), "compile named dependency"
    Set expression = New ATNamedExpression
    expression.Initialize "LOOP_A", "BOOL", "$LOOP_B"
    context.AddNamed expression
    Set expression = New ATNamedExpression
    expression.Initialize "LOOP_B", "BOOL", "$LOOP_A"
    context.AddNamed expression
    Set result = compiler.Compile("$LOOP_A", context)
    AT_Assert result.Diagnostics.Count > 0 And result.Root Is Nothing, "compile indirect cycle rejected"
    Set expression = New ATNamedExpression
    expression.Initialize "DIRECT", "BOOL", "$DIRECT"
    context.AddNamed expression
    Set result = compiler.Compile("$DIRECT", context)
    AT_Assert result.Diagnostics.Count > 0, "compile direct cycle rejected"
    Set expression = New ATNamedExpression
    expression.Initialize "OLD", "BOOL", "TRUE"
    expression.Status = "DEPRECATED"
    expression.ReplacedBy = "POSITIVE"
    context.AddNamed expression
    Set result = compiler.Compile("$OLD", context)
    AT_Assert Not result.Root Is Nothing And result.Diagnostics.Count = 1, "deprecated named remains executable with warning"
    If result.Diagnostics.Count > 0 Then
        AT_Assert result.Diagnostics(1).Severity = "WARNING", "deprecated named warning severity"
    End If
    Set expression = New ATNamedExpression
    expression.Initialize "PARAMS", "BOOL", "TRUE"
    expression.AddParameter "P", "BOOL"
    On Error Resume Next
    expression.AddParameter "p", "BOOL"
    rejected = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert rejected, "duplicate parameter canonical ID rejected"
    Set result = compiler.Compile("$OVER_LIMIT()", context)
    If Not result.Root Is Nothing Then
        AT_Assert result.Root.ParameterBindings.Count = 1, "compiled named call retains default binding"
        If result.Root.ParameterBindings.Count = 1 Then
            AT_Assert result.Root.ParameterBindings(1).Expression.Text = "10", "compiled default retains parsed AST"
            AT_Assert result.Root.ParameterBindings(1).DataType = "NUMBER", "compiled default binding keeps declared type"
        End If
    End If
    context.AddField "ROWS", "COLLECTION<ROW<ROW_TABLE>>"
    context.AddField "FLAG", "BOOL", "ROW_TABLE"
    Set expression = New ATNamedExpression
    expression.Initialize "DOC_CONDITION", "BOOL", "@FLAG"
    context.AddNamed expression
    Set result = compiler.Compile("ANYROW(@ROWS AS t, $DOC_CONDITION)", context)
    AT_Assert result.Root Is Nothing, "document named expression does not capture caller row"
    Set expression = New ATNamedExpression
    expression.Initialize "ROW_CONDITION", "BOOL", "@FLAG"
    expression.Scope = "ROW"
    context.AddNamed expression
    Set result = compiler.Compile("ANYROW(@ROWS AS t, $ROW_CONDITION)", context)
    AT_Assert Not result.Root Is Nothing, "row named expression explicitly inherits row context"
    Set expression = New ATNamedExpression
    expression.Initialize "DOC_DEFAULT", "BOOL", "$P"
    expression.AddParameter "P", "BOOL", "@FLAG"
    context.AddNamed expression
    Set result = compiler.Compile("ANYROW(@ROWS AS t, $DOC_DEFAULT)", context)
    AT_Assert result.Root Is Nothing, "document default does not capture caller row"
    Set expression = New ATNamedExpression
    expression.Initialize "ORDERED_DEFAULTS", "BOOL", "$SECOND"
    expression.AddParameter "FIRST", "BOOL", "TRUE"
    expression.AddParameter "SECOND", "BOOL", "$FIRST"
    context.AddNamed expression
    Set result = compiler.Compile("$ORDERED_DEFAULTS", context)
    AT_Assert Not result.Root Is Nothing, "default references an earlier declared parameter"
    Set result = compiler.Compile("@A + @B > 3", context)
    If Not result.Root Is Nothing Then
        AT_Assert result.Root.PrimitiveId = "GREATER", "compiled AST retains selected primitive ID"
        AT_Assert result.Root.Children(1).Signature.ArgumentTypes.Count = 2, "compiled AST retains selected signature"
        AT_Assert result.DSLVersion = "1.0.0", "compiled result retains language version"
    End If
End Sub
