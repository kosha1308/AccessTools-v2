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
End Sub
