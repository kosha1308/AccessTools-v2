Attribute VB_Name = "modATDiagnosticTests"
Option Explicit

Public Sub AT_TestDiagnostics()
    Dim context As New ATCompileContext, compiler As New ATCompiler
    Dim result As ATCompileResult, item As Variant, parts As Variant
    Dim diagnostic As ATDiagnostic, named As ATNamedExpression
    context.AddField "N", "NUMBER"
    context.AddField "NAME", "TEXT"
    For Each item In Array("1 2|PARSE_TRAILING_TOKEN|3|1", "(1|PARSE_EXPECTED_TOKEN|3|0", "@UNKNOWN = 1|RESOLVE_UNKNOWN_FIELD|1|8", "@N + @NAME > 0|TYPE_SIGNATURE_MISMATCH|1|10", "@N|TYPE_RULE_ROOT|1|2")
        parts = Split(CStr(item), "|")
        Set result = compiler.Compile(CStr(parts(0)), context)
        AT_Assert result.Root Is Nothing And result.Diagnostics.Count = 1, "diagnostic no partial executable root"
        If result.Diagnostics.Count = 1 Then
            Set diagnostic = result.Diagnostics(1)
            AT_Assert diagnostic.Code = parts(1), "diagnostic stable code"
            AT_Assert diagnostic.Start = CLng(parts(2)) And diagnostic.Length = CLng(parts(3)), "diagnostic exact original span"
        End If
    Next item
    For Each item In Array("DATE(""9999-99-99"") = TODAY()", "DATE(""2026-02-30"") = TODAY()", "DATE(""1900-02-29"") = TODAY()")
        Set result = compiler.Compile(CStr(item), context)
        If result.Diagnostics.Count > 0 Then AT_Assert result.Diagnostics(1).Code = "DATE_INVALID_LITERAL", "invalid dates have stable code independent of native error language"
    Next item
    Set named = New ATNamedExpression
    named.Initialize "BAD_DEFINITION", "BOOL", "@MISSING = 1"
    context.AddNamed named
    Set result = compiler.Compile("TRUE AND $BAD_DEFINITION", context)
    If result.Diagnostics.Count > 0 Then AT_Assert result.Diagnostics(1).SourceId = "NAMED:BAD_DEFINITION", "definition diagnostic names original source"
    Set named = New ATNamedExpression
    named.Initialize "BAD_SYNTAX", "BOOL", "1 +"
    context.AddNamed named
    Set result = compiler.Compile("$BAD_SYNTAX", context)
    If result.Diagnostics.Count > 0 Then
        AT_Assert result.Diagnostics(1).Code = "PARSE_EXPECTED_VALUE", "definition parser diagnostic preserved"
        AT_Assert result.Diagnostics(1).SourceId = "NAMED:BAD_SYNTAX", "definition parser source preserved"
    End If
    AT_TestRuntimeProvenance
End Sub

Private Sub AT_TestRuntimeProvenance()
    Dim context As New ATCompileContext, compiler As New ATCompiler
    Dim result As ATCompileResult, named As ATNamedExpression, node As ATAstNode
    context.AddField "N", "NUMBER"
    context.AddField "VALUES", "LIST<NUMBER>"
    context.AddField "ROWS", "COLLECTION<ROW<T>>"
    context.AddSource "S", "LIST<NUMBER>", "NUMBER"
    context.AddSourceColumn "S", "N", "NUMBER"
    Set named = New ATNamedExpression
    named.Initialize "OLD", "BOOL", "TRUE"
    named.Status = "DEPRECATED"
    context.AddNamed named
    Set named = New ATNamedExpression
    named.Initialize "OUTER", "BOOL", "$OLD"
    context.AddNamed named
    Set result = compiler.Compile("$OUTER", context)
    AT_Assert result.Diagnostics.Count = 1, "nested warning present"
    If result.Diagnostics.Count = 1 Then
        AT_Assert result.Diagnostics(1).SourceId = "NAMED:OUTER", "nested warning original source"
        AT_Assert result.Diagnostics(1).Start = 1 And result.Diagnostics(1).Length = 4, "nested warning exact span"
    End If
    Set result = compiler.Compile("SOURCE(""S"")", context, , "LIST<NUMBER>")
    AT_Assert Not result.Root Is Nothing, "source provenance compiles"
    If Not result.Root Is Nothing Then AT_Assert result.Root.Children(1).SourceId = "RULE", "source literal source"
    Set result = compiler.Compile("LOOKUP(""S"",@N,""N"") = 1", context)
    AT_Assert Not result.Root Is Nothing, "lookup provenance compiles"
    If Not result.Root Is Nothing Then
        Set node = result.Root.Children(1)
        AT_Assert node.Children(1).SourceId = "RULE" And node.Children(3).SourceId = "RULE", "lookup literals source"
        AT_Assert node.Children(1).Start = 8 And node.Children(3).Start = 15, "lookup literals exact span"
    End If
    Set result = compiler.Compile("ALLROWS(@ROWS AS t, TRUE)", context)
    AT_Assert Not result.Root Is Nothing, "binding provenance compiles"
    If Not result.Root Is Nothing Then AT_Assert result.Root.Children(1).SourceId = "RULE", "binding source"
    Set named = New ATNamedExpression
    named.Initialize "COVER_ONE", "BOOL", "@VALUES COVER: 1"
    context.AddNamed named
    Set result = compiler.Compile("$COVER_ONE", context)
    AT_Assert Not result.Root Is Nothing, "contextual list provenance compiles"
    If Not result.Root Is Nothing Then
        Set node = result.Root.Expansion.Children(2)
        AT_Assert node.SourceId = "NAMED:COVER_ONE", "contextual list source"
        AT_Assert node.Start = 16 And node.Length = 1, "contextual list span"
    End If
End Sub
