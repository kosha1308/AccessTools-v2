Attribute VB_Name = "modATRowBindingTests"
Option Explicit
Public Sub AT_TestRowBinding()
    Dim schema As New ATCompileContext, compiler As New ATCompiler, result As ATCompileResult, node As ATAstNode
    Dim named As New ATNamedExpression
    schema.AddField "ROWS", "COLLECTION<ROW<T>>"
    schema.AddField "N", "NUMBER", "T"
    schema.AddField "DOC", "NUMBER"
    Set result = compiler.Compile("ALLROWS(@ROWS AS t, @N > @DOC)", schema)
    AT_Assert Not result.Root Is Nothing, "row binding fixture compiles"
    If Not result.Root Is Nothing Then
        Set node = result.Root.Children(2).Children(1)
        AT_Assert node.FieldId = "N" And node.TableId = "T" And node.RowAlias = "T", "unqualified row bound exactly"
        Set node = result.Root.Children(2).Children(2)
        AT_Assert node.FieldId = "DOC" And node.RowAlias = "" And node.TableId = "", "document field bound separately"
    End If
    Set result = compiler.Compile("ANYROW(@ROWS AS r, @r.N > 0)", schema)
    If Not result.Root Is Nothing Then AT_Assert result.Root.Children(2).Children(1).RowAlias = "R", "qualified alias preserved"
    named.Initialize "ROW_CHECK", "BOOL", "@N > 0"
    named.Scope = "ROW"
    schema.AddNamed named
    Set result = compiler.Compile("ALLROWS(@ROWS AS t, $ROW_CHECK)", schema)
    AT_Assert Not result.Root Is Nothing, "ROW definition compiles"
    If Not result.Root Is Nothing Then
        Set node = result.Root.Children(2)
        AT_Assert node.NamedScope = "ROW", "ROW definition binding retained"
        AT_Assert node.Expansion.Children(1).RowAlias = "T", "ROW definition field retains caller alias"
    End If
    Set named = New ATNamedExpression
    named.Initialize "DOC_CHECK", "BOOL", "@DOC > 0"
    schema.AddNamed named
    Set result = compiler.Compile("ALLROWS(@ROWS AS t, $DOC_CHECK)", schema)
    If Not result.Root Is Nothing Then AT_Assert result.Root.Children(2).NamedScope = "DOCUMENT", "DOCUMENT definition binding retained"
End Sub
