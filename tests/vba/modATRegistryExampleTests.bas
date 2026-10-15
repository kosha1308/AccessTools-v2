Attribute VB_Name = "modATRegistryExampleTests"
Option Explicit

Public Sub AT_TestRegistryExamples()
    Dim context As New ATCompileContext, compiler As New ATCompiler
    Dim primitive As ATPrimitive, result As ATCompileResult
    Dim field As Variant, returnType As String
    For Each field In Array("NAME", "TEXT", "FIELD")
        context.AddField CStr(field), "TEXT"
    Next field
    For Each field In Array("A", "B", "NUMBER", "VALUE", "DEFAULT", "KEY")
        context.AddField CStr(field), "NUMBER"
    Next field
    For Each field In Array("IP", "OTHER", "ALLOWED")
        context.AddField CStr(field), "IP_EXPRESSION"
    Next field
    context.AddField "PORTS", "PORT_SET"
    context.AddField "NUMBERS", "LIST<NUMBER>"
    context.AddField "VALUES", "LIST<NUMBER>"
    context.AddField "START", "DATE"
    context.AddField "END", "DATE"
    context.AddField "ROWS", "COLLECTION<ROW<ROW_TABLE>>"
    context.AddSource "SOURCE_ID", "NUMBER", "NUMBER"
    context.AddSourceColumn "SOURCE_ID", "COLUMN_ID", "NUMBER"
    context.AddSource "ALLOWED_NETWORKS", "IP_EXPRESSION"
    For Each primitive In context.Registry.Entries
        If primitive.Category <> "DATATYPE" Then
            returnType = primitive.Signatures(1).ReturnType
            If returnType = "T" Then returnType = "NUMBER"
            If returnType = "COLLECTION<ROW<T>>" Then returnType = "COLLECTION<ROW<ROW_TABLE>>"
            Set result = compiler.Compile(primitive.Examples, context, "", returnType)
            AT_Assert Not result.Root Is Nothing, "registry example compiles: " & primitive.Id
        End If
    Next primitive
End Sub
