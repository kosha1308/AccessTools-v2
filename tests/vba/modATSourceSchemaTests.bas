Attribute VB_Name = "modATSourceSchemaTests"
Option Explicit

Public Sub AT_TestSourceSchema()
    Dim context As New ATCompileContext, compiler As New ATCompiler
    Dim result As ATCompileResult, serializer As New ATSerializer
    Dim item As Variant
    context.AddField "EMAIL", "EMAIL"
    context.AddField "DEPARTMENT", "TEXT"
    context.AddSource "EMPLOYEE_DIRECTORY", "COLLECTION<ROW<EMPLOYEE>>", "EMAIL"
    context.AddSourceColumn "EMPLOYEE_DIRECTORY", "DEPARTMENT", "TEXT"
    context.AddSourceColumn "EMPLOYEE_DIRECTORY", "ACTIVE", "BOOL"
    For Each item In Array("LOOKUP(""employee_directory"",@EMAIL,""Department"") = @DEPARTMENT", "SOURCE_VALUE(""EMPLOYEE_DIRECTORY"",@EMAIL,""ACTIVE"") = TRUE")
        Set result = compiler.Compile(CStr(item), context)
        AT_Assert Not result.Root Is Nothing, "lookup result type from source column schema"
        AT_Assert result.Dependencies.Contains("SOURCE:EMPLOYEE_DIRECTORY"), "lookup source dependency"
        AT_Assert result.Dependencies.Contains("SOURCE_COLUMN:EMPLOYEE_DIRECTORY:DEPARTMENT") Or result.Dependencies.Contains("SOURCE_COLUMN:EMPLOYEE_DIRECTORY:ACTIVE"), "lookup output column dependency"
    Next item
    Set result = compiler.Compile("SOURCE_VALUE(""employee_directory"",@EMAIL,""Department"") = @DEPARTMENT", context)
    If Not result.Root Is Nothing Then AT_Assert serializer.Serialize(result.Root) = "(LOOKUP(""EMPLOYEE_DIRECTORY"",@EMAIL,""DEPARTMENT"") = @DEPARTMENT)", "lookup canonical function/source/column identity"
    For Each item In Array("LOOKUP(""UNKNOWN"",@EMAIL,""DEPARTMENT"") = @DEPARTMENT", "LOOKUP(""EMPLOYEE_DIRECTORY"",@EMAIL,""UNKNOWN"") = @DEPARTMENT", "LOOKUP(""EMPLOYEE_DIRECTORY"",@DEPARTMENT,""DEPARTMENT"") = @DEPARTMENT", "LOOKUP(@DEPARTMENT,@EMAIL,""DEPARTMENT"") = @DEPARTMENT", "LOOKUP(""EMPLOYEE_DIRECTORY"",@EMAIL,@DEPARTMENT) = @DEPARTMENT", "LOOKUP(""EMPLOYEE_DIRECTORY"",@EMAIL) = @DEPARTMENT", "LOOKUP(""EMPLOYEE_DIRECTORY"",@EMAIL,""ACTIVE"") = @DEPARTMENT")
        Set result = compiler.Compile(CStr(item), context)
        AT_Assert result.Root Is Nothing And result.Diagnostics.Count > 0, "invalid lookup schema/arity/types rejected"
    Next item
End Sub
