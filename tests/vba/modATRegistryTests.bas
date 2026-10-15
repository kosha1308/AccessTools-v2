Attribute VB_Name = "modATRegistryTests"
Option Explicit

Public Sub AT_TestRegistry()
    Dim registry As New ATRegistry
    Dim primitive As ATPrimitive
    Dim diagnostic As New ATDiagnostic
    Dim signature As ATSignature
    Dim rejected As Boolean
    registry.Register AT_MakePrimitive("TEXT", "DATATYPE", "", "TEXT", "TEXT")
    Set primitive = AT_MakePrimitive("lower", "FUNCTION", "TEXT", "TEXT", "LOWER(@NAME)")
    registry.Register primitive
    AT_Assert registry.Contains("LOWER"), "registry canonical lookup"
    AT_Assert registry.GetPrimitive("lower").Id = "LOWER", "registry canonical ID"
    AT_Assert registry.GetPrimitive("lower").Signatures.Count = 1, "registry keeps signatures"
    Set signature = registry.GetPrimitive("LOWER").Signatures(1)
    AT_Assert signature.ArgumentTypes(1) = "TEXT", "registry operand type"
    On Error Resume Next
    signature.Initialize "NUMBER", "NUMBER"
    rejected = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert rejected, "registered signature immutable"
    On Error Resume Next
    registry.Register AT_MakePrimitive("UNKNOWN_TYPE", "FUNCTION", "TEXT", "UNREGISTERED", "example")
    rejected = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert rejected, "unregistered signature type rejected"
    Set primitive = AT_MakePrimitive("LOWER_ALIAS", "FUNCTION", "TEXT", "TEXT", "LOWER(@NAME)")
    primitive.SetSyntax "LOWER"
    On Error Resume Next
    registry.Register primitive
    rejected = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert rejected, "ambiguous registered syntax rejected"
    On Error Resume Next
    registry.Register primitive
    rejected = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert rejected, "duplicate ID rejected"
    On Error Resume Next
    Set primitive = AT_MakePrimitive("1BAD", "FUNCTION", "TEXT", "TEXT", "example")
    rejected = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert rejected, "invalid ID rejected"
    On Error Resume Next
    Set primitive = registry.GetPrimitive("UNKNOWN")
    rejected = Err.Number <> 0 And InStr(Err.Description, "REG_UNKNOWN_ID") > 0
    Err.Clear
    On Error GoTo 0
    AT_Assert rejected, "unknown ID diagnosed"
    Set primitive = New ATPrimitive
    On Error Resume Next
    primitive.Initialize "BROKEN", "FUNCTION", "description", "evaluator", "1.0.0", "ACTIVE", "PURE", "example", "positive negative"
    registry.Register primitive
    rejected = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert rejected, "incomplete primitive rejected"
    diagnostic.Initialize "LEX_UNEXPECTED", "LEXICAL", "ERROR", 12, 0, "Unexpected end"
    AT_Assert diagnostic.Start = 12 And diagnostic.Length = 0, "EOF diagnostic span"
    On Error Resume Next
    diagnostic.Initialize "BAD", "LEXICAL", "ERROR", 0, -1, "bad span"
    rejected = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert rejected, "invalid diagnostic span rejected"
    Set registry = AT_CreateRegistry()
    AT_Assert registry.Contains("BOOL") And registry.Contains("COVER"), "bootstrap types and operators"
    AT_Assert Not registry.Contains("EVAL"), "arbitrary evaluation absent"
    Set primitive = AT_MakePrimitive("NEGATED_ALIAS", "OPERATOR", "T,LIST<T>", "BOOL", "@A EXCLUDED_FROM [1]")
    primitive.SetSyntax "EXCLUDED_FROM"
    primitive.SetAliasOf registry.GetPrimitive("NOT_IN")
    AT_Assert primitive.NegatedAlias, "alias chaining preserves negation"
End Sub
