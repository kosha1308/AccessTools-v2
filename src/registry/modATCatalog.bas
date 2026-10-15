Attribute VB_Name = "modATCatalog"
Option Explicit

Public Sub AT_AddCatalog(ByVal registry As ATRegistry)
    Dim id As Variant, primitive As ATPrimitive, syntax As String
    For Each id In Array("STARTS_WITH", "ENDS_WITH", "LIKE", "MATCHES")
        Set primitive = AT_MakePrimitive(CStr(id), "OPERATOR", "TEXT,TEXT", "BOOL", "@NAME " & CStr(id) & " ""text""")
        primitive.SetPresentation CStr(id), "Text relation; locale-independent registered evaluator required"
        registry.Register primitive
    Next id
    For Each id In Array("EMPTY", "NOT_EMPTY", "VALID", "INVALID", "IS_NULL", "IS_NOT_NULL")
        syntax = CStr(id)
        If id = "IS_NULL" Or id = "IS_NOT_NULL" Then syntax = Replace(CStr(id), "_", " ")
        Set primitive = AT_MakePrimitive(CStr(id), "OPERATOR", "T", "BOOL", "@FIELD " & syntax)
        If id = "IS_NULL" Or id = "IS_NOT_NULL" Then primitive.SetSyntax Replace(CStr(id), "_", " ")
        primitive.SetPresentation CStr(id), "Explicit inspection of resolved value state; never silently coerce MISSING or INVALID"
        registry.Register primitive
    Next id
    Set primitive = AT_MakePrimitive("BETWEEN", "OPERATOR", "NUMBER,NUMBER,NUMBER", "BOOL", "@VALUE BETWEEN 1 AND 10")
    primitive.AddSignature "DATE,DATE,DATE", "BOOL"
    primitive.SetPresentation "BETWEEN", "Inclusive lower and upper bounds; AND separates bounds"
    registry.Register primitive
    Set primitive = AT_MakePrimitive("NOT_IN", "OPERATOR", "T,LIST<T>", "BOOL", "@VALUE NOT IN [1,2]")
    primitive.AddSignature "T,SET<T>", "BOOL"
    primitive.AddSignature "IP_EXPRESSION,IP_EXPRESSION", "BOOL"
    primitive.SetSyntax "NOT IN"
    primitive.SetAliasOf registry.GetPrimitive("IN"), True
    registry.Register primitive
    For Each id In Array("UPPER", "TRIM")
        registry.Register AT_MakePrimitive(CStr(id), "FUNCTION", "TEXT", "TEXT", CStr(id) & "(@TEXT)")
    Next id
    registry.Register AT_MakePrimitive("LEN", "FUNCTION", "TEXT", "NUMBER", "LEN(@TEXT)")
    registry.Register AT_MakePrimitive("ABS", "FUNCTION", "NUMBER", "NUMBER", "ABS(@NUMBER)")
    registry.Register AT_MakePrimitive("ROUND", "FUNCTION", "NUMBER,NUMBER", "NUMBER", "ROUND(@NUMBER,2)")
    For Each id In Array("MIN", "MAX")
        Set primitive = AT_MakePrimitive(CStr(id), "FUNCTION", "NUMBER,NUMBER", "NUMBER", CStr(id) & "(@A,@B)")
        primitive.AddSignature "DATE,DATE", "DATE"
        primitive.AddSignature "LIST<NUMBER>", "NUMBER"
        primitive.AddSignature "LIST<DATE>", "DATE"
        registry.Register primitive
    Next id
    For Each id In Array("SUM", "AVG")
        registry.Register AT_MakePrimitive(CStr(id), "FUNCTION", "LIST<NUMBER>", "NUMBER", CStr(id) & "(@NUMBERS)")
    Next id
    registry.Register AT_MakePrimitive("COALESCE", "FUNCTION", "T,T", "T", "COALESCE(@VALUE,@DEFAULT)")
    Set primitive = AT_MakePrimitive("TODAY", "FUNCTION", "", "DATE", "TODAY()")
    primitive.SetDeterminism "CONTEXT"
    primitive.SetPresentation "TODAY", "Date from fixed EvaluationContext run timestamp and configured timezone"
    registry.Register primitive
    registry.Register AT_MakePrimitive("DATE_DIFF", "FUNCTION", "DATE,DATE", "NUMBER", "DATE_DIFF(@START,@END)")
    Set primitive = AT_MakePrimitive("LOOKUP", "FUNCTION", "TEXT,T,TEXT", "T", "LOOKUP(""SOURCE_ID"",@KEY,""COLUMN_ID"")")
    primitive.SetPresentation "LOOKUP", "Source-schema lookup: key type and output column type are checked before execution"
    registry.Register primitive
    Set primitive = AT_MakePrimitive("SOURCE_VALUE", "FUNCTION", "TEXT,T,TEXT", "T", "SOURCE_VALUE(""SOURCE_ID"",@KEY,""COLUMN_ID"")")
    primitive.SetAliasOf registry.GetPrimitive("LOOKUP")
    registry.Register primitive
End Sub

Public Function AT_PrimitiveDescription(ByVal id As String) As String
    Select Case id
        Case "ANY", "OVERLAPS", "INTERSECTS": AT_PrimitiveDescription = "True when normalized actual and expected sets have a nonempty intersection"
        Case "NONE": AT_PrimitiveDescription = "True when normalized actual and expected sets have no intersection"
        Case "COVER", "CONTAINED_BY": AT_PrimitiveDescription = "Full coverage: every actual/requested element is inside the expected/allowed set"
        Case "CONTAINS": AT_PrimitiveDescription = "Contains expected text or all expected set elements according to the selected typed signature"
        Case "EXACT": AT_PrimitiveDescription = "Equality of normalized sets; ordering and duplicates do not change set equality"
        Case "IN": AT_PrimitiveDescription = "Typed membership, or full domain containment for IP expressions"
        Case "AND": AT_PrimitiveDescription = "Boolean conjunction with deterministic left-to-right short circuit"
        Case "OR": AT_PrimitiveDescription = "Boolean disjunction with deterministic left-to-right short circuit"
        Case "NOT": AT_PrimitiveDescription = "Negates an OK Boolean; non-OK states are propagated without Boolean coercion"
        Case "OP_EQUALS", "OP_NOT_EQUALS": AT_PrimitiveDescription = "Equality/inequality of compatible typed operands"
        Case "GREATER", "LESS", "GREATER_EQUAL", "LESS_EQUAL": AT_PrimitiveDescription = "Ordered comparison of compatible numbers, dates or text"
        Case "ADD", "SUBTRACT", "MULTIPLY", "DIVIDE", "REMAINDER", "UNARY_PLUS", "UNARY_MINUS": AT_PrimitiveDescription = "Typed numeric arithmetic; evaluator must return diagnostics for overflow and division by zero"
        Case "LOWER", "UPPER": AT_PrimitiveDescription = "Text case conversion under the registered normalization policy"
        Case "TRIM": AT_PrimitiveDescription = "Remove leading/trailing permitted whitespace without tokenizing free text"
        Case "LEN": AT_PrimitiveDescription = "Length of text in UTF-16 code units"
        Case "ABS": AT_PrimitiveDescription = "Absolute numeric value"
        Case "ROUND": AT_PrimitiveDescription = "Numeric rounding to the requested integral decimal precision; ties to even"
        Case "MIN", "MAX": AT_PrimitiveDescription = "Minimum/maximum of two compatible scalars or one typed list; empty lists require explicit policy"
        Case "SUM", "AVG": AT_PrimitiveDescription = "Sum/arithmetic mean of a numeric list; AVG on empty input is an error"
        Case "COALESCE": AT_PrimitiveDescription = "Two-argument typed default for explicit NULL; does not suppress ERROR, SKIPPED or MANUAL_REVIEW"
        Case "DATE_DIFF": AT_PrimitiveDescription = "Signed whole Gregorian days from start date to end date"
        Case "DATE_LITERAL": AT_PrimitiveDescription = "Locale-independent Gregorian ISO date constructor for years 0100 through 9999"
        Case "ALLROWS", "ANYROW", "NOROWS": AT_PrimitiveDescription = "Typed row predicate quantifier with lexical aliases and explicit empty-collection semantics"
        Case "FILTER": AT_PrimitiveDescription = "Select rows satisfying a typed Boolean predicate while preserving their schema"
        Case "COUNT", "DISTINCT_COUNT": AT_PrimitiveDescription = "Count collection elements or distinct normalized typed list elements"
        Case "SOURCE": AT_PrimitiveDescription = "Bind to a registered source schema and add a dependency without opening source data"
        Case Else: AT_PrimitiveDescription = "Versioned registered capability; signatures and normalizer/validation policy define its compiler contract"
    End Select
End Function
