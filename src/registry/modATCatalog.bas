Attribute VB_Name = "modATCatalog"
Option Explicit

Public Sub AT_AddCatalog(ByVal registry As ATRegistry)
    Dim id As Variant, primitive As ATPrimitive
    For Each id In Array("STARTS_WITH", "ENDS_WITH", "LIKE", "MATCHES")
        Set primitive = AT_MakePrimitive(CStr(id), "OPERATOR", "TEXT,TEXT", "BOOL", "@NAME " & CStr(id) & " ""text""")
        primitive.SetPresentation CStr(id), "Text relation; locale-independent registered evaluator required"
        registry.Register primitive
    Next id
    For Each id In Array("EMPTY", "NOT_EMPTY", "VALID", "INVALID", "IS_NULL", "IS_NOT_NULL")
        Set primitive = AT_MakePrimitive(CStr(id), "OPERATOR", "T", "BOOL", "@FIELD " & Replace(CStr(id), "_", " "))
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
End Sub
