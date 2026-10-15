Attribute VB_Name = "modATContracts"
Option Explicit

Public Enum ATEvalState
    atEvalOK = 0
    atEvalError = 1
    atEvalSkipped = 2
    atEvalManualReview = 3
End Enum

Public Enum ATValueState
    atValuePresent = 0
    atValueEmpty = 1
    atValueMissing = 2
    atValueInvalid = 3
End Enum

Public Function AT_CanonicalId(ByVal value As String) As String
    Dim i As Long
    Dim code As Long
    If Len(value) = 0 Then Err.Raise vbObjectError + 2140, "ATContracts", "REG_INVALID_ID: empty"
    For i = 1 To Len(value)
        code = AscW(Mid$(value, i, 1))
        If code >= 97 And code <= 122 Then code = code - 32
        If Not ((code >= 65 And code <= 90) Or (i > 1 And ((code >= 48 And code <= 57) Or code = 95))) Then
            Err.Raise vbObjectError + 2140, "ATContracts", "REG_INVALID_ID: " & value
        End If
        AT_CanonicalId = AT_CanonicalId & ChrW$(code)
    Next i
End Function

Public Function AT_MakePrimitive(ByVal idValue As String, ByVal categoryValue As String, ByVal argumentTypes As String, ByVal resultType As String, ByVal example As String) As ATPrimitive
    Dim item As New ATPrimitive
    item.Initialize idValue, categoryValue, idValue, idValue, "1.0.0", "ACTIVE", "PURE", example, "positive; negative; empty; boundary; malformed"
    If categoryValue = "DATATYPE" Then item.SetTypeContract idValue & "_V1", "FIELD_ONLY", "INVALID_IS_ERROR"
    item.AddSignature argumentTypes, resultType
    Set AT_MakePrimitive = item
End Function

Public Function AT_CreateRegistry() As ATRegistry
    Dim registry As New ATRegistry
    Dim item As Variant
    Dim primitive As ATPrimitive
    For Each item In Array("TEXT", "NUMBER", "BOOL", "DATE", "DATETIME", "EMAIL", "PROTOCOL", "HOSTNAME", "PORT", "PORT_SET", "IP", "IP_EXPRESSION", "ENUM", "FIELD_REF", "SOURCE_REF")
        Set primitive = AT_MakePrimitive(CStr(item), "DATATYPE", "", CStr(item), CStr(item))
        Select Case CStr(item)
            Case "TEXT": primitive.SetTypeContract "TEXT_V1", "STRING", "PRESERVE_FREE_TEXT"
            Case "NUMBER": primitive.SetTypeContract "DECIMAL_TEXT_V1", "DECIMAL", "INVARIANT_DECIMAL"
            Case "BOOL": primitive.SetTypeContract "BOOL_V1", "TRUE_FALSE", "EXPLICIT_MAPPING_ONLY"
            Case "DATE": primitive.SetTypeContract "ISO_DATE_V1", "DATE_CALL", "ISO_GREGORIAN_0100_9999"
            Case "PORT", "PORT_SET": primitive.SetTypeContract "PORT_SET_V1", "REGISTERED_NUMERIC_CONVERSION", "UNSIGNED_16_BIT_INTEGER"
        End Select
        registry.Register primitive
    Next item
    For Each item In Array("ANY", "NONE", "COVER", "CONTAINS", "EXACT", "OVERLAPS")
        Set primitive = AT_MakePrimitive(CStr(item), "OPERATOR", "PORT_SET,PORT_SET", "BOOL", "@PORTS " & CStr(item) & ": 443")
        primitive.AddSignature "IP_EXPRESSION,IP_EXPRESSION", "BOOL"
        primitive.AddSignature "LIST<T>,LIST<T>", "BOOL"
        primitive.AddSignature "SET<T>,SET<T>", "BOOL"
        If CStr(item) = "CONTAINS" Then primitive.AddSignature "TEXT,TEXT", "BOOL"
        registry.Register primitive
    Next item
    registry.Register AT_MakePrimitive("LOWER", "FUNCTION", "TEXT", "TEXT", "LOWER(@NAME)")
    registry.Register AT_MakePrimitive("INTERSECTS", "OPERATOR", "IP_EXPRESSION,IP_EXPRESSION", "BOOL", "@IP INTERSECTS @OTHER")
    registry.Register AT_MakePrimitive("CONTAINED_BY", "OPERATOR", "IP_EXPRESSION,IP_EXPRESSION", "BOOL", "@IP CONTAINED_BY @ALLOWED")
    Set primitive = AT_MakePrimitive("IN", "OPERATOR", "IP_EXPRESSION,IP_EXPRESSION", "BOOL", "@IP IN SOURCE(""ALLOWED_NETWORKS"")")
    primitive.AddSignature "T,LIST<T>", "BOOL"
    primitive.AddSignature "T,SET<T>", "BOOL"
    registry.Register primitive
    AT_AddSyntax registry, "OP_EQUALS", "=", "T,T", "BOOL"
    AT_AddSyntax registry, "OP_NOT_EQUALS", "!=", "T,T", "BOOL"
    For Each item In Array("AND", "OR")
        AT_AddSyntax registry, CStr(item), CStr(item), "BOOL,BOOL", "BOOL"
    Next item
    AT_AddSyntax registry, "NOT", "NOT", "BOOL", "BOOL"
    For Each item In Array("ADD|+", "SUBTRACT|-", "MULTIPLY|*", "DIVIDE|/", "REMAINDER|%")
        AT_AddSyntax registry, Split(CStr(item), "|")(0), Split(CStr(item), "|")(1), "NUMBER,NUMBER", "NUMBER"
    Next item
    AT_AddSyntax registry, "UNARY_PLUS", "UNARY_PLUS", "NUMBER", "NUMBER"
    AT_AddSyntax registry, "UNARY_MINUS", "UNARY_MINUS", "NUMBER", "NUMBER"
    For Each item In Array("GREATER|>", "LESS|<", "GREATER_EQUAL|>=", "LESS_EQUAL|<=")
        Set primitive = AT_MakePrimitive(Split(CStr(item), "|")(0), "OPERATOR", "NUMBER,NUMBER", "BOOL", "@A " & Split(CStr(item), "|")(1) & " @B")
        primitive.SetSyntax Split(CStr(item), "|")(1)
        primitive.AddSignature "DATE,DATE", "BOOL"
        primitive.AddSignature "TEXT,TEXT", "BOOL"
        registry.Register primitive
    Next item
    Set primitive = AT_MakePrimitive("DATE_LITERAL", "FUNCTION", "TEXT", "DATE", "DATE(""2026-10-15"")")
    primitive.SetSyntax "DATE"
    registry.Register primitive
    For Each item In Array("ALLROWS", "ANYROW", "NOROWS")
        registry.Register AT_MakePrimitive(CStr(item), "QUANTIFIER", "COLLECTION<ROW<T>>,BOOL", "BOOL", CStr(item) & "(@ROWS AS t, TRUE)")
    Next item
    registry.Register AT_MakePrimitive("FILTER", "QUANTIFIER", "COLLECTION<ROW<T>>,BOOL", "COLLECTION<ROW<T>>", "FILTER(@ROWS AS t, TRUE)")
    registry.Register AT_MakePrimitive("COUNT", "FUNCTION", "COLLECTION<ROW<T>>", "NUMBER", "COUNT(@ROWS)")
    registry.Register AT_MakePrimitive("DISTINCT_COUNT", "FUNCTION", "LIST<T>", "NUMBER", "DISTINCT_COUNT(@VALUES)")
    registry.Register AT_MakePrimitive("EXISTS", "FUNCTION", "COLLECTION<ROW<T>>", "BOOL", "EXISTS(@ROWS)")
    registry.Register AT_MakePrimitive("SOURCE", "FUNCTION", "TEXT", "T", "SOURCE(""SOURCE_ID"")")
    registry.AllowConversion "LIST<NUMBER>", "PORT_SET"
    registry.AllowConversion "NUMBER", "PORT_SET"
    AT_AddCatalog registry
    Set AT_CreateRegistry = registry
End Function

Private Sub AT_AddSyntax(ByVal registry As ATRegistry, ByVal idValue As String, ByVal syntaxValue As String, ByVal argumentTypes As String, ByVal returnType As String)
    Dim primitive As ATPrimitive
    Set primitive = AT_MakePrimitive(idValue, "OPERATOR", argumentTypes, returnType, "expression " & syntaxValue & " expression")
    primitive.SetSyntax syntaxValue
    registry.Register primitive
End Sub

