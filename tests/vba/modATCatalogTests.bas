Attribute VB_Name = "modATCatalogTests"
Option Explicit

Public Sub AT_TestCatalog()
    Dim context As New ATCompileContext, compiler As New ATCompiler
    Dim result As ATCompileResult, parser As New ATParser, serializer As New ATSerializer
    Dim parsed As ATParseResult, reparsed As ATParseResult
    Dim item As Variant, text As String
    Dim named As ATNamedExpression
    context.AddField "NAME", "TEXT"
    context.AddField "N", "NUMBER"
    context.AddField "PORTS", "PORT_SET"
    context.AddField "VALUES", "LIST<NUMBER>"
    context.AddField "WORDS", "LIST<TEXT>"
    context.AddField "FLAGS", "LIST<BOOL>"
    context.AddField "SET_WORDS", "SET<TEXT>"
    For Each item In Array("@SET_WORDS ANY [""a""]", "@SET_WORDS ANY: ""a""")
        Set result = compiler.Compile(CStr(item), context)
        AT_Assert Not result.Root Is Nothing, "set operator accepts canonical list literal"
    Next item
    For Each item In Array("@NAME CONTAINS ""abc""", "@NAME STARTS_WITH ""a""", "@NAME ENDS_WITH ""z""", "@NAME LIKE ""a*""", "@NAME MATCHES ""^[a-z]+$""", "@NAME EMPTY", "@NAME NOT_EMPTY", "@NAME IS NULL", "@NAME IS NOT NULL", "@NAME VALID", "@NAME INVALID", "@N BETWEEN 1 AND 10", "@N IN [1,2]", "@N NOT IN [3,4]", "@WORDS ANY [""a"",""b""]", "@WORDS COVER [""a""]", "@WORDS CONTAINS [""a""]", "UPPER(TRIM(@NAME)) = ""ABC""", "LEN(@NAME) > 3", "ABS(@N) > 0", "ROUND(@N,2) >= 0", "MIN(@N,2) <= MAX(@N,2)", "SUM(@VALUES) > AVG(@VALUES)", "MIN(@VALUES) <= MAX(@VALUES)", "COALESCE(@NAME,""fallback"") = ""abc""", "TODAY() >= DATE(""2026-01-01"")", "DATE_DIFF(DATE(""2026-01-01""),TODAY()) > 0")
        Set result = compiler.Compile(CStr(item), context)
        AT_Assert Not result.Root Is Nothing, "catalog accepts: " & CStr(item)
    Next item
    Set result = compiler.Compile("CONTAINS: ""abc""", context, "NAME")
    AT_Assert Not result.Root Is Nothing, "contextual scalar is not singleton list"
    Set result = compiler.Compile("EMPTY", context, "NAME")
    AT_Assert Not result.Root Is Nothing, "contextual unary emptiness"
    Set result = compiler.Compile("@FLAGS ANY: TRUE,FALSE", context)
    AT_Assert Not result.Root Is Nothing, "contextual Boolean list"
    For Each item In Array("@WORDS ANY: ""a""", "@FLAGS ANY: TRUE", "@N IN: 1")
        Set result = compiler.Compile(CStr(item), context)
        AT_Assert Not result.Root Is Nothing, "singleton contextual list preserves collection semantics"
    Next item
    Set named = New ATNamedExpression
    named.Initialize "HASPORT", "BOOL", "@PORTS ANY $P"
    named.AddParameter "P", "PORT_SET", "65536"
    context.AddNamed named
    For Each item In Array("$HASPORT(65536)", "$HASPORT(-1)", "$HASPORT([443,65536])", "$HASPORT()")
        Set result = compiler.Compile(CStr(item), context)
        AT_Assert result.Root Is Nothing, "named conversion validates port literals and defaults"
    Next item
    For Each item In Array("@N BETWEEN 1 AND 10 AND TRUE", "@N NOT IN [3,4]", "@NAME IS NOT NULL", "@NAME CONTAINS: ""abc""")
        Set parsed = parser.Parse(CStr(item))
        If Not parsed.Root Is Nothing Then
            text = serializer.Serialize(parsed.Root)
            Set reparsed = parser.Parse(text)
            AT_Assert reparsed.Diagnostics.Count = 0, "catalog round trip: " & CStr(item)
            If Not reparsed.Root Is Nothing Then AT_Assert text = serializer.Serialize(reparsed.Root), "catalog canonical stable"
        Else
            AT_Assert False, "catalog round trip initial parse"
        End If
    Next item
    Set parsed = parser.Parse("00012.3400")
    AT_Assert serializer.Serialize(parsed.Root) = "12.34", "decimal canonical form exact"
    Set parsed = parser.Parse("000.000")
    AT_Assert serializer.Serialize(parsed.Root) = "0", "zero canonical form"
    Set parsed = parser.Parse("123456789012345678901234567890.1200")
    AT_Assert serializer.Serialize(parsed.Root) = "123456789012345678901234567890.12", "canonical decimals do not use floating point"
    Set result = compiler.Compile("@PORTS COVER: 65536", context)
    AT_Assert result.Root Is Nothing, "invalid literal port upper bound rejected"
    Set result = compiler.Compile("@PORTS COVER: -1", context)
    AT_Assert result.Root Is Nothing, "negative literal port rejected"
    Set result = compiler.Compile("@PORTS COVER: 443.5", context)
    AT_Assert result.Root Is Nothing, "fractional literal port rejected"
End Sub
