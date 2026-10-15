Attribute VB_Name = "modATParserTests"
Option Explicit

Public Sub AT_TestParser()
    Dim parser As New ATParser
    Dim serializer As New ATSerializer
    Dim result As ATParseResult, other As ATParseResult
    Dim node As ATAstNode
    Dim text As String, item As Variant
    Set result = parser.Parse("1 + 2 * 3 > 4 AND NOT FALSE OR TRUE")
    AT_Assert result.Diagnostics.Count = 0, "parse precedence"
    Set node = result.Root
    AT_Assert node.Text = "OR" And node.Children(1).Text = "AND", "parse logical precedence"
    Set node = result.Root.Children(1).Children(1).Children(1)
    AT_Assert node.Text = "+" And node.Children(2).Text = "*", "parse arithmetic precedence"
    Set result = parser.Parse("COVER: 443,8443 AND NOT ANY: 9000", "TECH_DEST_PORT")
    Set other = parser.Parse("@TECH_DEST_PORT COVER: 443,8443 AND NOT @TECH_DEST_PORT ANY: 9000")
    AT_Assert result.Diagnostics.Count = 0 And other.Diagnostics.Count = 0, "parse contextual and full forms"
    AT_Assert serializer.Serialize(result.Root) = serializer.Serialize(other.Root), "parse contextual AST equivalence"
    text = "ALLROWS(@TECHNICAL_ROWS AS t, ANYROW(@ADDRESS_ROWS AS a, @a.IP INTERSECTS @t.IP))"
    Set result = parser.Parse(text)
    AT_Assert result.Diagnostics.Count = 0, "parse nested collection aliases"
    Set result = parser.Parse("$PORT_ALLOWED([443,8443])")
    AT_Assert result.Diagnostics.Count = 0 And result.Root.Children.Count = 1, "parse list as one named argument"
    For Each item In Array("@A = @B", "NOT NOT (@A = 2 OR @B >= 3)", "LOWER(""Text"") = ""text""", "@N / -2 + 3 % 2", "[1,2,3]", "DATE(""2026-10-15"")")
        Set result = parser.Parse(CStr(item))
        text = serializer.Serialize(result.Root)
        Set other = parser.Parse(text)
        AT_Assert other.Diagnostics.Count = 0, "round trip parses: " & CStr(item)
        AT_Assert text = serializer.Serialize(other.Root), "round trip stable: " & CStr(item)
    Next item
    For Each item In Array("1 2", "(1", "1 +", "[1,]", "COVER: 443", "@A = 2 ?", "EVAL(x)", "@A..B", "TRUE ""OR"" FALSE")
        Set result = parser.Parse(CStr(item))
        If CStr(item) = "EVAL(x)" Then
            AT_Assert result.Diagnostics.Count > 0, "parse bare unknown argument rejected"
        Else
            AT_Assert result.Diagnostics.Count > 0 And result.Root Is Nothing, "parse malformed input fails closed: " & CStr(item)
        End If
    Next item
End Sub
