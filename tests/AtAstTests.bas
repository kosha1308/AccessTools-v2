Attribute VB_Name = "AtAstTests"
Option Explicit

Private passed As Long, failed As Long, log As String

Public Function AtRunAstTests(ByVal repo As String, ByVal scratch As String) As String
    On Error GoTo Broken
    passed = 0: failed = 0: log = ""
    TestFixtures repo, scratch
    TestNodes repo
    TestScope repo
    TestLiterals
    Dim fso As Object, file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(scratch & "\ast-assertions.txt", True, True): file.Write log: file.Close
    AtRunAstTests = IIf(failed = 0, "PASS|AST|", "FAIL|AST|") & passed & "|" & failed
    If failed > 0 Then AtRunAstTests = AtRunAstTests & "|" & log
    Exit Function
Broken:
    AtRunAstTests = "FAIL|AST|" & Err.Number & "|" & Err.Description
End Function

Private Sub Check(ByVal condition As Boolean, ByVal label As String)
    If condition Then passed = passed + 1 Else failed = failed + 1
    log = log & IIf(condition, "PASS ", "FAIL ") & label & vbCrLf
End Sub

Private Function Catalog() As Object
    Dim cat As Object, args As New Collection, s As Object
    Set cat = AtDictionary(): cat.Add "OP_TEST@1.0.0", AtSignature("Bool", args)
    Set args = New Collection: args.Add "Text": args.Add "Text"
    Set s = AtSignature("Bool", args): s("Parameters").Add "MODE", "Text"
    cat.Add "OP_TEST_EQ@1.0.0", s
    Set args = New Collection: args.Add "Int64": args.Add "Int64"
    cat.Add "OP_TEST_LE@1.0.0", AtSignature("Bool", args)
    Set Catalog = cat
End Function

Private Function LoadSimple(ByVal repo As String, ByVal cat As Object) As CAtPackage
    Dim issues As Collection
    Set LoadSimple = AtLoad(repo & "\tests\fixtures\simple-ast.xml", cat, issues)
    If LoadSimple Is Nothing Or AtHasErrors(issues) Then Err.Raise 5, , "Positive AST fixture failed."
End Function

Private Function Definition(ByVal p As CAtPackage, ByVal id As String) As CAtDefinition
    Dim d As CAtDefinition
    For Each d In p.Definitions
        If d.Id = id Then Set Definition = d: Exit Function
    Next d
    Err.Raise 5, , "Missing definition " & id
End Function

Private Function Node(ByVal p As CAtPackage, ByVal id As String) As CAtNode
    Dim n As CAtNode
    For Each n In Definition(p, "EXPR_REQUIRED").Nodes
        If n.Id = id Then Set Node = n: Exit Function
    Next n
    Err.Raise 5, , "Missing node " & id
End Function

Private Function Has(ByVal issues As Collection, ByVal code As String, Optional ByVal nodeId As String = "") As Boolean
    Dim issue As Variant
    For Each issue In issues
        If issue("Code") = code And (nodeId = "" Or issue("NodeId") = nodeId) Then
            Has = issue("ObjectId") <> "": Exit Function
        End If
    Next issue
End Function

Private Sub Expect(ByVal p As CAtPackage, ByVal cat As Object, ByVal code As String, ByVal nodeId As String, ByVal label As String)
    Dim issues As Collection: Set issues = AtValidate(p, cat)
    Check Has(issues, code, nodeId), label
End Sub

Private Sub TestFixtures(ByVal repo As String, ByVal scratch As String)
    Dim cat As Object, p As CAtPackage, q As CAtPackage, issues As Collection, fixture As Variant, xml As String
    Set cat = Catalog()
    Set p = LoadSimple(repo, cat)
    Node(p, "CALL_EQ").Parameters.Add "EXTRA", AtTypedParameter("Text", "x")
    Expect p, cat, "AST_PARAMETER", "CALL_EQ", "unexpected operation parameter"
    Node(p, "CALL_EQ").Parameters.Remove "EXTRA"
    cat("OP_TEST_EQ@1.0.0")("Arguments").Remove 2
    Expect p, cat, "AST_ARITY", "CALL_EQ", "operation arity from trusted catalog"
    Set cat = Catalog(): cat("OP_TEST_EQ@1.0.0")("Arguments").Remove 1: cat("OP_TEST_EQ@1.0.0")("Arguments").Add "Bool", , 1
    Expect p, cat, "AST_TYPE", "VALUE_TEXT", "operation argument types exact"
    Set cat = Catalog(): cat("OP_TEST_EQ@1.0.0")("ReturnType") = "Int64"
    Expect p, cat, "AST_TYPE", "CALL_EQ", "operation return type exact"
    Set cat = Catalog(): cat("OP_TEST_EQ@1.0.0")("Propagation") = "ErrorToFalse"
    Expect p, cat, "AST_SIGNATURE", "CALL_EQ", "technical error coercion forbidden in signature"
    Set cat = Catalog()
    For Each fixture In Array("simple", "equipment-return")
        Set p = AtLoad(repo & "\tests\fixtures\" & fixture & "-ast.xml", cat, issues)
        Check Not p Is Nothing, "load AST " & fixture
        Check Not AtHasErrors(issues), "type check AST " & fixture
        If p Is Nothing Then Err.Raise 5, , "AST file load failed."
        Check Not Has(issues, "AST_DEFERRED"), "v2 is checked structurally"
        xml = AtToXml(p)
        Check AtSaveNew(p, scratch & "\" & fixture & "-ast.xml", cat, issues), "save AST " & fixture
        Set q = AtLoad(scratch & "\" & fixture & "-ast.xml", cat, issues)
        Check AtToXml(q) = xml, "AST full roundtrip " & fixture
        Set q = AtLoad(repo & "\tests\fixtures\" & fixture & "-ast-invalid.xml", cat, issues)
        Check Not q Is Nothing And AtHasErrors(issues), "negative fixture " & fixture
        If fixture = "simple" Then Check Has(issues, "AST_TYPE", "STATE_NODE"), "negative type NodeId"
        If fixture = "equipment-return" Then Check Has(issues, "AST_SCOPE", "LEFT"), "negative scope NodeId"
    Next fixture
    Set p = LoadSimple(repo, cat)
    Dim text As String
    text = ChrW(&H418) & ChrW(&H43C) & ChrW(&H44F) & " <&" & vbCrLf & vbTab & "%0D"
    Node(p, "LITERAL_TEXT").Values("Value") = text
    Check AtSaveNew(p, scratch & "\unicode-ast.xml", cat, issues), "AST UTF8 file save"
    Set q = AtLoad(scratch & "\unicode-ast.xml", cat, issues)
    Check Node(q, "LITERAL_TEXT").Values("Value") = text, "AST Unicode CR LF TAB percent roundtrip"
    xml = AtToXml(p)
    Definition(p, "LOC_NAME").Values("Address") = "D12": Definition(p, "FIELD_NAME").DisplayName = "Renamed"
    Set issues = AtValidate(p, cat): Check Not AtHasErrors(issues), "AST unaffected by coordinate move"
    Check Node(p, "FIELD_VALUE").Values("FieldId") = "FIELD_NAME" And Definition(p, "RULE_REQUIRED").Id = "RULE_REQUIRED", "stable FieldId RuleId NodeId"
    Set q = AtFromXml(Replace(xml, "schemaVersion=""2.0.0""", "schemaVersion=""1.0.0"""), cat, issues)
    Check q Is Nothing And Has(issues, "XML_SHAPE"), "no silent v2 as legacy"
    p.SchemaVersion = "1.0.0": Expect p, cat, "AST_SCHEMA", "ROOT", "memory AST forbidden in legacy"
    p.SchemaVersion = "2.0.0"
    Set q = AtFromXml(Replace(xml, "schemaVersion=""2.0.0""", "schemaVersion=""3.0.0"""), cat, issues)
    Check q Is Nothing And Has(issues, "SCHEMA_VERSION"), "unknown AST schema"
    Set q = AtFromXml("<!DOCTYPE package [<!ENTITY x SYSTEM 'file:///C:/not-read'>]>" & xml, cat, issues)
    Check q Is Nothing And Has(issues, "XML_INVALID"), "AST DTD rejected"
    Set q = AtFromXml(Replace(xml, "<values>", "<values><unknown/>"), cat, issues)
    Check q Is Nothing And Has(issues, "XML_SHAPE", "FIELD_VALUE"), "AST closed shape NodeId"
    Set q = AtFromXml(Replace(xml, "<parameters>", "<parameters><parameter name=""MODE"" type=""Text"" value=""x""/>"), cat, issues)
    Check q Is Nothing And Has(issues, "XML_SHAPE", "CALL_EQ"), "duplicate AST parameter rejected"
End Sub

Private Sub TestNodes(ByVal repo As String)
    Dim cat As Object, p As CAtPackage, n As CAtNode, d As CAtDefinition, issues As Collection
    Set cat = Catalog(): Set p = LoadSimple(repo, cat)
    Node(p, "FIELD_VALUE").Values("FieldId") = "MISSING": Expect p, cat, "AST_FIELD", "FIELD_VALUE", "missing AST field"
    Node(p, "FIELD_VALUE").Values("FieldId") = "FIELD_NAME"
    Node(p, "FIELD_VALUE").Values("FieldVersion") = "2.0.0": Expect p, cat, "AST_VERSION", "FIELD_VALUE", "field version pinned"
    Node(p, "FIELD_VALUE").Values("FieldVersion") = "1.0.0"
    Node(p, "VALUE_BOOL").ResultType = "Text": Expect p, cat, "AST_TYPE", "VALUE_BOOL", "no Text to Bool coercion"
    Node(p, "VALUE_BOOL").ResultType = "Bool"
    Node(p, "CALL_EQ").Values("OperationVersion") = "2.0.0": Expect p, cat, "AST_OPERATION", "CALL_EQ", "unknown operation version"
    Node(p, "CALL_EQ").Values("OperationVersion") = "1.0.0"
    Node(p, "CALL_EQ").Values("OperationId") = "ABSENT": Expect p, cat, "AST_OPERATION", "CALL_EQ", "unknown operation ID"
    Node(p, "CALL_EQ").Values("OperationId") = "OP_TEST_EQ"
    Node(p, "CALL_EQ").Parameters("MODE")("Type") = "Bool": Expect p, cat, "AST_PARAMETER", "CALL_EQ", "parameter type checked"
    Node(p, "CALL_EQ").Parameters("MODE")("Type") = "Text"
    Node(p, "CALL_EQ").Parameters.Remove "MODE": Expect p, cat, "AST_PARAMETER", "CALL_EQ", "required parameter"
    Node(p, "CALL_EQ").Parameters.Add "MODE", AtTypedParameter("Text", "binary")
    cat.Remove "OP_TEST_EQ@1.0.0": cat.Add "OP_TEST_EQ@1.0.0", "Bool"
    Expect p, cat, "AST_SIGNATURE", "CALL_EQ", "legacy catalog not an AST signature"
    Set cat = Catalog(): cat("OP_TEST_EQ@1.0.0")("States").Add "FALSE"
    Expect p, cat, "AST_SIGNATURE", "CALL_EQ", "unknown operation state"
    Set cat = Catalog()
    Node(p, "STATE_NODE").Values("States") = "ERROR|LIMIT|UNAVAILABLE|CONFLICT|MISSING|INVALID|NOT_APPLICABLE|EMPTY|OK"
    Set issues = AtValidate(p, cat): Check Not AtHasErrors(issues), "all nine states valid contract"
    Node(p, "STATE_NODE").Values("States") = "FALSE": Expect p, cat, "AST_STATE", "STATE_NODE", "FALSE is not EvalState"
    Node(p, "STATE_NODE").Values("States") = "OK"
    Node(p, "LITERAL_TRUE").Values("State") = "ERROR": Expect p, cat, "AST_LITERAL", "LITERAL_TRUE", "technical error not a Bool payload"
    Node(p, "LITERAL_TRUE").Values("State") = "OK"
    Node(p, "NOT_NODE").Children.Remove 1: Expect p, cat, "AST_TREE", "LITERAL_FALSE", "orphan node rejected"
    Definition(p, "EXPR_REQUIRED").Nodes.Remove 12
    Expect p, cat, "AST_ARITY", "NOT_NODE", "Not arity checked"
    Set p = LoadSimple(repo, cat): Node(p, "NOT_NODE").Children.Add "ROOT"
    Expect p, cat, "AST_CYCLE", "ROOT", "cyclic nodes rejected"
    Set p = LoadSimple(repo, cat): Node(p, "NOT_NODE").Children.Add "NO_NODE"
    Expect p, cat, "AST_REFERENCE", "NOT_NODE", "dangling child NodeId"
    Set p = LoadSimple(repo, cat): Set n = AtNewNode("ROOT", "TypedLiteral", "Bool"): Definition(p, "EXPR_REQUIRED").Nodes.Add n
    Expect p, cat, "AST_DUPLICATE", "ROOT", "duplicate NodeId"
    Set p = LoadSimple(repo, cat): Node(p, "NOT_NODE").Kind = "EvaluateText"
    Expect p, cat, "AST_KIND", "NOT_NODE", "unknown node kind rejected"
    Set p = LoadSimple(repo, cat): Node(p, "NOT_NODE").ResultType = "Boolish"
    Expect p, cat, "AST_TYPE", "NOT_NODE", "unknown result type"
    Set p = LoadSimple(repo, cat)
    Dim index As Object, item As CAtDefinition
    Set index = AtDictionary()
    For Each item In p.Definitions: index.Add item.Kind & "/" & item.Id, item: Next item
    Set issues = New Collection: AtValidateAst p, index, cat, issues, 10, 64
    Check Has(issues, "AST_LIMIT", "ROOT"), "host node budget"
    Set issues = New Collection: AtValidateAst p, index, cat, issues, 512, 2
    Check Has(issues, "AST_LIMIT"), "host depth budget"
    Set issues = New Collection: AtValidateAst p, index, cat, issues, 513, 64
    Check Has(issues, "AST_BUDGET"), "host cannot raise safety budget"
    ' Hard limits, independent of XML recursion and host-requested lower limits.
    Set d = Definition(p, "EXPR_REQUIRED")
    Dim i As Long
    For i = d.Nodes.Count + 1 To 513
        Set n = AtNewNode("N_" & i, "TypedLiteral", "Bool"): d.Nodes.Add n
    Next i
    Expect p, cat, "AST_LIMIT", "ROOT", "hard 512 node limit"
    Set p = LoadSimple(repo, cat): Set d = Definition(p, "EXPR_REQUIRED")
    Set d.Nodes = New Collection: d.RootNodeId = "DEPTH_1"
    For i = 1 To 65
        Set n = AtNewNode("DEPTH_" & i, IIf(i = 65, "TypedLiteral", "Not"), "Bool")
        If i < 65 Then n.Children.Add "DEPTH_" & (i + 1)
        If i = 65 Then n.Values.Add "State", "OK": n.Values.Add "Value", "true"
        d.Nodes.Add n
    Next i
    Expect p, cat, "AST_LIMIT", "DEPTH_65", "hard depth 64 limit"
    Set p = LoadSimple(repo, cat)
    For i = 1 To 9
        Set d = AtNewDefinition("Expression", "EXTRA_" & i, "Synthetic budget fixture")
        d.RootNodeId = "N_1": d.ScopeTemplateId = "SIMPLE"
        d.Values.Add "Deferred", "false": d.Values.Add "ReturnType", "Bool"
        Dim j As Long
        For j = 1 To 512
            Set n = AtNewNode("N_" & j, IIf(j = 1, "AllOf", "TypedLiteral"), "Bool")
            If j = 1 Then
                Dim k As Long
                For k = 2 To 512: n.Children.Add "N_" & k: Next k
            Else
                n.Values.Add "State", "OK": n.Values.Add "Value", "true"
            End If
            d.Nodes.Add n
        Next j
        p.Definitions.Add d
    Next i
    Expect p, cat, "AST_LIMIT", "N_1", "hard package 4096 node limit"
End Sub

Private Sub TestScope(ByVal repo As String)
    Dim p As CAtPackage, cat As Object, issues As Collection, inner As CAtNode, value As CAtNode
    Set cat = Catalog(): Set p = LoadSimple(repo, cat)
    Node(p, "VALUE_TEXT").Values("BindingId") = "B_BOOL": Expect p, cat, "AST_SCOPE", "VALUE_TEXT", "Let forward self reference rejected"
    Node(p, "VALUE_TEXT").Values("BindingId") = "B_TEXT"
    Node(p, "VALUE_BOOL").Values("BindingId") = "OUTSIDE": Expect p, cat, "AST_SCOPE", "VALUE_BOOL", "missing binding"
    Node(p, "VALUE_BOOL").Values("BindingId") = "B_BOOL"
    Set inner = AtNewNode("INNER", "Let", "Bool"): inner.Bindings.Add "B_INNER"
    inner.Children.Add "INNER_VALUE": inner.Children.Add "THEN_NODE"
    Set value = AtNewNode("INNER_VALUE", "TypedLiteral", "Bool"): value.Values.Add "State", "OK": value.Values.Add "Value", "true"
    Definition(p, "EXPR_REQUIRED").Nodes.Add inner: Definition(p, "EXPR_REQUIRED").Nodes.Add value
    Node(p, "WHEN_ROOT").Children.Remove 2: Node(p, "WHEN_ROOT").Children.Add "INNER", , 2
    Node(p, "VALUE_BOOL").Values("BindingId") = "B_INNER"
    Set issues = AtValidate(p, cat): Check Not AtHasErrors(issues), "nested Let lexical binding valid"
    Node(p, "LITERAL_TRUE").Kind = "ValueRef": Node(p, "LITERAL_TRUE").Values.RemoveAll
    Node(p, "LITERAL_TRUE").Values.Add "BindingId", "B_INNER"
    Expect p, cat, "AST_SCOPE", "LITERAL_TRUE", "nested binding cannot leak to sibling"
    inner.Bindings.Remove 1: inner.Bindings.Add "B_BOOL"
    Expect p, cat, "AST_BINDING", "INNER", "stable BindingId cannot be reused"
    Set p = LoadSimple(repo, cat)
    Definition(p, "EXPR_REQUIRED").ScopeTemplateId = "ABSENT"
    Expect p, cat, "AST_SCOPE", "ROOT", "template context required"
End Sub

Private Sub TestLiterals()
    Check AtLiteralValid("Int64", "9223372036854775807"), "Int64 max without Double rounding"
    Check AtLiteralValid("Int64", "-9223372036854775808"), "Int64 min"
    Check Not AtLiteralValid("Int64", "9223372036854775808"), "Int64 overflow"
    Check Not AtLiteralValid("Int64", "01"), "Int64 canonical spelling"
    Check Not AtLiteralValid("Bool", "TRUE"), "Bool exact payload"
    Check AtLiteralValid("Decimal", "123.45"), "Decimal canonical"
    Check AtLiteralValid("Decimal", "0.1234567890123456789012345678"), "Decimal 28 significant fractional digits"
    Check AtLiteralValid("Decimal", "0.0000000000000000000000000001"), "Decimal leading zeros not precision"
    Check Not AtLiteralValid("Decimal", "12345678901234567890123456789"), "Decimal precision bound"
    Check Not AtLiteralValid("Decimal", "0.00000000000000000000000000001"), "Decimal scale bound"
    Check Not AtLiteralValid("Decimal", "1e3"), "no implicit exponential number"
    Check AtLiteralValid("Date", "2024-02-29"), "valid calendar date"
    Check Not AtLiteralValid("Date", "2026-02-29"), "invalid calendar date"
    Check AtLiteralValid("DateTime", "2026-10-19T12:34:56Z"), "UTC DateTime"
    Check Not AtLiteralValid("DateTime", "2026-10-19T25:34:56Z"), "invalid time"
    Check Not AtLiteralValid("IpSet", "10.0.0.1"), "no implicit Text to IP literal codec"
End Sub
