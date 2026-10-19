Attribute VB_Name = "AtEvalTests"
Option Explicit

Private passed As Long, failed As Long, transcript As String

Public Function AtRunEvalTests(ByVal repo As String, ByVal scratch As String) As String
    On Error GoTo Broken
    passed = 0: failed = 0: transcript = ""
    TestExamples repo
    TestControl repo
    TestRequirements repo
    TestLet repo
    TestBudgets repo
    TestRegistry
    TestSideEffects repo
    Dim fso As Object, file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(scratch & "\eval-assertions.txt", True, True): file.Write transcript: file.Close
    AtRunEvalTests = IIf(failed = 0, "PASS|Evaluator|", "FAIL|Evaluator|") & passed & "|" & failed
    If failed > 0 Then AtRunEvalTests = AtRunEvalTests & "|" & transcript
    Exit Function
Broken:
    AtRunEvalTests = "FAIL|Evaluator|" & Err.Number & "|" & Err.Description
End Function

Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    If ok Then passed = passed + 1 Else failed = failed + 1
    transcript = transcript & IIf(ok, "PASS ", "FAIL ") & label & vbCrLf
End Sub

Private Sub TestSideEffects(ByVal repo As String)
    Dim p As CAtPackage, n As CAtNode, r As CAtEvalResult, cell As Object
    Dim saved As Boolean, active As String, events As Boolean, screen As Boolean, calculation As Long
    Set p = LoadExample(repo): Set n = Root(p, "TypedLiteral")
    n.Values.Add "State", "OK": n.Values.Add "Value", "true"
    Set cell = ThisWorkbook.Worksheets(1).Range("A1")
    cell.NumberFormat = "@": cell.Value2 = "Evaluator sentinel": cell.Interior.Color = RGB(17, 34, 51)
    cell.AddComment "User annotation sentinel"
    saved = ThisWorkbook.Saved: active = Application.ActiveWorkbook.Name
    events = Application.EnableEvents: screen = Application.ScreenUpdating: calculation = Application.Calculation
    Set r = Run(p)
    Check r.IsComplete And r.Value = "true", "side effect test expression executes"
    Check cell.Value2 = "Evaluator sentinel" And cell.NumberFormat = "@" And cell.Interior.Color = RGB(17, 34, 51), "evaluation leaves worksheet data and formatting unchanged"
    Check cell.Comment.Text = "User annotation sentinel" And ThisWorkbook.Saved = saved, "evaluation preserves comment and does not save workbook"
    Check Application.ActiveWorkbook.Name = active And Application.EnableEvents = events And Application.ScreenUpdating = screen And Application.Calculation = calculation, "evaluation leaves workbook selection and application flags unchanged"
End Sub

Private Function LoadExample(ByVal repo As String, Optional ByVal equipment As Boolean = False) As CAtPackage
    Dim registry As New CAtOperationRegistry, cat As Object, args As New Collection, p As CAtPackage, issues As Collection
    Dim d As CAtDefinition, n As CAtNode, reference As Variant
    Set cat = registry.Catalog
    cat.Add "OP_TEST@1.0.0", AtSignature("Bool", args)
    cat.Add "OP_TEST_EQ@1.0.0", cat("OP_TEXT_EQUAL@1.0.0")
    cat.Add "OP_TEST_LE@1.0.0", cat("OP_INT64_LE@1.0.0")
    Set p = AtLoad(repo & "\tests\fixtures\" & IIf(equipment, "equipment-return", "simple") & "-ast.xml", cat, issues)
    If p Is Nothing Or AtHasErrors(issues) Then Err.Raise 5, , "Example load failed"
    For Each d In p.Definitions
        For Each reference In d.References
            If reference("Kind") = "Operation" Then reference("Id") = "OP_TEXT_EQUAL"
        Next reference
        For Each n In d.Nodes
            If n.Kind = "Call" Then n.Values("OperationId") = IIf(equipment, "OP_INT64_LE", "OP_TEXT_EQUAL")
        Next n
    Next d
    Set LoadExample = p
End Function

Private Function Expr(ByVal p As CAtPackage) As CAtDefinition
    Dim d As CAtDefinition
    For Each d In p.Definitions
        If d.Kind = "Expression" Then Set Expr = d: Exit Function
    Next d
End Function

Private Function Node(ByVal p As CAtPackage, ByVal id As String) As CAtNode
    Dim n As CAtNode
    For Each n In Expr(p).Nodes
        If n.Id = id Then Set Node = n: Exit Function
    Next n
    Err.Raise 5, , "Missing runtime test node"
End Function

Private Function Root(ByVal p As CAtPackage, ByVal kind As String) As CAtNode
    Dim n As CAtNode
    Set Expr(p).Nodes = New Collection: Expr(p).RootNodeId = "ROOT"
    Set n = AtNewNode("ROOT", kind, "Bool"): Expr(p).Nodes.Add n: Set Root = n
End Function

Private Function Literal(ByVal p As CAtPackage, ByVal id As String, ByVal typeName As String, ByVal value As String) As CAtNode
    Dim n As CAtNode: Set n = AtNewNode(id, "TypedLiteral", typeName)
    n.Values.Add "State", "OK": n.Values.Add "Value", value: Expr(p).Nodes.Add n
    Set Literal = n
End Function

Private Function Field(ByVal p As CAtPackage, ByVal id As String, Optional ByVal fieldId As String = "FIELD_NAME") As CAtNode
    Dim n As CAtNode: Set n = AtNewNode(id, "FieldRef", "Text")
    n.Values.Add "FieldId", fieldId: n.Values.Add "FieldVersion", "1.0.0": Expr(p).Nodes.Add n
    Set Field = n
End Function

Private Function StateNode(ByVal p As CAtPackage, ByVal id As String, ByVal states As String, ByVal childId As String) As CAtNode
    Dim n As CAtNode: Set n = AtNewNode(id, "StateIs", "Bool")
    n.Values.Add "States", states: n.Children.Add childId: Expr(p).Nodes.Add n: Set StateNode = n
End Function

Private Function RefValue(ByVal p As CAtPackage, ByVal id As String, ByVal binding As String) As CAtNode
    Dim n As CAtNode: Set n = AtNewNode(id, "ValueRef", "Text")
    n.Values.Add "BindingId", binding: Expr(p).Nodes.Add n: Set RefValue = n
End Function

Private Function CallNode(ByVal p As CAtPackage, ByVal id As String, ByVal operationId As String, ByVal typeName As String) As CAtNode
    Dim n As CAtNode: Set n = AtNewNode(id, "Call", typeName)
    n.Values.Add "OperationId", operationId: n.Values.Add "OperationVersion", "1.0.0"
    If operationId = "OP_TEXT_EQUAL" Then n.Parameters.Add "MODE", AtTypedParameter("Text", "binary")
    Expr(p).Nodes.Add n: Set CallNode = n
End Function

Private Function Context(ByVal p As CAtPackage, Optional ByVal fields As Object = Nothing) As CAtEvalContext
    If fields Is Nothing Then Set fields = AtDictionary()
    Set Context = AtEvalContext(p.Id, p.Version, Expr(p).ScopeTemplateId, fields)
End Function

Private Function Run(ByVal p As CAtPackage, Optional ByVal fields As Object = Nothing) As CAtEvalResult
    Dim evaluator As New CAtEvaluator
    Set Run = evaluator.Evaluate(p, Expr(p).Id, Context(p, fields))
End Function

Private Function Events(ByVal r As CAtEvalResult, ByVal id As String, ByVal eventName As String) As Long
    Dim item As Variant
    For Each item In r.Trace
        If item("NodeId") = id And item("Event") = eventName Then Events = Events + 1
    Next item
End Function

Private Sub TestExamples(ByVal repo As String)
    Dim p As CAtPackage, r As CAtEvalResult, fields As Object, n As CAtNode, before As String, equipment As Variant
    For Each equipment In Array(False, True)
        Set p = LoadExample(repo, CBool(equipment)): Set fields = AtDictionary()
        If Not equipment Then fields.Add "FIELD_NAME", AtEvalValue("Text", "OK", Node(p, "LITERAL_TEXT").Values("Value"))
        before = AtToXml(p): Set r = Run(p, fields)
        Check r.State = "OK" And r.Value = "true" And r.IsComplete, "original typed AST executes for template " & Expr(p).ScopeTemplateId
        Check AtToXml(p) = before, "evaluator leaves model and XML unchanged"
        ' Identical AST and NodeIds in both independent template contexts.
        Set n = Root(p, "Call"): n.Values.Add "OperationId", "OP_TEXT_EQUAL": n.Values.Add "OperationVersion", "1.0.0"
        n.Parameters.Add "MODE", AtTypedParameter("Text", "binary"): n.Children.Add "A": n.Children.Add "B"
        Literal p, "A", "Text", "x"
        Set n = CallNode(p, "B", "OP_TEXT_TRIM", "Text"): n.Children.Add "C"
        Literal p, "C", "Text", " x "
        Set r = Run(p): Check r.State = "OK" And r.Value = "true" And r.IsComplete, "identical rule for template " & Expr(p).ScopeTemplateId
    Next equipment
    Set p = LoadExample(repo): Set fields = AtDictionary()
    Dim original As CAtEvalResult, snapshot As CAtEvalContext, evaluator As New CAtEvaluator, returned As CAtEvalResult
    Set original = AtEvalValue("Text", "OK", Node(p, "LITERAL_TEXT").Values("Value"), "Source/CATALOG")
    fields.Add "FIELD_NAME", original: Set snapshot = Context(p, fields)
    original.SetData "Text", "OK", "mutated": fields.Remove "FIELD_NAME"
    Set returned = snapshot.ReadField("FIELD_NAME"): returned.SetData "Text", "OK", "also mutated"
    Set r = evaluator.Evaluate(p, Expr(p).Id, snapshot)
    Check r.Value = "true" And r.IsComplete, "context snapshots caller and returned envelope mutations"
    Check r.Provenance.Count > 0, "result carries node field source provenance"
    Dim records As Collection: Set records = r.Provenance: records(1)("Id") = "modified"
    Check r.Provenance(1)("Id") <> "modified", "returned record collections are defensive copies"
    Check ThisWorkbook.Worksheets.Count > 0, "own test workbook remains open"
    Set r = AtEvalValue("Bool", "OK", "true")
    r.AddDiagnostic "E_SOURCE", "Context", "", "Source failed."
    Check Not r.IsComplete, "error diagnostic alone cannot imply complete success"
    Check RejectInconsistentContext(p), "immutable context rejects OK input with error diagnostics"
End Sub

Private Function RejectInconsistentContext(ByVal p As CAtPackage) As Boolean
    On Error GoTo Rejected
    Dim fields As Object, r As CAtEvalResult, c As CAtEvalContext
    Set fields = AtDictionary(): Set r = AtEvalValue("Text", "OK", "x")
    r.AddDiagnostic "E_SOURCE", "Context", "", "Source failed."
    fields.Add "FIELD_NAME", r: Set c = Context(p, fields)
    Exit Function
Rejected:
    RejectInconsistentContext = Err.Number = 5
End Function

Private Sub TestControl(ByVal repo As String)
    Dim p As CAtPackage, n As CAtNode, r As CAtEvalResult, kind As Variant
    For Each kind In Array("AllOf", "AnyOf")
        Set p = LoadExample(repo): Set n = Root(p, CStr(kind)): n.Children.Add "LEFT": n.Children.Add "STATE"
        Literal p, "LEFT", "Bool", IIf(kind = "AllOf", "false", "true")
        StateNode p, "STATE", "OK", "FIELD": Field p, "FIELD"
        Set r = Run(p)
        Check r.State = "OK" And r.Value = IIf(kind = "AllOf", "false", "true") And r.IsComplete, "short-circuit " & kind
        Check r.Requirements.Count = 0 And Events(r, "FIELD", "NOT_EVALUATED") = 1 And Events(r, "FIELD", "EXECUTED") = 0, "skipped required input creates no requirement " & kind
    Next kind
    Set p = LoadExample(repo): Set n = Root(p, "When"): n.Children.Add "GATE": n.Children.Add "YES": n.Children.Add "NO"
    Set n = CallNode(p, "GATE", "OP_TEXT_EQUAL", "Bool"): n.Children.Add "FIELD": n.Children.Add "TEXT"
    Field p, "FIELD": Literal p, "TEXT", "Text", "x": Literal p, "YES", "Bool", "true": Literal p, "NO", "Bool", "false"
    Set r = Run(p)
    Check r.State = "MISSING" And r.Value = "" And Not r.IsComplete, "non-OK gate does not select else"
    Check Events(r, "YES", "NOT_EVALUATED") = 1 And Events(r, "NO", "NOT_EVALUATED") = 1, "both When branches skipped for non-OK gate"
    Check Events(r, "TEXT", "EXECUTED") = 1 And Events(r, "GATE", "OPERATION") = 0, "Call arguments eager ordered but algorithm not invoked for non-OK"
    Set n = Root(p, "Not"): n.Children.Add "FALSE": Literal p, "FALSE", "Bool", "false"
    Set r = Run(p): Check r.State = "OK" And r.Value = "true", "Not false"
End Sub

Private Sub TestRequirements(ByVal repo As String)
    Dim p As CAtPackage, n As CAtNode, r As CAtEvalResult, state As Variant, fields As Object
    For Each state In Array("EMPTY", "MISSING", "INVALID", "UNAVAILABLE", "ERROR", "CONFLICT", "LIMIT", "NOT_APPLICABLE")
        Set p = LoadExample(repo): Set fields = AtDictionary(): fields.Add "FIELD_NAME", AtEvalValue("Text", CStr(state), "", "Source/CATALOG")
        Set r = Run(p, fields)
        Check r.State = state And Not r.HasValue And r.Value = "" And Not r.IsComplete, "strict state propagation " & state
        Set n = Root(p, "StateIs"): n.Values.Add "States", CStr(state): n.Children.Add "FIELD": Field p, "FIELD"
        Set r = Run(p, fields)
        Check r.State = "OK" And r.HasValue And r.Value = "true" And Not r.IsComplete, "StateIs observes but cannot hide " & state
        Check r.Requirements.Count = 1 And r.Requirements(1)("Disposition") = "UNRESOLVED" And r.Requirements(1)("OriginNodeId") = "FIELD", "requirement origin preserved " & state
        Check r.Diagnostics.Count > 0, "raw non-OK diagnostic retained " & state
    Next state
    Set p = LoadExample(repo): Set n = Root(p, "When"): n.Children.Add "GATE": n.Children.Add "YES": n.Children.Add "NO"
    StateNode p, "GATE", "UNAVAILABLE", "FIELD": Field p, "FIELD"
    Literal p, "YES", "Bool", "true": Literal p, "NO", "Bool", "false"
    Set fields = AtDictionary(): fields.Add "FIELD_NAME", AtEvalValue("Text", "UNAVAILABLE", "", "Source/CATALOG")
    Set r = Run(p, fields)
    Check r.State = "OK" And r.Value = "true" And Not r.IsComplete And r.Requirements.Count = 1, "When on StateIs cannot declare successful completeness"
    Dim requirements As Collection: Set requirements = r.Requirements
    requirements(1)("Disposition") = "RESOLVED": requirements.Remove 1
    Check r.Requirements.Count = 1 And Not r.IsComplete And r.Requirements(1)("Disposition") = "UNRESOLVED", "caller cannot erase required technical failure through returned collection"
    Set p = LoadExample(repo): Set n = Root(p, "Call")
    n.Values.Add "OperationId", "OP_TEXT_EQUAL": n.Values.Add "OperationVersion", "1.0.0"
    n.Parameters.Add "MODE", AtTypedParameter("Text", "unsafe")
    n.Children.Add "PRODUCER": n.Children.Add "TEXT"
    Set n = AtNewNode("PRODUCER", "Let", "Text"): n.Bindings.Add "UNUSED": n.Children.Add "FIELD": n.Children.Add "BODY": Expr(p).Nodes.Add n
    Field p, "FIELD": Literal p, "BODY", "Text", "x": Literal p, "TEXT", "Text", "x"
    Set r = Run(p, fields)
    Check r.State = "INVALID" And r.Requirements.Count = 2 And Not r.IsComplete, "own operation failure retained beside inherited requirement"
    Set p = LoadExample(repo): Set fields = AtDictionary(): fields.Add "FIELD_NAME", AtEvalValue("Bool", "OK", "true")
    Set r = Run(p, fields): Check r.State = "ERROR" And r.Value = "" And Not r.IsComplete, "snapshot type mismatch not coerced"
End Sub

Private Sub TestLet(ByVal repo As String)
    Dim p As CAtPackage, n As CAtNode, r As CAtEvalResult
    Set p = LoadExample(repo): Set n = Root(p, "Let"): n.Bindings.Add "B": n.Children.Add "BOUND": n.Children.Add "BODY"
    Set n = CallNode(p, "BOUND", "OP_TEXT_TRIM", "Text"): n.Children.Add "TEXT": Literal p, "TEXT", "Text", " x "
    Set n = CallNode(p, "BODY", "OP_TEXT_EQUAL", "Bool"): n.Children.Add "V1": n.Children.Add "V2"
    RefValue p, "V1", "B": RefValue p, "V2", "B"
    Set r = Run(p)
    Check r.State = "OK" And r.Value = "true" And r.IsComplete, "Let bound value reused"
    Check Events(r, "BOUND", "OPERATION") = 1 And Events(r, "V1", "EXECUTED") = 1 And Events(r, "V2", "EXECUTED") = 1, "Let computes once for two ValueRefs"
    Node(p, "V1").Values("BindingId") = "ABSENT": Set r = Run(p)
    Check r.State = "ERROR" And r.Trace.Count = 0 And Not r.IsComplete, "out-of-scope binding rejected before execution"
    Set p = LoadExample(repo): Set n = Root(p, "Let"): n.Bindings.Add "UNUSED": n.Children.Add "FIELD": n.Children.Add "BODY"
    Field p, "FIELD": Literal p, "BODY", "Bool", "true": Set r = Run(p)
    Check r.State = "OK" And r.Value = "true" And Not r.IsComplete, "unused reached Let binding cannot discard requirement"
    Set p = LoadExample(repo): Set n = Root(p, "Let"): n.Bindings.Add "OUTER": n.Children.Add "A": n.Children.Add "INNER"
    Literal p, "A", "Text", "x"
    Set n = AtNewNode("INNER", "Let", "Bool"): n.Bindings.Add "INSIDE": n.Children.Add "OUTER_REF": n.Children.Add "BODY": Expr(p).Nodes.Add n
    RefValue p, "OUTER_REF", "OUTER"
    Set n = CallNode(p, "BODY", "OP_TEXT_EQUAL", "Bool"): n.Children.Add "V1": n.Children.Add "V2"
    RefValue p, "V1", "OUTER": RefValue p, "V2", "INSIDE"
    Set r = Run(p): Check r.Value = "true" And r.IsComplete, "nested Let sees outer binding and its own value"
    Set p = LoadExample(repo): Set n = Root(p, "Let"): n.Bindings.Add "EMPTY_CONST": n.Children.Add "A": n.Children.Add "BODY"
    Set n = Literal(p, "A", "Text", ""): n.Values("State") = "EMPTY"
    StateNode p, "BODY", "EMPTY", "V1": RefValue p, "V1", "EMPTY_CONST"
    Set r = Run(p): Check r.State = "OK" And r.Value = "true" And r.IsComplete And r.Requirements.Count = 0, "EMPTY constant keeps no external obligation through ValueRef"
End Sub

Private Sub TestBudgets(ByVal repo As String)
    Dim p As CAtPackage, n As CAtNode, r As CAtEvalResult, evaluator As New CAtEvaluator, c As CAtEvalContext
    Set p = LoadExample(repo): Set c = Context(p)
    Set r = evaluator.Evaluate(p, Expr(p).Id, c, 1)
    Check r.State = "LIMIT" And r.Value = "" And Not r.IsComplete, "step budget LIMIT"
    Set r = evaluator.Evaluate(p, Expr(p).Id, c, 1024, 1)
    Check r.State = "LIMIT", "depth budget LIMIT"
    Set r = evaluator.Evaluate(p, Expr(p).Id, c, 1024, 64, 128, 1)
    Check r.State = "LIMIT" And r.Trace.Count <= 1, "bounded trace LIMIT"
    Set n = Root(p, "Let"): n.Bindings.Add "A": n.Bindings.Add "B": n.Children.Add "F1": n.Children.Add "F2": n.Children.Add "BODY"
    Field p, "F1": Field p, "F2", "FIELD_NOTE": Literal p, "BODY", "Bool", "true"
    Set r = evaluator.Evaluate(p, Expr(p).Id, c, 1024, 64, 1)
    Check r.State = "LIMIT" And r.Diagnostics.Count <= 1 And r.Requirements.Count >= 2, "diagnostic cap retains unresolved requirements"
    Set p = LoadExample(repo): Set c = Context(p)
    Set r = evaluator.Evaluate(p, Expr(p).Id, c, 1024, 64, 128, 2048, 2)
    Check r.State = "LIMIT" And r.Value = "", "payload cap LIMIT"
    Set r = evaluator.Evaluate(p, Expr(p).Id, c, 1025)
    Check r.State = "ERROR" And r.Trace.Count = 0, "host cannot raise hard cap"
    Set r = evaluator.Evaluate(p, Expr(p).Id, AtEvalContext(p.Id, "2.0.0", Expr(p).ScopeTemplateId, AtDictionary()))
    Check r.State = "ERROR" And r.Trace.Count = 0, "snapshot identity pinned"
    Set p = LoadExample(repo): p.SchemaVersion = "1.0.0": Set r = Run(p)
    Check r.State = "ERROR" And r.Trace.Count = 0, "legacy deferred XML never executable"
End Sub

Private Sub TestRegistry()
    Dim registry As New CAtOperationRegistry, args As Collection, parameters As Object, r As CAtEvalResult, catalog As Object
    Set args = New Collection: args.Add AtEvalValue("Int64", "OK", "-9223372036854775808"): args.Add AtEvalValue("Int64", "OK", "9223372036854775807")
    Set r = registry.ExecuteOperation("OP_INT64_LE", "1.0.0", args, AtDictionary())
    Check r.State = "OK" And r.Value = "true", "Int64 comparison without Double rounding"
    Set r = registry.ExecuteOperation("OP_INT64_LE", "2.0.0", args, AtDictionary())
    Check r.State = "ERROR" And r.Value = "", "unknown exact runtime operation version"
    Set args = New Collection: args.Add AtEvalValue("Text", "OK", "1"): args.Add AtEvalValue("Int64", "OK", "1")
    Set r = registry.ExecuteOperation("OP_INT64_LE", "1.0.0", args, AtDictionary())
    Check r.State = "ERROR" And r.Value = "", "runtime operation type check no Text coercion"
    Set args = New Collection: args.Add AtEvalValue("Text", "OK", "x"): args.Add AtEvalValue("Text", "OK", "x")
    Set parameters = AtDictionary(): parameters.Add "MODE", AtTypedParameter("Text", "unsafe")
    Set r = registry.ExecuteOperation("OP_TEXT_EQUAL", "1.0.0", args, parameters)
    Check r.State = "INVALID" And r.Value = "", "invalid operation parameter value"
    Set catalog = registry.Catalog: catalog.RemoveAll
    parameters("MODE")("Value") = "binary"
    Set r = registry.ExecuteOperation("OP_TEXT_EQUAL", "1.0.0", args, parameters)
    Check r.State = "OK" And r.Value = "true", "catalog mutation cannot register replace or remove runtime implementation"
    Set args = New Collection: args.Add AtEvalValue("Bool", "OK", "false"): args.Add AtEvalValue("Bool", "OK", "true")
    Set r = registry.ExecuteOperation("OP_BOOL_EQUAL", "1.0.0", args, AtDictionary())
    Check r.State = "OK" And r.Value = "false", "Bool equality returns typed false"
    Dim pair As Variant
    For Each pair In Array(Array("-10", "-2", "true"), Array("-2", "-10", "false"), Array("0", "0", "true"), Array("9223372036854775807", "9223372036854775806", "false"))
        Set args = New Collection: args.Add AtEvalValue("Int64", "OK", CStr(pair(0))): args.Add AtEvalValue("Int64", "OK", CStr(pair(1)))
        Set r = registry.ExecuteOperation("OP_INT64_LE", "1.0.0", args, AtDictionary())
        Check r.State = "OK" And r.Value = pair(2), "Int64 order " & pair(0) & " <= " & pair(1)
    Next pair
End Sub
