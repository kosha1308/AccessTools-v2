Attribute VB_Name = "AtSchemaTests"
Option Explicit

Private passed As Long
Private failed As Long
Private failures As String
Private transcript As String

Public Function AtRunSchemaTests(ByVal repo As String, ByVal scratch As String) As String
    On Error GoTo Broken
    passed = 0: failed = 0: failures = "": transcript = ""
    Dim p As Object
    ' Bootstrap is late-bound so a missing foundation produces an observable RED.
    Set p = Application.Run("'" & ThisWorkbook.Name & "'!AtNewPackage", "SIMPLE", "Simple")
    TestModels repo, scratch
    TestValidationEdges repo
    TestNameRegressions repo
    WriteText scratch & "\assertions.txt", transcript
    AtRunSchemaTests = IIf(failed = 0, "PASS|", "FAIL|") & passed & "|" & failed & "|" & failures & "|" & HostBits()
    Exit Function
Broken:
    AtRunSchemaTests = "FAIL|" & passed & "|runtime|" & Err.Number & "|" & Err.Description
End Function

Private Function HostBits() As String
#If Win64 Then
    HostBits = "Office64"
#Else
    HostBits = "Office32"
#End If
End Function

Private Sub Check(ByVal ok As Boolean, ByVal label As String)
    transcript = transcript & IIf(ok, "PASS ", "FAIL ") & label & vbCrLf
    If ok Then
        passed = passed + 1
    Else
        failed = failed + 1: failures = failures & label & ";"
    End If
End Sub

Private Sub TestModels(ByVal repo As String, ByVal scratch As String)
    Dim p As CAtPackage, q As CAtPackage, d As Collection, cat As Object
    Dim x As CAtDefinition, original As String, xml As String, path As String, previous As String
    Set cat = AtDictionary(): cat.Add "OP_TEST@1.0.0", "Bool"
    Dim fixture As Variant
    For Each fixture In Array("simple", "equipment-return")
        Set p = AtLoad(repo & "\tests\fixtures\" & fixture & ".xml", cat, d)
        Check Not p Is Nothing, "load " & fixture
        Check Not AtHasErrors(d), "valid " & fixture
        If p Is Nothing Then Err.Raise 5, , "Fixture failed: " & fixture & " " & DiagnosticsText(d)
        original = AtToXml(p)
        path = scratch & "\" & fixture & ".xml"
        Check AtSaveNew(p, path, cat, d), "save " & fixture
        Set q = AtLoad(path, cat, d)
        Check Not q Is Nothing, "reload " & fixture
        Check AtToXml(q) = original, "roundtrip " & fixture
        previous = FileBytes(path)
        Check Not AtSaveNew(p, path, cat, d), "refuse overwrite"
        Check FileBytes(path) = previous, "previous bytes unchanged"
        Check HasCode(d, "FILE_EXISTS"), "overwrite diagnostic"
        Set q = AtLoad(path, cat, d)
        Check AtToXml(q) = original, "previous survives write failure"
    Next fixture
    Set p = AtLoad(repo & "\tests\fixtures\simple.xml", cat, d)
    Set x = Entity(p, "LOC_NAME")
    x.Values("Sheet") = "Bad/Sheet": Expect p, cat, "SHEET_INVALID", "invalid sheet selector"
    x.Values("Sheet") = "Request"
    x.Values("Address") = "$XFD$1048576"
    Entity(p, "FIELD_NAME").DisplayName = "Renamed"
    Set d = AtValidate(p, cat)
    Check Not AtHasErrors(d), "move and rename valid"
    Check Entity(p, "RULE_REQUIRED").References(3)("Id") = "FIELD_NAME", "stable field reference"
    x.Values("Address") = "XFE1": Expect p, cat, "LOCATOR_ADDRESS", "column overflow"
    x.Values("Address") = "A0": Expect p, cat, "LOCATOR_ADDRESS", "zero row"
    x.Values("Address") = "B2:A1": Expect p, cat, "LOCATOR_ADDRESS", "reverse range"
    x.Values("Address") = "A1"
    Entity(p, "RULE_REQUIRED").References(3)("Id") = "MISSING"
    Expect p, cat, "REF_MISSING", "missing field"
    Entity(p, "RULE_REQUIRED").References(3)("Id") = "FIELD_NAME"
    Set x = AtNewDefinition("Field", "FIELD_NAME", "Duplicate"): p.Definitions.Add x
    Expect p, cat, "ID_DUPLICATE", "duplicate id": p.Definitions.Remove p.Definitions.Count
    Entity(p, "FIELD_NAME").Values("DataType") = "Bogus"
    Expect p, cat, "TYPE_UNKNOWN", "invalid type"
    Entity(p, "FIELD_NAME").Values("DataType") = "Int64"
    Expect p, cat, "REF_TYPE", "reference type mismatch"
    Entity(p, "FIELD_NAME").Values("DataType") = "Text"
    Entity(p, "RULE_REQUIRED").Dependencies.Add "RULE_REQUIRED"
    Expect p, cat, "DEPENDENCY_CYCLE", "self cycle"
    Entity(p, "RULE_REQUIRED").Dependencies.Remove 1
    AtAddReference Entity(p, "RULE_REQUIRED"), "Operation", "Operation", "OP_UNKNOWN", "1.0.0", "Bool"
    Expect p, cat, "OPERATION_UNKNOWN", "unknown operation"
    Entity(p, "RULE_REQUIRED").References.Remove 5
    Entity(p, "RULE_REQUIRED").References(4)("Version") = "2.0.0"
    Expect p, cat, "OPERATION_UNKNOWN", "unknown operation version"
    Entity(p, "RULE_REQUIRED").References(4)("Version") = "1.0.0"
    p.SchemaVersion = "2.0.0": Expect p, cat, "SCHEMA_VERSION", "schema mismatch"
    p.SchemaVersion = "1.0.0": p.State = "Published"
    Expect p, cat, "PUBLISHED_REQUIRED", "published metadata missing"
    p.State = "Draft"
    Entity(p, "FIELD_NAME").Version = "1.x"
    Expect p, cat, "VERSION_INVALID", "entity version"
    Entity(p, "FIELD_NAME").Version = "1.0.0"
    Entity(p, "FIELD_NAME").DisplayName = ChrW(&H418) & ChrW(&H43C) & ChrW(&H44F) & " <&" & vbCrLf & "line" & vbTab & "%0D %25" & vbCr & "end"
    xml = AtToXml(p): Set q = AtFromXml(xml, cat, d)
    Check Entity(q, "FIELD_NAME").DisplayName = Entity(p, "FIELD_NAME").DisplayName, "unicode and whitespace"
    path = scratch & "\unicode.xml"
    Check AtSaveNew(p, path, cat, d), "unicode file save"
    Set q = AtLoad(path, cat, d)
    Check Entity(q, "FIELD_NAME").DisplayName = Entity(p, "FIELD_NAME").DisplayName, "unicode UTF8 disk"
    Set q = AtFromXml("<broken", cat, d): Check HasCode(d, "XML_INVALID"), "corrupt xml"
    Set q = AtFromXml(Replace(xml, "schemaVersion=""1.0.0""", "schemaVersion=""9.0.0"""), cat, d)
    Check q Is Nothing And HasCode(d, "SCHEMA_VERSION"), "unsupported XML schema"
    Set q = AtFromXml("<!DOCTYPE package [<!ENTITY e SYSTEM 'file:///C:/never-read'>]>" & xml, cat, d)
    Check q Is Nothing And HasCode(d, "XML_INVALID"), "DTD forbidden"
    Set q = AtFromXml(Replace(xml, "<definitions>", "<definitions><unknown/>"), cat, d)
    Check q Is Nothing And HasCode(d, "XML_SHAPE"), "unknown element rejected"
    Set q = AtFromXml(Replace(xml, "<properties>", "<properties><property name=""Injected"" value=""x""/>"), cat, d)
    Check AtHasErrors(d), "unknown property rejected"
    Set q = AtFromXml(Replace(xml, "<properties>", "<properties><property name=""DataType"" value=""Text""/><property name=""DataType"" value=""Text""/>"), cat, d)
    Check q Is Nothing And HasCode(d, "XML_SHAPE"), "duplicate XML property rejected"
    Set q = AtFromXml(Replace(xml, "<definitions>", "<definitions unexpected=""true"">"), cat, d)
    Check q Is Nothing And HasCode(d, "XML_SHAPE"), "unknown attribute rejected"
    Set q = AtFromXml(Replace(xml, "<definitions>", "<definitions><?run macro?>"), cat, d)
    Check q Is Nothing And HasCode(d, "XML_SHAPE"), "processing instruction rejected"
    Set q = AtFromXml(String$(524289, "x"), cat, d)
    Check q Is Nothing And HasCode(d, "FILE_LIMIT"), "text size bound"
    Dim deep As String, i As Long
    For i = 1 To 34: deep = deep & "<x>": Next i
    For i = 1 To 34: deep = deep & "</x>": Next i
    Set q = AtFromXml(deep, cat, d)
    Check q Is Nothing And HasCode(d, "XML_INVALID"), "parser depth bound"
    WriteText scratch & "\corrupt.xml", "<damaged"
    Set q = AtLoad(scratch & "\corrupt.xml", cat, d)
    Check q Is Nothing And HasCode(d, "XML_INVALID"), "corrupt file load"
    Set q = AtLoad(scratch & "\does-not-exist.xml", cat, d)
    Check q Is Nothing And HasCode(d, "IO_READ"), "missing file diagnostic"
    previous = FileBytes(path)
    Check Not AtSaveNew(p, scratch & "\absent\fail.xml", cat, d), "IO failure"
    Check HasCode(d, "IO_WRITE"), "IO error diagnostic"
    Check FileBytes(path) = previous, "previous bytes survive IO failure"
    Set q = AtLoad(path, cat, d)
    Check AtToXml(q) = AtToXml(p), "previous survives IO failure"
End Sub

Private Sub TestValidationEdges(ByVal repo As String)
    Dim p As CAtPackage, d As Collection, cat As Object, x As CAtDefinition, y As CAtDefinition
    Set cat = AtDictionary(): cat.Add "OP_TEST@1.0.0", "Bool"
    Set p = AtLoad(repo & "\tests\fixtures\simple.xml", cat, d)
    Set x = Entity(p, "FIELD_NAME")
    x.Id = "lower-id": Expect p, cat, "ID_INVALID", "ID grammar": x.Id = "FIELD_NAME"
    x.References(2)("Version") = "1.0.1": Expect p, cat, "REF_VERSION", "exact entity version"
    x.References(2)("Version") = "1.0.0"
    x.References.Remove 2: Expect p, cat, "REF_REQUIRED", "mandatory locator"
    AtAddReference x, "Locator", "Locator", "LOC_NAME", "1.0.0"
    Set y = AtNewDefinition("Template", "OTHER", "Other"): p.Definitions.Add y
    AtAddReference y, "Rule", "Rule", "RULE_REQUIRED", "1.0.0"
    Expect p, cat, "REF_OWNER", "template cannot include foreign rule"
    y.References.Remove 1
    Entity(p, "LOC_NAME").References(1)("Id") = "OTHER"
    Expect p, cat, "REF_OWNER", "foreign template locator"
    Entity(p, "LOC_NAME").References(1)("Id") = "SIMPLE"
    Set x = Entity(p, "LOC_NAME")
    x.Values.Remove "Address": x.Values("Strategy") = "DefinedName": x.Values.Add "DefinedName", "REQUEST_NAME"
    Set d = AtValidate(p, cat): Check Not AtHasErrors(d), "defined name valid"
    x.Values("DefinedName") = "A1": Expect p, cat, "LOCATOR_NAME", "name cannot be cell"
    x.Values.Remove "DefinedName": x.Values("Strategy") = "LabelOffset"
    x.Values.Add "Label", "Name": x.Values.Add "RowOffset", "1": x.Values.Add "ColumnOffset", "-1"
    Set d = AtValidate(p, cat): Check Not AtHasErrors(d), "label offsets valid"
    x.Values("RowOffset") = "1.5": Expect p, cat, "LOCATOR_OFFSET", "noninteger offset"
    x.Values("RowOffset") = "1": x.Values.Add "Address", "A1"
    Expect p, cat, "LOCATOR_SETTINGS", "conflicting locator settings"
    x.Values.Remove "Label": x.Values.Remove "RowOffset": x.Values.Remove "ColumnOffset": x.Values("Strategy") = "A1"
    AtAddReference Entity(p, "RULE_REQUIRED"), "Field", "Field", "FIELD_NAME", "1.0.0", "Text"
    Expect p, cat, "REF_DUPLICATE", "duplicate reference"
    Entity(p, "RULE_REQUIRED").References.Remove 5
    Set y = AtNewDefinition("Rule", "RULE_SECOND", "Second")
    AtAddReference y, "OwnerTemplate", "Template", "SIMPLE", "1.0.0"
    AtAddReference y, "Expression", "Expression", "EXPR_REQUIRED", "1.0.0", "Bool"
    p.Definitions.Add y
    y.References(1)("Id") = "OTHER"
    Entity(p, "RULE_REQUIRED").Dependencies.Add "RULE_SECOND"
    Expect p, cat, "REF_OWNER", "foreign rule dependency"
    Entity(p, "RULE_REQUIRED").Dependencies.Remove 1
    y.References(1)("Id") = "SIMPLE"
    Entity(p, "RULE_REQUIRED").Dependencies.Add "RULE_SECOND"
    y.Dependencies.Add "RULE_REQUIRED": Expect p, cat, "DEPENDENCY_CYCLE", "two-rule cycle"
    y.Dependencies.Remove 1
    Set d = AtValidate(p, cat): Check Not AtHasErrors(d), "acyclic dependency valid"
    y.Dependencies.Add "ABSENT": Expect p, cat, "DEPENDENCY_MISSING", "missing dependency"
    y.Dependencies.Remove 1
    Entity(p, "EXPR_REQUIRED").Values("Deferred") = "false"
    Expect p, cat, "AST_UNSUPPORTED", "AST bodies unsupported"
    Entity(p, "EXPR_REQUIRED").Values("Deferred") = "true"
    p.State = "Published": p.PublishedBy = "Synthetic test": p.PublishedAt = "2026-10-19T12:34:56Z"
    Set d = AtValidate(p, cat): Check Not AtHasErrors(d), "published structural metadata valid"
    Check HasCode(d, "AST_DEFERRED"), "published is not execution approval"
    p.PublishedAt = "2026-02-30T12:34:56Z": Expect p, cat, "PUBLISHED_REQUIRED", "invalid publication date"
    Set p = AtLoad(repo & "\tests\fixtures\equipment-return.xml", cat, d)
    Set y = AtNewDefinition("Column", "ITEM_SERIAL", "Serial")
    y.Values.Add "DataType", "Text": AtAddReference y, "OwnerTable", "Table", "ITEMS", "1.0.0": p.Definitions.Add y
    Set x = AtNewDefinition("Locator", "LOC_COLUMN", "Column location")
    x.Values.Add "Strategy", "TableColumn": x.Values.Add "Sheet", "Return"
    AtAddReference x, "OwnerTemplate", "Template", "EQUIPMENT_RETURN", "1.0.0"
    AtAddReference x, "Table", "Table", "ITEMS", "1.0.0"
    AtAddReference x, "Column", "Column", "ITEM_CODE", "1.0.0", "Text": p.Definitions.Add x
    Set d = AtValidate(p, cat): Check Not AtHasErrors(d), "table column locator valid"
    AtAddReference x, "Column", "Column", "ITEM_SERIAL", "1.0.0", "Text"
    Expect p, cat, "REF_REQUIRED", "table column cannot select two columns"
    x.References.Remove 4: x.Values("Strategy") = "A1": x.Values.Add "Address", "A1"
    Expect p, cat, "LOCATOR_SETTINGS", "A1 rejects table column references"
    p.Definitions.Remove p.Definitions.Count
    Set x = Entity(p, "CATALOG_CODE")
    x.References(1)("Role") = "OwnerTable": x.References(1)("Kind") = "Table": x.References(1)("Id") = "ITEMS"
    Expect p, cat, "REF_OWNER", "binding column from wrong source"
    Set p = Nothing: Set d = AtValidate(p, cat)
    Check HasCode(d, "MODEL_INVALID"), "malformed in-memory model diagnostic"
End Sub

Private Function FileBytes(ByVal path As String) As String
    Dim handle As Integer, buffer As String
    handle = FreeFile
    Open path For Binary Access Read As #handle
    buffer = Space$(LOF(handle)): Get #handle, , buffer
    Close #handle
    FileBytes = buffer
End Function

Private Sub TestNameRegressions(ByVal repo As String)
    Dim p As CAtPackage, d As Collection, cat As Object, x As CAtDefinition, value As Variant
    Set cat = AtDictionary(): cat.Add "OP_TEST@1.0.0", "Bool"
    Set p = AtLoad(repo & "\tests\fixtures\simple.xml", cat, d)
    Set x = Entity(p, "LOC_NAME")
    x.Values.Remove "Address": x.Values("Strategy") = "DefinedName"
    For Each value In Array("R1C1", "R2C3", "r1c1", "RC", "rc", "R1C", "RC1", "R", "c", "A1", "a1")
        x.Values("DefinedName") = CStr(value)
        Expect p, cat, "LOCATOR_NAME", "reserved DefinedName " & value
    Next value
    For Each value In Array("R1C1_NOTE", "_R1C1", "REQUEST_NAME", "RC_NOTE")
        x.Values("DefinedName") = CStr(value): Set d = AtValidate(p, cat)
        Check Not AtHasErrors(d), "allowed DefinedName " & value
    Next value
    Set p = AtLoad(repo & "\tests\fixtures\equipment-return.xml", cat, d)
    Set x = Entity(p, "CATALOG")
    Dim i As Long, ok As Boolean
    ok = True
    For i = 0 To 31
        x.Values("FileName") = "catalog" & ChrW(i) & ".xlsx": Set d = AtValidate(p, cat)
        If Not HasCode(d, "SOURCE_PATH") Then ok = False
    Next i
    Check ok, "FileName rejects all controls U+0000-U+001F"
    ok = True
    For Each value In Array("CON", "nul.xlsx", "PrN.csv", "AUX.backup.xlsx", "COM1", "com9.xlsx", "LPT1.txt", "Lpt9", "COM" & ChrW(&HB9) & ".xlsx", "LPT" & ChrW(&HB2), "COM" & ChrW(&HB3), "CON .xlsx")
        x.Values("FileName") = CStr(value): Set d = AtValidate(p, cat)
        If Not HasCode(d, "SOURCE_PATH") Then ok = False
    Next value
    Check ok, "FileName rejects device names regardless of case and extension"
    ok = True
    For Each value In Array("catalog.xlsx", "CONTRACT.xlsx", "NUL_data.xlsx", "COM10.xlsx", "LPT0.csv", "archive.backup.xlsx")
        x.Values("FileName") = CStr(value): Set d = AtValidate(p, cat)
        If AtHasErrors(d) Then ok = False
    Next value
    Check ok, "FileName allows ordinary names without blanket device prefix ban"
    ok = True
    For Each value In Array("sub\catalog.xlsx", "catalog.xlsx ", "catalog.xlsx.", "bad?.xlsx")
        x.Values("FileName") = CStr(value): Set d = AtValidate(p, cat)
        If Not HasCode(d, "SOURCE_PATH") Then ok = False
    Next value
    Check ok, "FileName retains existing leaf restrictions"
End Sub

Private Sub WriteText(ByVal path As String, ByVal text As String)
    Dim fso As Object, file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(path, True, True)
    file.Write text: file.Close
End Sub

Private Function DiagnosticsText(ByVal issues As Collection) As String
    Dim issue As Variant
    For Each issue In issues
        DiagnosticsText = DiagnosticsText & issue("Code") & ":" & issue("Reason") & ";"
    Next issue
End Function

Private Function Entity(ByVal p As CAtPackage, ByVal id As String) As CAtDefinition
    Dim x As CAtDefinition
    For Each x In p.Definitions
        If x.Id = id Then Set Entity = x: Exit Function
    Next x
    Err.Raise 5, , "Missing test entity " & id
End Function

Private Function HasCode(ByVal d As Collection, ByVal code As String) As Boolean
    Dim item As Variant
    For Each item In d
        If item("Code") = code Then HasCode = True: Exit Function
    Next item
End Function

Private Sub Expect(ByVal p As CAtPackage, ByVal cat As Object, ByVal code As String, ByVal label As String)
    Dim d As Collection: Set d = AtValidate(p, cat)
    Check HasCode(d, code), label
End Sub
