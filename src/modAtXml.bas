Attribute VB_Name = "modAtXml"
Option Explicit

Private Const NS As String = "urn:accesstools:definitions:1"
Private Const MAX_BYTES As Long = 1048576

Private Function NewDom() As Object
    Dim dom As Object: Set dom = CreateObject("Msxml2.DOMDocument.6.0")
    dom.async = False: dom.validateOnParse = False: dom.resolveExternals = False
    dom.preserveWhiteSpace = True
    dom.setProperty "ProhibitDTD", True
    ' MSXML counts blocks of 1024 characters, not megabytes.
    dom.setProperty "MaxXMLSize", 1024
    dom.setProperty "MaxElementDepth", 32
    Set NewDom = dom
End Function

Public Function AtToXml(ByVal p As CAtPackage) As String
    Dim dom As Object, root As Object, list As Object, node As Object, group As Object
    Dim item As Object, d As CAtDefinition, prop As Variant, r As Variant, dep As Variant
    If p.SchemaVersion <> "1.0.0" And p.SchemaVersion <> "2.0.0" Then Err.Raise 5, , "Unsupported schema"
    Set dom = NewDom()
    dom.appendChild dom.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    Set root = Element(dom, dom, "package")
    Attr root, "schemaVersion", p.SchemaVersion: Attr root, "id", p.Id
    Attr root, "version", p.Version: Attr root, "displayName", p.DisplayName
    Attr root, "state", p.State: Attr root, "publishedBy", p.PublishedBy: Attr root, "publishedAt", p.PublishedAt
    Set list = Element(dom, root, "definitions")
    For Each d In p.Definitions
        If p.SchemaVersion = "1.0.0" And (d.Nodes.Count > 0 Or d.RootNodeId <> "" Or d.ScopeTemplateId <> "") Then Err.Raise 5, , "AST cannot be written as schema 1.0.0"
        Set node = Element(dom, list, "definition")
        Attr node, "kind", d.Kind: Attr node, "id", d.Id
        Attr node, "version", d.Version: Attr node, "displayName", d.DisplayName
        Set group = Element(dom, node, "properties")
        For Each prop In d.Values.Keys
            Set item = Element(dom, group, "property")
            Attr item, "name", CStr(prop): Attr item, "value", CStr(d.Values(prop))
        Next prop
        Set group = Element(dom, node, "references")
        For Each r In d.References
            Set item = Element(dom, group, "ref")
            Attr item, "role", CStr(r("Role")): Attr item, "kind", CStr(r("Kind"))
            Attr item, "id", CStr(r("Id")): Attr item, "version", CStr(r("Version"))
            Attr item, "expectedType", CStr(r("ExpectedType"))
        Next r
        Set group = Element(dom, node, "dependencies")
        For Each dep In d.Dependencies
            Set item = Element(dom, group, "dependency"): Attr item, "id", CStr(dep)
        Next dep
        If p.SchemaVersion = "2.0.0" And d.Kind = "Expression" Then WriteAst dom, node, d
    Next d
    AtToXml = dom.XML
End Function

Private Function Element(ByVal dom As Object, ByVal parent As Object, ByVal name As String) As Object
    Dim node As Object: Set node = dom.createNode(1, name, NS)
    parent.appendChild node
    Set Element = node
End Function

Private Sub Attr(ByVal node As Object, ByVal name As String, ByVal value As String)
    node.setAttribute name, EncodeScalar(value)
End Sub

' MSXML normalizes CRLF even in attribute character references. Escape only
' percent and whitespace controls before entering DOM; ordinary Unicode is XML.
Private Function EncodeScalar(ByVal value As String) As String
    value = Replace(value, "%", "%25")
    value = Replace(value, vbCr, "%0D")
    value = Replace(value, vbLf, "%0A")
    EncodeScalar = Replace(value, vbTab, "%09")
End Function

Private Function ReadAttr(ByVal node As Object, ByVal name As String) As String
    Dim raw As String, value As String
    raw = node.getAttribute(name)
    value = Replace(raw, "%0D", vbCr)
    value = Replace(value, "%0A", vbLf)
    value = Replace(value, "%09", vbTab)
    value = Replace(value, "%25", "%")
    If EncodeScalar(value) <> raw Then Err.Raise 5, , "Noncanonical scalar escape"
    ReadAttr = value
End Function

Public Function AtFromXml(ByVal text As String, ByVal catalog As Object, ByRef diagnostics As Collection) As CAtPackage
    On Error GoTo Broken
    Set diagnostics = New Collection
    If Len(text) > MAX_BYTES \ 2 Then
        AtProblem diagnostics, "FILE_LIMIT", "Package", "Xml", "Text exceeds 524288 UTF-16 code units", "Use a smaller definition file.": Exit Function
    End If
    Dim dom As Object: Set dom = NewDom()
    If Not dom.loadXML(text) Then
        ParseProblem dom, diagnostics: Exit Function
    End If
    Set AtFromXml = Decode(dom, catalog, diagnostics)
    Exit Function
Broken:
    AtProblem diagnostics, "XML_RUNTIME", "Package", "Xml", CStr(Err.Number), "XML processing failed. Check that MSXML6 is available."
End Function

Public Function AtLoad(ByVal path As String, ByVal catalog As Object, ByRef diagnostics As Collection) As CAtPackage
    Set diagnostics = New Collection
    On Error GoTo Broken
    Dim fso As Object, dom As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Err.Raise 53
    If fso.GetFile(path).Size > MAX_BYTES Then
        AtProblem diagnostics, "FILE_LIMIT", "Package", "File", "File exceeds 1 MiB", "Use a smaller definition file.": Exit Function
    End If
    Set dom = NewDom()
    If Not dom.Load(path) Then ParseProblem dom, diagnostics: Exit Function
    Set AtLoad = Decode(dom, catalog, diagnostics)
    Exit Function
Broken:
    AtProblem diagnostics, "IO_READ", "Package", "File", CStr(Err.Number), "Cannot read the local definition file. Check the path, access rights and MSXML6 availability."
End Function

Private Sub ParseProblem(ByVal dom As Object, ByVal diagnostics As Collection)
    AtProblem diagnostics, "XML_INVALID", "Package", "Xml", "Parser " & dom.parseError.errorCode & " at line " & dom.parseError.Line, "The XML file is damaged or contains prohibited DTD content."
End Sub

Private Function Decode(ByVal dom As Object, ByVal catalog As Object, ByRef diagnostics As Collection) As CAtPackage
    On Error GoTo ShapeError
    Dim root As Object, definitions As Object, node As Object, group As Object, item As Object
    Dim p As CAtPackage, d As CAtDefinition, groups As Object, v As String, schema As String
    Dim objectId As String, nodeId As String
    For Each node In dom.childNodes
        If node.nodeType <> 1 Then
            If Not (node.nodeType = 7 And node.nodeName = "xml") Then Err.Raise 5
        End If
    Next node
    Set root = dom.documentElement
    Shape root, "package", "schemaVersion|id|version|displayName|state|publishedBy|publishedAt"
    schema = ReadAttr(root, "schemaVersion")
    If schema <> "1.0.0" And schema <> "2.0.0" Then
        AtProblem diagnostics, "SCHEMA_VERSION", "Package", "SchemaVersion", CStr(ReadAttr(root, "schemaVersion")), "This schema version is not supported. No definitions were interpreted.": Exit Function
    End If
    Set p = AtNewPackage(ReadAttr(root, "id"), ReadAttr(root, "displayName"))
    p.SchemaVersion = schema
    p.Version = ReadAttr(root, "version"): p.State = ReadAttr(root, "state")
    p.PublishedBy = ReadAttr(root, "publishedBy"): p.PublishedAt = ReadAttr(root, "publishedAt")
    Set definitions = OnlyElement(root, "definitions")
    Shape definitions, "definitions", ""
    For Each node In definitions.childNodes
        If node.nodeType = 1 Then
            If p.Definitions.Count >= 2000 Then Err.Raise 5
            Shape node, "definition", "kind|id|version|displayName"
            Set d = AtNewDefinition(ReadAttr(node, "kind"), ReadAttr(node, "id"), ReadAttr(node, "displayName"))
            objectId = d.Id: nodeId = ""
            d.Version = ReadAttr(node, "version")
            Set groups = AtDictionary()
            For Each group In node.childNodes
                If group.nodeType = 1 Then
                    v = group.baseName
                    If groups.Exists(v) Then Err.Raise 5
                    groups.Add v, True
                    If v <> "ast" Then Shape group, v, ""
                    Select Case v
                        Case "properties"
                            For Each item In group.childNodes
                                If item.nodeType = 1 Then
                                    Shape item, "property", "name|value": NoChildren item
                                    If d.Values.Exists(ReadAttr(item, "name")) Then Err.Raise 5
                                    d.Values.Add ReadAttr(item, "name"), ReadAttr(item, "value")
                                End If
                            Next item
                        Case "references"
                            For Each item In group.childNodes
                                If item.nodeType = 1 Then
                                    Shape item, "ref", "role|kind|id|version|expectedType": NoChildren item
                                    AtAddReference d, ReadAttr(item, "role"), ReadAttr(item, "kind"), ReadAttr(item, "id"), ReadAttr(item, "version"), ReadAttr(item, "expectedType")
                                End If
                            Next item
                        Case "dependencies"
                            For Each item In group.childNodes
                                If item.nodeType = 1 Then
                                    Shape item, "dependency", "id": NoChildren item
                                    d.Dependencies.Add ReadAttr(item, "id")
                                End If
                            Next item
                        Case "ast"
                            If schema <> "2.0.0" Or d.Kind <> "Expression" Then Err.Raise 5
                            ReadAst group, d, nodeId
                        Case Else: Err.Raise 5
                    End Select
                End If
            Next group
            If groups.Count <> IIf(schema = "2.0.0" And d.Kind = "Expression", 4, 3) Then Err.Raise 5
            If schema = "2.0.0" And d.Kind = "Expression" And Not groups.Exists("ast") Then Err.Raise 5
            p.Definitions.Add d
        End If
    Next node
    Set diagnostics = AtValidate(p, catalog)
    ' Invalid drafts can be inspected/repaired by Admin; never imply executable readiness.
    Set Decode = p
    Exit Function
ShapeError:
    If objectId <> "" Then
        AtAstProblem diagnostics, "XML_SHAPE", objectId, nodeId, "Unexpected or duplicate element/attribute; no definitions loaded."
    Else
        AtProblem diagnostics, "XML_SHAPE", "Package", "Xml", "Unexpected or duplicate element/attribute", "The file does not match its declared schema. No definitions were loaded."
    End If
End Function

Private Sub WriteAst(ByVal dom As Object, ByVal parent As Object, ByVal d As CAtDefinition)
    Dim ast As Object, node As Object, group As Object, item As Object, n As CAtNode
    Dim key As Variant, id As Variant, parameter As Object
    Set ast = Element(dom, parent, "ast")
    Attr ast, "root", d.RootNodeId: Attr ast, "template", d.ScopeTemplateId
    For Each n In d.Nodes
        Set node = Element(dom, ast, "node")
        Attr node, "id", n.Id: Attr node, "kind", n.Kind: Attr node, "type", n.ResultType
        Set group = Element(dom, node, "values")
        For Each key In n.Values.Keys
            Set item = Element(dom, group, "value")
            Attr item, "name", CStr(key): Attr item, "value", CStr(n.Values(key))
        Next key
        Set group = Element(dom, node, "parameters")
        For Each key In n.Parameters.Keys
            Set parameter = n.Parameters(key): Set item = Element(dom, group, "parameter")
            Attr item, "name", CStr(key): Attr item, "type", CStr(parameter("Type")): Attr item, "value", CStr(parameter("Value"))
        Next key
        Set group = Element(dom, node, "bindings")
        For Each id In n.Bindings
            Set item = Element(dom, group, "binding"): Attr item, "id", CStr(id)
        Next id
        Set group = Element(dom, node, "children")
        For Each id In n.Children
            Set item = Element(dom, group, "child"): Attr item, "id", CStr(id)
        Next id
    Next n
End Sub

Private Sub ReadAst(ByVal ast As Object, ByVal d As CAtDefinition, ByRef nodeId As String)
    Dim node As Object, group As Object, item As Object, n As CAtNode, groups As Object, key As String
    Shape ast, "ast", "root|template"
    d.RootNodeId = ReadAttr(ast, "root"): d.ScopeTemplateId = ReadAttr(ast, "template")
    For Each node In ast.childNodes
        If node.nodeType = 1 Then
            nodeId = CStr(node.getAttribute("id"))
            Shape node, "node", "id|kind|type"
            Set n = AtNewNode(ReadAttr(node, "id"), ReadAttr(node, "kind"), ReadAttr(node, "type"))
            nodeId = n.Id: Set groups = AtDictionary()
            For Each group In node.childNodes
                If group.nodeType = 1 Then
                    key = group.baseName: Shape group, key, ""
                    If groups.Exists(key) Then Err.Raise 5
                    groups.Add key, True
                    For Each item In group.childNodes
                        If item.nodeType = 1 Then
                            Select Case key
                                Case "values"
                                    Shape item, "value", "name|value": NoChildren item
                                    If n.Values.Exists(ReadAttr(item, "name")) Then Err.Raise 5
                                    n.Values.Add ReadAttr(item, "name"), ReadAttr(item, "value")
                                Case "parameters"
                                    Shape item, "parameter", "name|type|value": NoChildren item
                                    If n.Parameters.Exists(ReadAttr(item, "name")) Then Err.Raise 5
                                    n.Parameters.Add ReadAttr(item, "name"), AtTypedParameter(ReadAttr(item, "type"), ReadAttr(item, "value"))
                                Case "bindings"
                                    Shape item, "binding", "id": NoChildren item: n.Bindings.Add ReadAttr(item, "id")
                                Case "children"
                                    Shape item, "child", "id": NoChildren item: n.Children.Add ReadAttr(item, "id")
                                Case Else: Err.Raise 5
                            End Select
                        End If
                    Next item
                    If key <> "values" And key <> "parameters" And key <> "bindings" And key <> "children" Then Err.Raise 5
                End If
            Next group
            If groups.Count <> 4 Then Err.Raise 5
            d.Nodes.Add n
        End If
    Next node
End Sub

Private Sub Shape(ByVal node As Object, ByVal name As String, ByVal attributes As String)
    Dim a As Object, child As Object, expected As Long
    If node Is Nothing Then Err.Raise 5
    If node.nodeType <> 1 Or node.baseName <> name Or node.namespaceURI <> NS Then Err.Raise 5
    For Each a In node.Attributes
        If a.namespaceURI <> "http://www.w3.org/2000/xmlns/" Then
            If a.namespaceURI <> "" Or InStr(1, "|" & attributes & "|", "|" & a.nodeName & "|", vbBinaryCompare) = 0 Then Err.Raise 5
            expected = expected + 1
        End If
    Next a
    If attributes <> "" Then
        If expected <> UBound(Split(attributes, "|")) + 1 Then Err.Raise 5
    ElseIf expected <> 0 Then
        Err.Raise 5
    End If
    For Each child In node.childNodes
        If child.nodeType <> 1 Then
            If child.nodeType <> 3 Or Trim$(Replace(Replace(Replace(child.Text, vbCr, ""), vbLf, ""), vbTab, "")) <> "" Then Err.Raise 5
        End If
    Next child
End Sub

Private Sub NoChildren(ByVal node As Object)
    Dim child As Object
    For Each child In node.childNodes
        If child.nodeType = 1 Then Err.Raise 5
    Next child
End Sub

Private Function OnlyElement(ByVal parent As Object, ByVal name As String) As Object
    Dim child As Object, count As Long
    For Each child In parent.childNodes
        If child.nodeType = 1 Then
            count = count + 1
            If child.baseName <> name Then Err.Raise 5
            Set OnlyElement = child
        End If
    Next child
    If count <> 1 Then Err.Raise 5
End Function

' Append-only commit: existing revisions are NEVER overwritten, including on failure.
' The caller selects the new revision path. No active pointer/publisher is implemented.
Public Function AtSaveNew(ByVal p As CAtPackage, ByVal path As String, ByVal catalog As Object, ByRef diagnostics As Collection) As Boolean
    On Error GoTo Broken
    Set diagnostics = AtValidate(p, catalog)
    If AtHasErrors(diagnostics) Then Exit Function
    Dim fso As Object, temp As String, owned As Boolean, stream As Object
    Dim dom As Object, q As CAtPackage, readIssues As Collection, xml As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.GetAbsolutePathName(path)
    If fso.FileExists(path) Or fso.FolderExists(path) Then
        AtProblem diagnostics, "FILE_EXISTS", p.Id, "File", "Destination already exists", "Choose a new revision file. Existing files are never replaced.": Exit Function
    End If
    temp = fso.BuildPath(fso.GetParentFolderName(path), fso.GetTempName)
    Set stream = fso.CreateTextFile(temp, False): owned = True
    stream.Close: Set stream = Nothing
    xml = AtToXml(p): Set dom = NewDom()
    If Not dom.loadXML(xml) Then Err.Raise 5
    dom.Save temp
    Set q = AtLoad(temp, catalog, readIssues)
    If q Is Nothing Then Err.Raise 5
    If AtHasErrors(readIssues) Then Err.Raise 5
    If AtToXml(q) <> xml Then Err.Raise 5, , "Round-trip mismatch"
    ' VBA Name refuses an existing target. Temp and final are on the same volume.
    Name temp As path
    owned = False: AtSaveNew = True
    Exit Function
Broken:
    Dim reason As String: reason = CStr(Err.Number)
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If owned Then fso.DeleteFile temp, True
    On Error GoTo 0
    If diagnostics Is Nothing Then Set diagnostics = New Collection
    AtProblem diagnostics, "IO_WRITE", "Package", "File", reason, "The new revision was not committed. Previous revision files were not changed. Check path, permissions and free space."
End Function
