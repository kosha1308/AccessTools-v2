Attribute VB_Name = "modAtValidation"
Option Explicit

Private Const KINDS As String = "|Template|Field|Locator|Table|Column|Source|Binding|Rule|Action|Expression|"
Private Const TYPES As String = "|Text|Bool|Int64|Decimal|Date|DateTime|Email|Hostname|IpSet|PortSet|"

' operationCatalog is trusted host metadata: exact ID@version -> declared return type.
' Configurations cannot register operations; no operation is executed here.
Public Function AtValidate(ByVal p As CAtPackage, ByVal operationCatalog As Object) As Collection
    Dim issues As New Collection, index As Object, colors As Object
    Dim d As CAtDefinition, r As Variant, key As String, prop As Variant
    On Error GoTo InvalidModel
    Set index = AtDictionary(): Set colors = AtDictionary()
    If p Is Nothing Then Err.Raise 5, , "Package is Nothing"
    If p.SchemaVersion <> "1.0.0" Then AtProblem issues, "SCHEMA_VERSION", p.Id, "SchemaVersion", p.SchemaVersion, "This schema version is not supported. Use reader 1.0.0."
    Identity issues, "Package", p.Id, p.Version, p.DisplayName
    If p.State <> "Draft" And p.State <> "Published" Then AtProblem issues, "STATE_INVALID", p.Id, "State", p.State, "Choose Draft or Published."
    If p.Definitions.Count = 0 Then AtProblem issues, "DEFINITIONS_REQUIRED", p.Id, "Definitions", "Empty package", "Add at least one template definition."
    If p.Definitions.Count > 2000 Then AtProblem issues, "LIMIT", p.Id, "Definitions", "More than 2000 definitions", "Split the package into smaller packages.": GoTo Done
    If p.State = "Published" Then
        If Len(Trim$(p.PublishedBy)) = 0 Or Not UtcTimestamp(p.PublishedAt) Then AtProblem issues, "PUBLISHED_REQUIRED", p.Id, "PublishedBy/PublishedAt", "Missing author or invalid UTC timestamp", "Published metadata requires author and YYYY-MM-DDTHH:MM:SSZ timestamp."
    ElseIf p.PublishedBy <> "" Or p.PublishedAt <> "" Then
        AtProblem issues, "DRAFT_METADATA", p.Id, "PublishedBy/PublishedAt", "Draft has publication metadata", "Remove publication metadata from the draft."
    End If
    For Each d In p.Definitions
        Identity issues, d.Kind, d.Id, d.Version, d.DisplayName
        If Not InList(d.Kind, KINDS) Then AtProblem issues, "KIND_UNKNOWN", d.Id, "Kind", d.Kind, "Choose a supported definition kind."
        key = d.Kind & "/" & d.Id
        If index.Exists(key) Then
            AtProblem issues, "ID_DUPLICATE", d.Id, "Id", key, "An ID may occur only once within its definition kind."
        Else
            index.Add key, d
        End If
        For Each prop In d.Values.Keys
            If Not InList(CStr(prop), Properties(d.Kind)) Then AtProblem issues, "PROPERTY_UNKNOWN", d.Id, CStr(prop), "Not allowed for " & d.Kind, "Remove the unknown property or use a supported schema."
            If VarType(d.Values(prop)) <> vbString Then AtProblem issues, "PROPERTY_TYPE", d.Id, CStr(prop), "Storage scalars must be strings", "Use the declared lexical form of the property."
        Next prop
        ValidateProperties d, issues
    Next d
    If Not ContainsKind(p, "Template") Then AtProblem issues, "TEMPLATE_REQUIRED", p.Id, "Definitions", "No Template", "Add a template definition."
    For Each d In p.Definitions
        ValidateReferences d, index, operationCatalog, issues
    Next d
    For Each d In p.Definitions
        If d.Kind = "Rule" Or d.Kind = "Expression" Then Visit d, index, colors, issues, 0
    Next d
    AtProblem issues, "AST_DEFERRED", p.Id, "Expressions", "Only expression identity and declared result type validated", "Structural validation does not prove AST correctness or publication readiness (PR-03).", "Info"
Done:
    Set AtValidate = issues
    Exit Function
InvalidModel:
    AtProblem issues, "MODEL_INVALID", "Package", "Model", CStr(Err.Number) & ": " & Err.Description, "The in-memory model is malformed. Recreate the affected definition."
    Resume Done
End Function

Private Sub Identity(ByVal issues As Collection, ByVal kind As String, ByVal id As String, ByVal version As String, ByVal name As String)
    If Not AtValidId(id) Then AtProblem issues, "ID_INVALID", id, "Id", kind, "Use 1-64 uppercase ASCII letters, digits or underscores, starting with a letter."
    If Not AtValidVersion(version) Then AtProblem issues, "VERSION_INVALID", id, "Version", version, "Use an exact three-part numeric version, for example 1.0.0."
    If Len(Trim$(name)) = 0 Then AtProblem issues, "NAME_REQUIRED", id, "DisplayName", "Empty name", "Enter a display name."
End Sub

Public Function AtValidId(ByVal value As String) As Boolean
    Dim i As Long, c As Long
    If Len(value) = 0 Or Len(value) > 64 Then Exit Function
    For i = 1 To Len(value)
        c = AscW(Mid$(value, i, 1))
        If Not (c >= 65 And c <= 90) Then
            If i = 1 Or Not ((c >= 48 And c <= 57) Or c = 95) Then Exit Function
        End If
    Next i
    AtValidId = True
End Function

Public Function AtValidVersion(ByVal value As String) As Boolean
    Dim parts As Variant, item As Variant
    parts = Split(value, ".")
    If UBound(parts) <> 2 Then Exit Function
    For Each item In parts
        If Not Digits(CStr(item)) Or Len(CStr(item)) > 5 Then Exit Function
        If Len(CStr(item)) > 1 And Left$(CStr(item), 1) = "0" Then Exit Function
        If CLng(item) > 65535 Then Exit Function
    Next item
    AtValidVersion = True
End Function

Private Function Digits(ByVal value As String) As Boolean
    Dim i As Long, c As Long
    If Len(value) = 0 Then Exit Function
    For i = 1 To Len(value)
        c = AscW(Mid$(value, i, 1))
        If c < 48 Or c > 57 Then Exit Function
    Next i
    Digits = True
End Function

Private Function InList(ByVal value As String, ByVal choices As String) As Boolean
    If value <> "" And InStr(value, "|") = 0 Then InList = InStr(1, choices, "|" & value & "|", vbBinaryCompare) > 0
End Function

Private Function Properties(ByVal kind As String) As String
    Select Case kind
        Case "Field", "Column": Properties = "|DataType|"
        Case "Locator": Properties = "|Strategy|Sheet|Address|DefinedName|Label|RowOffset|ColumnOffset|"
        Case "Source": Properties = "|Adapter|FileName|Sheet|"
        Case "Expression": Properties = "|ReturnType|Deferred|"
        Case "Action": Properties = "|Trigger|Audience|Channel|"
        Case Else: Properties = "|"
    End Select
End Function

Private Sub Required(ByVal d As CAtDefinition, ByVal member As String, ByVal issues As Collection)
    If Len(Trim$(AtValue(d, member))) = 0 Then AtProblem issues, "PROPERTY_REQUIRED", d.Id, member, "Missing value", "Enter the required property: " & member & "."
End Sub

Private Sub ValidateProperties(ByVal d As CAtDefinition, ByVal issues As Collection)
    Dim strategy As String, prop As Variant
    If d.Kind = "Source" Or d.Kind = "Locator" Then
        If Not SheetNameValid(AtValue(d, "Sheet")) Then AtProblem issues, "SHEET_INVALID", d.Id, "Sheet", "Invalid literal worksheet name", "Use a nonempty worksheet name up to 31 characters without control characters or : \\ / ? * [ ]."
    End If
    Select Case d.Kind
        Case "Field", "Column", "Expression"
            prop = IIf(d.Kind = "Expression", "ReturnType", "DataType")
            If Not InList(AtValue(d, CStr(prop)), TYPES) Then AtProblem issues, "TYPE_UNKNOWN", d.Id, CStr(prop), AtValue(d, CStr(prop)), "Choose a supported declared type. No implicit conversion is performed."
            If d.Kind = "Expression" And AtValue(d, "Deferred") <> "true" Then AtProblem issues, "AST_UNSUPPORTED", d.Id, "Deferred", "Only structural AST references supported", "Set Deferred=true; AST content belongs to PR-03."
        Case "Source"
            If AtValue(d, "Adapter") <> "Excel" Then AtProblem issues, "SOURCE_ADAPTER", d.Id, "Adapter", AtValue(d, "Adapter"), "This structural schema supports Excel source metadata."
            Required d, "Sheet", issues: Required d, "FileName", issues
            If Not SafeLeaf(AtValue(d, "FileName")) Then AtProblem issues, "SOURCE_PATH", d.Id, "FileName", "Not a leaf file name", "Store a file name only; choosing a trusted source path is a separate host responsibility."
        Case "Locator"
            Required d, "Sheet", issues
            strategy = AtValue(d, "Strategy")
            If Not InList(strategy, "|A1|DefinedName|LabelOffset|TableColumn|") Then AtProblem issues, "LOCATOR_STRATEGY", d.Id, "Strategy", strategy, "Choose A1, DefinedName, LabelOffset or TableColumn."
            Select Case strategy
                Case "A1"
                    If Not AtValidAddress(AtValue(d, "Address")) Then AtProblem issues, "LOCATOR_ADDRESS", d.Id, "Address", AtValue(d, "Address"), "Use a bounded A1 cell or ascending rectangular range within XFD1048576."
                Case "DefinedName"
                    If Not SafeName(AtValue(d, "DefinedName")) Then AtProblem issues, "LOCATOR_NAME", d.Id, "DefinedName", AtValue(d, "DefinedName"), "Use a simple ASCII workbook name, not an address or formula."
                Case "LabelOffset"
                    Required d, "Label", issues
                    If Not OffsetValid(AtValue(d, "RowOffset"), 1048575) Or Not OffsetValid(AtValue(d, "ColumnOffset"), 16383) Then AtProblem issues, "LOCATOR_OFFSET", d.Id, "Offsets", "Invalid integer offset", "Use integer offsets inside Excel bounds."
            End Select
            For Each prop In Array("Address", "DefinedName", "Label", "RowOffset", "ColumnOffset")
                If d.Values.Exists(prop) Then
                    If Not ((strategy = "A1" And prop = "Address") Or (strategy = "DefinedName" And prop = "DefinedName") Or (strategy = "LabelOffset" And InList(CStr(prop), "|Label|RowOffset|ColumnOffset|"))) Then AtProblem issues, "LOCATOR_SETTINGS", d.Id, CStr(prop), "Property conflicts with strategy", "Remove settings belonging to another locator strategy."
                End If
            Next prop
        Case "Action"
            If Not InList(AtValue(d, "Trigger"), "|Error|Warning|Info|Approval|") Then AtProblem issues, "ACTION_TRIGGER", d.Id, "Trigger", AtValue(d, "Trigger"), "Choose Error, Warning, Info or Approval."
            If Not InList(AtValue(d, "Audience"), "|Initiator|FirstLine|Coordinator|") Then AtProblem issues, "ACTION_AUDIENCE", d.Id, "Audience", AtValue(d, "Audience"), "Choose the intended audience."
            If Not InList(AtValue(d, "Channel"), "|Results|Comment|Report|") Then AtProblem issues, "ACTION_CHANNEL", d.Id, "Channel", AtValue(d, "Channel"), "Choose Results, Comment or Report. No output is executed by this component."
    End Select
End Sub

Private Function SheetNameValid(ByVal text As String) As Boolean
    Dim c As Variant, i As Long
    If Len(Trim$(text)) = 0 Or Len(text) > 31 Then Exit Function
    If Left$(text, 1) = "'" Or Right$(text, 1) = "'" Then Exit Function
    For Each c In Array(":", "\", "/", "?", "*", "[", "]")
        If InStr(text, CStr(c)) > 0 Then Exit Function
    Next c
    For i = 1 To Len(text)
        If AscW(Mid$(text, i, 1)) >= 0 And AscW(Mid$(text, i, 1)) < 32 Then Exit Function
    Next i
    SheetNameValid = True
End Function

Private Function RoleKind(ByVal kind As String, ByVal role As String) As String
    Select Case role
        Case "OwnerTemplate": If InList(kind, "|Field|Locator|Table|Binding|Rule|Action|") Then RoleKind = "Template"
        Case "Locator": If kind = "Field" Or kind = "Table" Then RoleKind = "Locator"
        Case "OwnerTable": If kind = "Column" Then RoleKind = "Table"
        Case "OwnerSource": If kind = "Column" Then RoleKind = "Source"
        Case "Source": If kind = "Binding" Then RoleKind = "Source"
        Case "Column": If InList(kind, "|Binding|Locator|") Then RoleKind = "Column"
        Case "Table": If kind = "Locator" Or kind = "Rule" Then RoleKind = "Table"
        Case "Expression": If kind = "Rule" Then RoleKind = "Expression"
        Case "Field": If kind = "Rule" Or kind = "Action" Then RoleKind = "Field"
        Case "Rule": If kind = "Action" Or kind = "Template" Then RoleKind = "Rule"
        Case "Operation": If kind = "Rule" Or kind = "Expression" Then RoleKind = "Operation"
    End Select
End Function

Private Sub ValidateReferences(ByVal d As CAtDefinition, ByVal index As Object, ByVal catalog As Object, ByVal issues As Collection)
    Dim r As Variant, counts As Object, seen As Object, key As String, target As CAtDefinition
    Dim actual As String, role As String, member As Variant, expected As String, dep As Variant
    Set counts = AtDictionary(): Set seen = AtDictionary()
    For Each r In d.References
        role = CStr(r("Role")): key = CStr(r("Kind")) & "/" & CStr(r("Id"))
        expected = RoleKind(d.Kind, role)
        If expected = "" Or expected <> r("Kind") Then AtProblem issues, "REF_ROLE", d.Id, role, key, "The reference kind is not allowed for this role."
        If d.Kind = "Locator" And (role = "Table" Or role = "Column") Then
            If AtValue(d, "Strategy") <> "TableColumn" Then AtProblem issues, "LOCATOR_SETTINGS", d.Id, role, "Reference conflicts with strategy", "Table and Column references require the TableColumn strategy."
        End If
        If Not counts.Exists(role) Then counts.Add role, 0
        counts(role) = counts(role) + 1
        If seen.Exists(role & "/" & key) Then AtProblem issues, "REF_DUPLICATE", d.Id, role, key, "Remove the duplicate reference."
        seen(role & "/" & key) = True
        If Not AtValidId(CStr(r("Id"))) Or Not AtValidVersion(CStr(r("Version"))) Then AtProblem issues, "REF_ID_VERSION", d.Id, role, key, "Reference requires a valid ID and exact version."
        actual = ""
        If r("Kind") = "Operation" Then
            key = r("Id") & "@" & r("Version")
            If catalog Is Nothing Then
                AtProblem issues, "OPERATION_UNKNOWN", d.Id, role, key, "The host operation catalog is unavailable."
            ElseIf Not catalog.Exists(key) Then
                AtProblem issues, "OPERATION_UNKNOWN", d.Id, role, key, "Install or select the exact registered operation version."
            Else
                actual = CStr(catalog(key))
                If Not InList(actual, TYPES) Then AtProblem issues, "CATALOG_TYPE", d.Id, role, actual, "The trusted operation metadata has an unsupported result type."
            End If
        ElseIf Not index.Exists(key) Then
            AtProblem issues, "REF_MISSING", d.Id, role, key, "The referenced definition does not exist in this package."
        Else
            Set target = index(key)
            If target.Version <> r("Version") Then AtProblem issues, "REF_VERSION", d.Id, role, key, "The referenced definition version does not match the exact package version."
            If target.Kind = "Field" Or target.Kind = "Column" Then actual = AtValue(target, "DataType")
            If target.Kind = "Expression" Then actual = AtValue(target, "ReturnType")
            If role = "Expression" And actual <> "Bool" Then AtProblem issues, "REF_TYPE", d.Id, role, actual, "A rule predicate must declare Bool. AST checking is deferred."
        End If
        If r("ExpectedType") <> "" Then
            If Not InList(CStr(r("ExpectedType")), TYPES) Or actual <> r("ExpectedType") Then AtProblem issues, "REF_TYPE", d.Id, role, CStr(r("ExpectedType")) & " vs " & actual, "The expected and declared reference types must match exactly."
        ElseIf InList(CStr(r("Kind")), "|Field|Column|Expression|Operation|") Then
            AtProblem issues, "REF_TYPE_REQUIRED", d.Id, role, key, "Enter the expected declared result type."
        End If
    Next r
    Select Case d.Kind
        Case "Field", "Table": Need counts, "OwnerTemplate", d, issues: Need counts, "Locator", d, issues
        Case "Locator": Need counts, "OwnerTemplate", d, issues
            If AtValue(d, "Strategy") = "TableColumn" Then Need counts, "Table", d, issues: Need counts, "Column", d, issues
        Case "Column"
            If CountOf(counts, "OwnerTable") + CountOf(counts, "OwnerSource") <> 1 Then AtProblem issues, "REF_REQUIRED", d.Id, "Owner", "Exactly one owner required", "Select one owning table or source."
        Case "Binding": Need counts, "OwnerTemplate", d, issues: Need counts, "Source", d, issues: Need counts, "Column", d, issues
        Case "Rule": Need counts, "OwnerTemplate", d, issues: Need counts, "Expression", d, issues
        Case "Action": Need counts, "OwnerTemplate", d, issues: Need counts, "Rule", d, issues
    End Select
    For Each member In Array("OwnerTemplate", "Locator", "OwnerTable", "OwnerSource", "Source", "Expression", "Table")
        If CountOf(counts, CStr(member)) > 1 Then AtProblem issues, "REF_CARDINALITY", d.Id, CStr(member), "Multiple singleton references", "Select at most one reference for this role."
    Next member
    For Each dep In d.Dependencies
        If d.Kind <> "Rule" And d.Kind <> "Expression" Then AtProblem issues, "DEPENDENCY_ROLE", d.Id, "Dependencies", d.Kind, "Only rules and expressions have execution dependencies."
        If Not index.Exists(d.Kind & "/" & CStr(dep)) Then
            AtProblem issues, "DEPENDENCY_MISSING", d.Id, "Dependencies", CStr(dep), "The dependency must exist in the same definition kind."
        ElseIf d.Kind = "Rule" Then
            Set target = index(d.Kind & "/" & CStr(dep))
            If RefId(target, "OwnerTemplate") <> RefId(d, "OwnerTemplate") Then AtProblem issues, "REF_OWNER", d.Id, "Dependencies", CStr(dep), "Rule dependencies must belong to the same template."
        End If
    Next dep
    ValidateOwnership d, index, issues
End Sub

Private Sub ValidateOwnership(ByVal d As CAtDefinition, ByVal index As Object, ByVal issues As Collection)
    Dim r As Variant, target As CAtDefinition, owner As String
    owner = RefId(d, "OwnerTemplate")
    If d.Kind = "Template" Then owner = d.Id
    For Each r In d.References
        If index.Exists(r("Kind") & "/" & r("Id")) Then
            Set target = index(r("Kind") & "/" & r("Id"))
            If owner <> "" And RefId(target, "OwnerTemplate") <> "" Then
                If owner <> RefId(target, "OwnerTemplate") Then AtProblem issues, "REF_OWNER", d.Id, CStr(r("Role")), target.Id, "The referenced definition belongs to another template."
            End If
            If d.Kind = "Binding" And r("Role") = "Column" Then
                If RefId(target, "OwnerSource") <> RefId(d, "Source") Then AtProblem issues, "REF_OWNER", d.Id, "Column", target.Id, "The bound column must belong to the selected source."
            End If
            If d.Kind = "Locator" And AtValue(d, "Strategy") = "TableColumn" And r("Role") = "Column" Then
                If RefId(target, "OwnerTable") <> RefId(d, "Table") Then AtProblem issues, "REF_OWNER", d.Id, "Column", target.Id, "The located column must belong to the selected table."
            End If
        End If
    Next r
End Sub

Private Function RefId(ByVal d As CAtDefinition, ByVal role As String) As String
    Dim r As Variant
    For Each r In d.References
        If r("Role") = role Then RefId = r("Id"): Exit Function
    Next r
End Function

Private Sub Need(ByVal counts As Object, ByVal role As String, ByVal d As CAtDefinition, ByVal issues As Collection)
    If CountOf(counts, role) <> 1 Then AtProblem issues, "REF_REQUIRED", d.Id, role, "Exactly one reference required", "Select the required reference: " & role & "."
End Sub

Private Function CountOf(ByVal counts As Object, ByVal key As String) As Long
    If counts.Exists(key) Then CountOf = counts(key)
End Function

Private Sub Visit(ByVal d As CAtDefinition, ByVal index As Object, ByVal colors As Object, ByVal issues As Collection, ByVal depth As Long)
    Dim key As String, dep As Variant, child As CAtDefinition
    key = d.Kind & "/" & d.Id
    If colors.Exists(key) Then
        If colors(key) = 1 Then AtProblem issues, "DEPENDENCY_CYCLE", d.Id, "Dependencies", key, "Remove the circular execution dependency. Ownership links are not execution dependencies."
        Exit Sub
    End If
    If depth > 128 Then AtProblem issues, "DEPENDENCY_DEPTH", d.Id, "Dependencies", "Depth exceeds 128", "Reduce dependency depth.": Exit Sub
    colors.Add key, 1
    For Each dep In d.Dependencies
        If index.Exists(d.Kind & "/" & dep) Then
            Set child = index(d.Kind & "/" & dep)
            Visit child, index, colors, issues, depth + 1
        End If
    Next dep
    colors(key) = 2
End Sub

Private Function ContainsKind(ByVal p As CAtPackage, ByVal kind As String) As Boolean
    Dim d As CAtDefinition
    For Each d In p.Definitions
        If d.Kind = kind Then ContainsKind = True: Exit Function
    Next d
End Function

Public Function AtValidAddress(ByVal address As String) As Boolean
    Dim parts As Variant, c1 As Long, r1 As Long, c2 As Long, r2 As Long
    parts = Split(address, ":")
    If UBound(parts) > 1 Then Exit Function
    If Not CellParts(CStr(parts(0)), c1, r1) Then Exit Function
    If UBound(parts) = 1 Then
        If Not CellParts(CStr(parts(1)), c2, r2) Then Exit Function
        If c2 < c1 Or r2 < r1 Then Exit Function
    End If
    AtValidAddress = True
End Function

Private Function CellParts(ByVal text As String, ByRef col As Long, ByRef row As Long) As Boolean
    Dim i As Long, c As Long, start As Long
    i = 1: col = 0: row = 0
    If Left$(text, 1) = "$" Then i = i + 1
    start = i
    Do While i <= Len(text)
        c = AscW(Mid$(text, i, 1))
        If c < 65 Or c > 90 Then Exit Do
        col = col * 26 + c - 64
        If col > 16384 Then Exit Function
        i = i + 1
    Loop
    If i = start Or col = 0 Then Exit Function
    If Mid$(text, i, 1) = "$" Then i = i + 1
    text = Mid$(text, i)
    If Not Digits(text) Or Len(text) > 7 Then Exit Function
    If Left$(text, 1) = "0" Then Exit Function
    row = CLng(text)
    CellParts = row >= 1 And row <= 1048576
End Function

Private Function OffsetValid(ByVal text As String, ByVal bound As Long) As Boolean
    Dim unsigned As String
    unsigned = text
    If Left$(unsigned, 1) = "-" Then unsigned = Mid$(unsigned, 2)
    If Not Digits(unsigned) Or Len(unsigned) > 7 Then Exit Function
    OffsetValid = Abs(CDbl(text)) <= bound
End Function

Private Function SafeName(ByVal text As String) As Boolean
    Dim i As Long, c As String
    If Len(text) = 0 Or Len(text) > 255 Or AtValidAddress(UCase$(text)) Then Exit Function
    For i = 1 To Len(text)
        c = Mid$(text, i, 1)
        If Not (c Like "[A-Za-z_]" Or (i > 1 And c Like "[0-9.]")) Then Exit Function
    Next i
    If UCase$(text) = "R" Or UCase$(text) = "C" Then Exit Function
    SafeName = True
End Function

Private Function SafeLeaf(ByVal text As String) As Boolean
    Dim c As Variant
    If Len(text) = 0 Or Len(text) > 255 Or text = "." Or text = ".." Then Exit Function
    For Each c In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        If InStr(text, CStr(c)) > 0 Then Exit Function
    Next c
    SafeLeaf = Right$(text, 1) <> "." And Right$(text, 1) <> " "
End Function

Private Function UtcTimestamp(ByVal text As String) As Boolean
    On Error GoTo Invalid
    If Not text Like "####-##-##T##:##:##Z" Then Exit Function
    Dim actual As Date
    actual = DateSerial(CInt(Left$(text, 4)), CInt(Mid$(text, 6, 2)), CInt(Mid$(text, 9, 2))) + TimeSerial(CInt(Mid$(text, 12, 2)), CInt(Mid$(text, 15, 2)), CInt(Mid$(text, 18, 2)))
    UtcTimestamp = Format$(actual, "yyyy-mm-dd") & "T" & Format$(actual, "hh:nn:ss") & "Z" = text
Invalid:
End Function
