VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtOperationRegistry"
Option Explicit

' Metadata and dispatch are paired in trusted VBA. No user callback or macro name.
Public Property Get Catalog() As Object
    Dim c As Object, args As Collection, s As Object
    Set c = AtDictionary()
    Set args = New Collection: args.Add "Text": args.Add "Text"
    Set s = AtSignature("Bool", args): s("Parameters").Add "MODE", "Text": s("States").Add "INVALID"
    c.Add "OP_TEXT_EQUAL@1.0.0", s
    Set args = New Collection: args.Add "Text"
    c.Add "OP_TEXT_TRIM@1.0.0", AtSignature("Text", args)
    Set args = New Collection: args.Add "Bool": args.Add "Bool"
    c.Add "OP_BOOL_EQUAL@1.0.0", AtSignature("Bool", args)
    Set args = New Collection: args.Add "Int64": args.Add "Int64"
    c.Add "OP_INT64_LE@1.0.0", AtSignature("Bool", args)
    Set Catalog = c
End Property

Public Function ExecuteOperation(ByVal operationId As String, ByVal version As String, ByVal args As Collection, ByVal parameters As Object) As CAtEvalResult
    On Error GoTo Broken
    Dim c As Object, s As Object, key As String, r As CAtEvalResult, a As CAtEvalResult, b As CAtEvalResult
    Dim typeName As String, i As Long, prop As Variant, parameter As Object, ok As Boolean
    typeName = "Text": Set c = Catalog: key = operationId & "@" & version
    If Not c.Exists(key) Then GoTo Broken
    Set s = c(key): typeName = s("ReturnType")
    If args.Count <> s("Arguments").Count Or parameters.Count <> s("Parameters").Count Then GoTo Broken
    For Each prop In parameters.Keys
        If Not s("Parameters").Exists(prop) Then GoTo Broken
        Set parameter = parameters(prop)
        If parameter.Count <> 2 Then GoTo Broken
        If parameter("Type") <> s("Parameters")(prop) Or VarType(parameter("Value")) <> vbString Then GoTo Broken
        If Not AtLiteralValid(CStr(parameter("Type")), CStr(parameter("Value"))) Then GoTo Broken
        If Len(CStr(parameter("Value"))) > 65536 Then GoTo Broken
    Next prop
    ' Validate all envelopes before strict state propagation; no implicit coercion.
    For i = 1 To args.Count
        Set a = args(i)
        If a.ResultType <> s("Arguments")(i) Or Not AtAstState(a.State) Then GoTo Broken
        If Len(a.Value) > 65536 Then GoTo Broken
        If a.State = "OK" Then
            If Not AtLiteralValid(a.ResultType, a.Value) Then GoTo Broken
        ElseIf a.Value <> "" Then
            GoTo Broken
        End If
    Next i
    For Each a In args
        If a.State <> "OK" Then
            Set r = AtEvalValue(typeName, a.State, ""): GoTo Effects
        End If
    Next a
    Set a = args(1)
    Select Case key
        Case "OP_TEXT_EQUAL@1.0.0"
            If parameters("MODE")("Value") <> "binary" Then
                Set r = AtEvalValue(typeName, "INVALID", "")
                r.AddDiagnostic "E_PARAMETER", "Operation", "", "MODE must be binary.": GoTo Effects
            End If
            Set b = args(2): ok = StrComp(a.Value, b.Value, vbBinaryCompare) = 0
        Case "OP_BOOL_EQUAL@1.0.0"
            Set b = args(2): ok = a.Value = b.Value
        Case "OP_INT64_LE@1.0.0"
            Set b = args(2): ok = Int64Compare(a.Value, b.Value) <= 0
        Case "OP_TEXT_TRIM@1.0.0"
            Set r = AtEvalValue("Text", "OK", Trim$(a.Value)): GoTo Effects
        Case Else: GoTo Broken
    End Select
    Set r = AtEvalValue(typeName, "OK", IIf(ok, "true", "false"))
Effects:
    For Each a In args: r.MergeEffects a: Next a
    Set ExecuteOperation = r: Exit Function
Broken:
    If typeName = "" Then typeName = "Text"
    Set r = AtEvalValue(typeName, "ERROR", "")
    r.AddDiagnostic "E_OPERATION", "Operation", "", "Unknown exact operation or invalid typed arguments/parameters."
    Set ExecuteOperation = r
End Function

Private Function Int64Compare(ByVal a As String, ByVal b As String) As Long
    Dim an As Boolean, bn As Boolean, comparison As Long
    an = Left$(a, 1) = "-": bn = Left$(b, 1) = "-"
    If an <> bn Then Int64Compare = IIf(an, -1, 1): Exit Function
    If an Then a = Mid$(a, 2): b = Mid$(b, 2)
    If Len(a) <> Len(b) Then comparison = Sgn(Len(a) - Len(b)) Else comparison = StrComp(a, b, vbBinaryCompare)
    Int64Compare = IIf(an, -comparison, comparison)
End Function
