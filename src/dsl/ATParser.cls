VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATParser"
Option Explicit

Private mTokens As Collection
Private mIndex As Long
Private mResult As ATParseResult
Private mContext As String
Private mRegistry As ATRegistry
Private mDepth As Long

Public Function Parse(ByVal text As String, Optional ByVal contextualFieldId As String = "", Optional ByVal registry As ATRegistry = Nothing) As ATParseResult
    Dim lexed As ATLexResult
    Dim root As ATAstNode
    Dim diagnostic As ATDiagnostic
    Set mResult = New ATParseResult
    mContext = contextualFieldId
    mDepth = 0
    If registry Is Nothing Then Set registry = AT_CreateRegistry()
    Set mRegistry = registry
    Set lexed = AT_Tokenize(text)
    For Each diagnostic In lexed.Diagnostics
        mResult.Diagnostics.Add diagnostic
    Next diagnostic
    If mResult.Diagnostics.Count > 0 Then
        Set Parse = mResult
        Exit Function
    End If
    Set mTokens = lexed.Tokens
    mIndex = 1
    On Error GoTo Failed
    If Len(mContext) > 0 Then mContext = AT_CanonicalId(mContext)
    Set root = ParseOr()
    If Current.Kind <> "EOF" Then Fail "PARSE_TRAILING_TOKEN", "Unexpected trailing token"
    Set mResult.Root = root
    Set Parse = mResult
    Exit Function
Failed:
    If mResult.Diagnostics.Count = 0 Then
        Set diagnostic = New ATDiagnostic
        diagnostic.Initialize "PARSE_INTERNAL", "SYNTAX", "ERROR", Current.Start, Current.Length, Err.Description
        mResult.Diagnostics.Add diagnostic
    End If
    Set mResult.Root = Nothing
    Set Parse = mResult
End Function

Private Property Get Current() As ATToken
    Set Current = mTokens(mIndex)
End Property

Private Function At(ByVal value As String) As Boolean
    At = Current.Value = value And (Current.Kind = "IDENT" Or Current.Kind = "SYMBOL")
End Function

Private Function Take(ByVal value As String) As Boolean
    If At(value) Then
        mIndex = mIndex + 1
        Take = True
    End If
End Function

Private Sub Expect(ByVal value As String)
    If Not Take(value) Then Fail "PARSE_EXPECTED_TOKEN", "Expected " & value
End Sub

Private Sub Fail(ByVal code As String, ByVal message As String)
    Dim diagnostic As New ATDiagnostic
    diagnostic.Initialize code, "SYNTAX", "ERROR", Current.Start, Current.Length, message
    mResult.Diagnostics.Add diagnostic
    Err.Raise vbObjectError + 2200, "ATParser", message
End Sub

Private Function Node(ByVal kind As String, ByVal text As String, ByVal start As Long) As ATAstNode
    Dim result As New ATAstNode
    result.Kind = kind
    result.Text = text
    result.Start = start
    result.Length = Current.Start - start
    Set Node = result
End Function

Private Function Binary(ByVal kind As String, ByVal op As String, ByVal left As ATAstNode, ByVal right As ATAstNode) As ATAstNode
    Dim result As ATAstNode
    Set result = Node(kind, op, left.Start)
    result.Children.Add left
    result.Children.Add right
    Set Binary = result
End Function

Private Function ParseOr() As ATAstNode
    Dim left As ATAstNode
    Set left = ParseAnd()
    Do While Take("OR")
        Set left = Binary("BINARY", "OR", left, ParseAnd())
    Loop
    Set ParseOr = left
End Function

Private Function ParseAnd() As ATAstNode
    Dim left As ATAstNode
    Set left = ParseNot()
    Do While Take("AND")
        Set left = Binary("BINARY", "AND", left, ParseNot())
    Loop
    Set ParseAnd = left
End Function

Private Function ParseNot() As ATAstNode
    Dim start As Long
    Dim child As ATAstNode
    start = Current.Start
    If Take("NOT") Then
        mDepth = mDepth + 1
        If mDepth > 64 Then Fail "LIMIT_DEPTH", "Expression nesting exceeded"
        Set child = ParseNot()
        Set ParseNot = Node("UNARY", "NOT", start)
        ParseNot.Children.Add child
        mDepth = mDepth - 1
    Else
        Set ParseNot = ParseRelation()
    End If
End Function

Private Function IsRelation() As Boolean
    Dim value As String
    value = Current.Value
    If Current.Kind = "SYMBOL" Then
        IsRelation = value = "=" Or value = "!=" Or value = "<>" Or value = "<" Or value = ">" Or value = "<=" Or value = ">="
    ElseIf Current.Kind = "IDENT" Then
        If value = "AND" Or value = "OR" Or value = "NOT" Then Exit Function
        If mRegistry.Contains(value) Then IsRelation = mRegistry.GetPrimitive(value).Category = "OPERATOR"
    End If
End Function

Private Function ParseRelation() As ATAstNode
    Dim left As ATAstNode, right As ATAstNode
    Dim op As String, colon As Boolean, start As Long
    start = Current.Start
    If IsRelation() Then
        op = Current.Value
        mIndex = mIndex + 1
        Expect ":"
        If Len(mContext) = 0 Then Fail "PARSE_CONTEXT_REQUIRED", "Contextual operator requires a FieldId"
        Set left = Node("FIELD", mContext, start)
        Set right = ParseContextValues()
        Set ParseRelation = Binary("RELATION", op, left, right)
        Exit Function
    End If
    Set left = ParseAdditive()
    If IsRelation() Then
        op = Current.Value
        If op = "<>" Then op = "!="
        mIndex = mIndex + 1
        colon = Take(":")
        If colon Then
            Set right = ParseContextValues()
        Else
            Set right = ParseAdditive()
        End If
        Set left = Binary("RELATION", op, left, right)
    End If
    Set ParseRelation = left
End Function

Private Function ParseContextValues() As ATAstNode
    Dim list As ATAstNode, item As ATAstNode
    If At("[") Then
        Set ParseContextValues = ParsePrimary()
        Exit Function
    End If
    Set list = Node("LIST", "", Current.Start)
    Set item = ParseAdditive()
    list.Children.Add item
    Do While At(",")
        If mIndex + 1 > mTokens.Count Then Exit Do
        If mTokens(mIndex + 1).Kind <> "NUMBER" And mTokens(mIndex + 1).Kind <> "STRING" Then Exit Do
        mIndex = mIndex + 1
        Set item = ParseAdditive()
        list.Children.Add item
    Loop
    Set ParseContextValues = list
End Function

Private Function ParseAdditive() As ATAstNode
    Dim left As ATAstNode, op As String
    Set left = ParseMultiplicative()
    Do While At("+") Or At("-")
        op = Current.Value
        mIndex = mIndex + 1
        Set left = Binary("BINARY", op, left, ParseMultiplicative())
    Loop
    Set ParseAdditive = left
End Function

Private Function ParseMultiplicative() As ATAstNode
    Dim left As ATAstNode, op As String
    Set left = ParseUnary()
    Do While At("*") Or At("/") Or At("%")
        op = Current.Value
        mIndex = mIndex + 1
        Set left = Binary("BINARY", op, left, ParseUnary())
    Loop
    Set ParseMultiplicative = left
End Function

Private Function ParseUnary() As ATAstNode
    Dim result As ATAstNode, op As String, start As Long
    If At("+") Or At("-") Then
        start = Current.Start
        op = Current.Value
        mIndex = mIndex + 1
        Set result = Node("UNARY", op, start)
        mDepth = mDepth + 1
        If mDepth > 64 Then Fail "LIMIT_DEPTH", "Expression nesting exceeded"
        result.Children.Add ParseUnary()
        mDepth = mDepth - 1
        Set ParseUnary = result
    Else
        Set ParseUnary = ParsePrimary()
    End If
End Function

Private Function ParsePrimary() As ATAstNode
    Dim result As ATAstNode, child As ATAstNode, binding As ATAstNode
    Dim kind As String, value As String, start As Long, named As Boolean
    start = Current.Start
    mDepth = mDepth + 1
    If mDepth > 64 Then Fail "LIMIT_DEPTH", "Expression nesting exceeded"
    If Take("(") Then
        Set result = ParseOr()
        Expect ")"
    ElseIf Take("[") Then
        Set result = Node("LIST", "", start)
        If Not Take("]") Then
            Do
                result.Children.Add ParseOr()
                If Not Take(",") Then Exit Do
            Loop
            Expect "]"
        End If
    ElseIf Take("@") Then
        value = ReadIdentifier()
        If Take(".") Then value = value & "." & ReadIdentifier()
        Set result = Node("FIELD", value, start)
    ElseIf Current.Kind = "NUMBER" Or Current.Kind = "STRING" Then
        kind = Current.Kind
        value = Current.Value
        mIndex = mIndex + 1
        Set result = Node(kind, value, start)
    ElseIf At("TRUE") Or At("FALSE") Then
        value = Current.Value
        mIndex = mIndex + 1
        Set result = Node("BOOL", value, start)
    Else
        named = Take("$")
        If Current.Kind <> "IDENT" Then Fail "PARSE_EXPECTED_VALUE", "Expected literal, reference or function call"
        value = ReadIdentifier()
        If named Then
            Set result = Node("NAMED", value, start)
        Else
            Set result = Node("CALL", value, start)
        End If
        If Take("(") Then
            If Not Take(")") Then
                Do
                    Set child = ParseOr()
                    If Take("AS") Then
                        Set binding = Node("BIND", ReadIdentifier(), child.Start)
                        binding.Children.Add child
                        Set child = binding
                    End If
                    result.Children.Add child
                    If Not Take(",") Then Exit Do
                Loop
                Expect ")"
            End If
        ElseIf Not named Then
            Fail "PARSE_BARE_IDENTIFIER", "Bare identifiers are not string literals"
        End If
    End If
    mDepth = mDepth - 1
    result.Length = mTokens(mIndex - 1).Start + mTokens(mIndex - 1).Length - result.Start
    Set ParsePrimary = result
End Function

Private Function ReadIdentifier() As String
    If Current.Kind <> "IDENT" Then Fail "PARSE_EXPECTED_ID", "Expected identifier"
    Select Case Current.Value
        Case "AND", "OR", "NOT", "AS", "TRUE", "FALSE"
            Fail "PARSE_RESERVED_ID", "Reserved keyword cannot be an identifier"
    End Select
    ReadIdentifier = Current.Value
    mIndex = mIndex + 1
End Function

