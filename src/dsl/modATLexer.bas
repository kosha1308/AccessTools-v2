Attribute VB_Name = "modATLexer"
Option Explicit

Public Function AT_Tokenize(ByVal text As String, Optional ByVal limits As ATLimits = Nothing) As ATLexResult
    Dim result As New ATLexResult
    Dim i As Long, first As Long, count As Long
    Dim ch As String, pair As String, value As String, raw As String
    Dim closed As Boolean, badEscape As Boolean
    Dim kind As String
    Dim diagnostic As ATDiagnostic
    On Error GoTo Failed
    If limits Is Nothing Then Set limits = New ATLimits
    limits.Validate
    If Len(text) > limits.MaxInputLength Then
        AT_LexError result, "LIMIT_INPUT_LENGTH", 1, Len(text), "Input length budget exceeded"
        Set AT_Tokenize = result
        Exit Function
    End If
    i = 1
    count = Len(text)
    Do While i <= count
        If result.Tokens.Count >= limits.MaxTokens Then
            AT_LexError result, "LIMIT_TOKENS", i, 0, "Token budget exceeded"
            Set AT_Tokenize = result
            Exit Function
        End If
        ch = Mid$(text, i, 1)
        first = i
        Select Case ch
            Case " ", vbTab, vbCr, vbLf
                i = i + 1
            Case Chr$(34)
                value = ""
                closed = False
                badEscape = False
                i = i + 1
                Do While i <= count
                    ch = Mid$(text, i, 1)
                    If ch = Chr$(34) Then
                        closed = True
                        i = i + 1
                        Exit Do
                    ElseIf ch = "\" Then
                        i = i + 1
                        If i > count Then Exit Do
                        ch = Mid$(text, i, 1)
                        Select Case ch
                            Case Chr$(34), "\": value = value & ch
                            Case "n": value = value & vbLf
                            Case "r": value = value & vbCr
                            Case "t": value = value & vbTab
                            Case Else: badEscape = True
                        End Select
                    Else
                        value = value & ch
                    End If
                    i = i + 1
                Loop
                If Not closed Then
                    AT_LexError result, "LEX_UNTERMINATED_STRING", first, count - first + 1, "Unterminated string literal"
                ElseIf badEscape Then
                    AT_LexError result, "LEX_INVALID_ESCAPE", first, i - first, "Unsupported string escape"
                Else
                    AT_AddToken result, "STRING", Mid$(text, first, i - first), value, first
                End If
            Case Else
                If AT_IsLetter(ch) Then
                    i = i + 1
                    Do While i <= count
                        ch = Mid$(text, i, 1)
                        If Not (AT_IsLetter(ch) Or AT_IsDigit(ch) Or ch = "_") Then Exit Do
                        i = i + 1
                    Loop
                    raw = Mid$(text, first, i - first)
                    AT_AddToken result, "IDENT", raw, AT_CanonicalId(raw), first
                ElseIf AT_IsDigit(ch) Then
                    i = i + 1
                    Do While i <= count
                        If Not AT_IsDigit(Mid$(text, i, 1)) Then Exit Do
                        i = i + 1
                    Loop
                    If Mid$(text, i, 1) = "." And AT_IsDigit(Mid$(text, i + 1, 1)) Then
                        i = i + 1
                        Do While i <= count
                            If Not AT_IsDigit(Mid$(text, i, 1)) Then Exit Do
                            i = i + 1
                        Loop
                    End If
                    raw = Mid$(text, first, i - first)
                    AT_AddToken result, "NUMBER", raw, raw, first
                Else
                    pair = Mid$(text, i, 2)
                    If pair = "!=" Or pair = "<>" Or pair = ">=" Or pair = "<=" Then
                        AT_AddToken result, "SYMBOL", pair, pair, first
                        i = i + 2
                    ElseIf InStr(1, "@$.,:()[]+-*/%=<>", ch, vbBinaryCompare) > 0 Then
                        AT_AddToken result, "SYMBOL", ch, ch, first
                        i = i + 1
                    Else
                        AT_LexError result, "LEX_UNEXPECTED_CHARACTER", first, 1, "Unexpected character"
                        i = i + 1
                    End If
                End If
        End Select
    Loop
    If result.Tokens.Count >= limits.MaxTokens Then
        AT_LexError result, "LIMIT_TOKENS", count + 1, 0, "Token budget exceeded"
    Else
        AT_AddToken result, "EOF", "", "", count + 1
    End If
    Set AT_Tokenize = result
    Exit Function
Failed:
    Set diagnostic = New ATDiagnostic
    diagnostic.Initialize "LEX_CONFIGURATION", "LEXICAL", "ERROR", 1, 0, Err.Description
    result.Diagnostics.Add diagnostic
    Set AT_Tokenize = result
End Function

Private Sub AT_AddToken(ByVal result As ATLexResult, ByVal kind As String, ByVal raw As String, ByVal value As String, ByVal first As Long)
    Dim token As New ATToken
    token.Kind = kind
    token.Raw = raw
    token.Value = value
    token.Start = first
    token.Length = Len(raw)
    result.Tokens.Add token
End Sub

Private Sub AT_LexError(ByVal result As ATLexResult, ByVal code As String, ByVal first As Long, ByVal length As Long, ByVal message As String)
    Dim diagnostic As New ATDiagnostic
    diagnostic.Initialize code, "LEXICAL", "ERROR", first, length, message
    result.Diagnostics.Add diagnostic
End Sub

Private Function AT_IsLetter(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    AT_IsLetter = (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z")
End Function

Private Function AT_IsDigit(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    AT_IsDigit = (ch >= "0" And ch <= "9")
End Function

