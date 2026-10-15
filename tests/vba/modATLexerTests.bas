Attribute VB_Name = "modATLexerTests"
Option Explicit

Public Sub AT_TestLexer()
    Dim result As ATLexResult
    Dim token As ATToken
    Dim quote As String
    Dim text As String
    quote = Chr$(34)
    Set result = AT_Tokenize("@t.FIELD >= 12.5 AND $RULE([443,8443])")
    AT_Assert result.Diagnostics.Count = 0, "lex references numbers calls"
    AT_Assert result.Tokens.Count = 17, "lex token count including EOF"
    Set token = result.Tokens(6)
    AT_Assert token.Kind = "NUMBER" And token.Value = "12.5", "lex decimal locale independent"
    Set token = result.Tokens(2)
    AT_Assert token.Value = "T" And token.Raw = "t" And token.Start = 2, "lex raw and canonical identifier"
    text = quote & ChrW$(&H416) & "\" & quote & "\\" & "\n" & quote
    Set result = AT_Tokenize(text)
    Set token = result.Tokens(1)
    AT_Assert token.Value = ChrW$(&H416) & quote & "\" & vbLf, "lex Unicode and escapes"
    AT_Assert token.Length = Len(text), "lex string source span"
    Set result = AT_Tokenize("true" & vbCrLf & "FALSE")
    Set token = result.Tokens(2)
    AT_Assert token.Value = "FALSE" And token.Start = 7, "lex CRLF offsets"
    Set result = AT_Tokenize("DATE(" & quote & "2026-10-15" & quote & ")")
    AT_Assert result.Diagnostics.Count = 0, "lex explicit date call"
    Set result = AT_Tokenize(quote & "unterminated")
    AT_Assert result.Diagnostics.Count = 1, "lex unterminated string rejected"
    Set result = AT_Tokenize("1 + 2 ?")
    AT_Assert result.Diagnostics.Count = 1, "lex malformed suffix rejected"
    Set result = AT_Tokenize(quote & "\q" & quote)
    AT_Assert result.Diagnostics.Count = 1, "lex unknown escape rejected"
    Set result = AT_Tokenize("1,5")
    AT_Assert result.Tokens.Count = 4, "lex comma is delimiter not decimal separator"
    Set result = AT_Tokenize("!= <> <= >= + - * / % = : . @ $")
    AT_Assert result.Diagnostics.Count = 0, "lex punctuation"
End Sub
