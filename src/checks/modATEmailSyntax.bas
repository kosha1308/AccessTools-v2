Attribute VB_Name = "modATEmailSyntax"
Option Explicit
Public Function AT_EmailSyntaxValid(ByVal text As String) As Boolean
    Dim at As Long, dot As Long, i As Long, ch As String
    If Len(text) > 65536 Then Exit Function
    text = AT_AsciiCase(Trim$(text), True)
    at = InStr(text, "@")
    If at < 2 Or at = Len(text) Then Exit Function
    If InStr(at + 1, text, "@") > 0 Then Exit Function
    For i = 1 To at - 1
        ch = Mid$(text, i, 1)
        If Not AlphaNumeric(ch) Then
            If InStr(1, "._%+-", ch, vbBinaryCompare) = 0 Then Exit Function
        End If
    Next i
    dot = InStrRev(text, ".")
    If dot < at + 2 Or Len(text) - dot < 2 Then Exit Function
    For i = at + 1 To dot - 1
        ch = Mid$(text, i, 1)
        If Not AlphaNumeric(ch) And ch <> "." And ch <> "-" Then Exit Function
    Next i
    For i = dot + 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch < "A" Or ch > "Z" Then Exit Function
    Next i
    AT_EmailSyntaxValid = True
End Function
Private Function AlphaNumeric(ByVal ch As String) As Boolean
    AlphaNumeric = ((ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9"))
End Function
