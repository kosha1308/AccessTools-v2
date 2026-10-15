Attribute VB_Name = "modATLiteralContracts"
Option Explicit

Public Function AT_CanonicalDecimal(ByVal value As String) As String
    Dim parts As Variant, whole As String, fraction As String
    Dim first As Long, last As Long, i As Long, ch As String
    If Len(value) = 0 Then Err.Raise vbObjectError + 2450, "ATLiteral", "NUMBER_INVALID_LITERAL"
    parts = Split(value, ".")
    If UBound(parts) > 1 Then Err.Raise vbObjectError + 2450, "ATLiteral", "NUMBER_INVALID_LITERAL"
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch <> "." And (ch < "0" Or ch > "9") Then Err.Raise vbObjectError + 2450, "ATLiteral", "NUMBER_INVALID_LITERAL"
    Next i
    whole = parts(0)
    If Len(whole) = 0 Then Err.Raise vbObjectError + 2450, "ATLiteral", "NUMBER_INVALID_LITERAL"
    first = 1
    Do While first < Len(whole)
        If Mid$(whole, first, 1) <> "0" Then Exit Do
        first = first + 1
    Loop
    whole = Mid$(whole, first)
    If UBound(parts) = 1 Then
        fraction = parts(1)
        If Len(fraction) = 0 Then Err.Raise vbObjectError + 2450, "ATLiteral", "NUMBER_INVALID_LITERAL"
        last = Len(fraction)
        Do While last > 0
            If Mid$(fraction, last, 1) <> "0" Then Exit Do
            last = last - 1
        Loop
        fraction = Left$(fraction, last)
    End If
    AT_CanonicalDecimal = whole
    If Len(fraction) > 0 Then AT_CanonicalDecimal = whole & "." & fraction
End Function

Public Sub AT_ValidatePortLiteral(ByVal node As ATAstNode)
    Dim child As ATAstNode, value As String
    Select Case node.Kind
        Case "LIST"
            For Each child In node.Children
                AT_ValidatePortLiteral child
            Next child
        Case "NUMBER"
            value = AT_CanonicalDecimal(node.Text)
            If InStr(value, ".") > 0 Or Len(value) > 5 Then Err.Raise vbObjectError + 2451, "ATLiteral", "PORT_INVALID_LITERAL"
            If CLng(value) > 65535 Then Err.Raise vbObjectError + 2451, "ATLiteral", "PORT_INVALID_LITERAL"
        Case "UNARY"
            If node.Text = "-" Then
                If IsNumericLiteral(node.Children(1)) Then Err.Raise vbObjectError + 2451, "ATLiteral", "PORT_INVALID_LITERAL"
            ElseIf node.Text = "+" Then
                AT_ValidatePortLiteral node.Children(1)
            End If
    End Select
End Sub

Private Function IsNumericLiteral(ByVal node As ATAstNode) As Boolean
    If node.Kind = "NUMBER" Then
        IsNumericLiteral = True
    ElseIf node.Kind = "UNARY" And (node.Text = "+" Or node.Text = "-") Then
        IsNumericLiteral = IsNumericLiteral(node.Children(1))
    End If
End Function
