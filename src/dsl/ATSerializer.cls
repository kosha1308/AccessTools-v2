VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATSerializer"
Option Explicit

Public Function Serialize(ByVal node As ATAstNode) As String
    Dim result As String
    Dim child As ATAstNode
    If node Is Nothing Then Err.Raise vbObjectError + 2210, "ATSerializer", "SERIALIZE_MISSING_ROOT"
    Select Case node.Kind
        Case "NUMBER", "BOOL"
            result = node.Text
        Case "STRING"
            result = Replace(node.Text, "\", "\\")
            result = Replace(result, Chr$(34), "\" & Chr$(34))
            result = Replace(result, vbCr, "\r")
            result = Replace(result, vbLf, "\n")
            result = Replace(result, vbTab, "\t")
            result = Chr$(34) & result & Chr$(34)
        Case "FIELD"
            result = "@" & node.Text
        Case "UNARY"
            result = "(" & node.Text & " " & Serialize(node.Children(1)) & ")"
        Case "BINARY", "RELATION"
            result = "(" & Serialize(node.Children(1)) & " " & node.Text & " " & Serialize(node.Children(2)) & ")"
        Case "BIND"
            result = Serialize(node.Children(1)) & " AS " & node.Text
        Case "LIST", "CALL", "NAMED"
            For Each child In node.Children
                If Len(result) > 0 Then result = result & ","
                result = result & Serialize(child)
            Next child
            If node.Kind = "LIST" Then
                result = "[" & result & "]"
            ElseIf node.Kind = "NAMED" Then
                result = "$" & node.Text & "(" & result & ")"
            Else
                result = node.Text & "(" & result & ")"
            End If
        Case Else
            Err.Raise vbObjectError + 2211, "ATSerializer", "SERIALIZE_UNKNOWN_NODE"
    End Select
    Serialize = result
End Function

