Attribute VB_Name = "modATLanguageDocs"
Option Explicit

Public Function AT_ExportLanguageCatalog() As String
    Dim registry As ATRegistry, primitive As ATPrimitive, signature As ATSignature
    Dim argument As Variant, signatures As String, arguments As String, result As String
    Set registry = AT_CreateRegistry()
    result = "# AccessTools DSL 1.0.0 language catalog" & vbLf & vbLf
    result = result & "Generated from AT_CreateRegistry; edit registry metadata, then regenerate this file." & vbLf & vbLf
    result = result & "These are compiler signatures and versioned evaluator/normalizer binding contracts. Runtime evaluators are not enabled by this catalog." & vbLf & vbLf
    For Each primitive In registry.Entries
        result = result & "## " & primitive.Id & vbLf & vbLf
        result = result & "- Category: " & primitive.Category & vbLf
        result = result & "- Display name: " & EscapeMarkdown(primitive.DisplayName) & vbLf
        result = result & "- Description: " & EscapeMarkdown(primitive.Description) & vbLf
        result = result & "- Syntax: " & EscapeMarkdown(primitive.Syntax) & vbLf
        result = result & "- Canonical ID/syntax: " & primitive.CanonicalId & " / " & EscapeMarkdown(primitive.CanonicalSyntax) & vbLf
        result = result & "- Negated alias: " & IIf(primitive.NegatedAlias, "TRUE", "FALSE") & vbLf
        signatures = ""
        For Each signature In primitive.Signatures
            arguments = ""
            For Each argument In signature.ArgumentTypes
                If Len(arguments) > 0 Then arguments = arguments & ", "
                arguments = arguments & CStr(argument)
            Next argument
            If Len(signatures) > 0 Then signatures = signatures & "; "
            signatures = signatures & "(" & arguments & ") -> " & signature.ReturnType
        Next signature
        result = result & "- Signatures: " & EscapeMarkdown(signatures) & vbLf
        result = result & "- Evaluator binding: " & primitive.EvaluatorId & vbLf
        result = result & "- Version/status: " & primitive.Version & " / " & primitive.Status & vbLf
        result = result & "- Determinism/cache: " & primitive.Determinism & " / " & primitive.CachePolicy & vbLf
        result = result & "- Precedence/associativity: " & CStr(primitive.Precedence) & " / " & primitive.Associativity & vbLf
        If primitive.Category = "DATATYPE" Then
            result = result & "- Normalizer binding: " & primitive.NormalizerId & vbLf
            result = result & "- Literal policy: " & primitive.LiteralPolicy & vbLf
            result = result & "- Validation policy: " & primitive.ValidationPolicy & vbLf
        End If
        result = result & "- Example: " & EscapeMarkdown(primitive.Examples) & vbLf
        result = result & "- Test contract: " & EscapeMarkdown(primitive.TestContract) & vbLf & vbLf
    Next primitive
    AT_ExportLanguageCatalog = result
End Function

Private Function EscapeMarkdown(ByVal value As String) As String
    value = Replace(Replace(Replace(value, "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
    value = Replace(Replace(Replace(value, vbCrLf, " "), vbCr, " "), vbLf, " ")
    EscapeMarkdown = value
End Function
