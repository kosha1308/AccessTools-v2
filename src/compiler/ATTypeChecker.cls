VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATTypeChecker"
Option Explicit
Public LastSignature As ATSignature
Public LastContextWrap As Long

Public Function ResolveSignature(ByVal primitive As ATPrimitive, ByVal types As Collection, ByVal registry As ATRegistry, Optional ByVal contextualIndex As Long = 0) As String
    Dim signature As ATSignature, parameters As Collection
    Dim index As Long, binding As String, valid As Boolean, wrapped As Boolean, wrapIndex As Long
    Dim conversionPass As Long
    Set LastSignature = Nothing
    LastContextWrap = 0
    For conversionPass = 0 To 1
    For Each signature In primitive.Signatures
        Set parameters = signature.ArgumentTypes
        If parameters.Count = types.Count Then
            valid = True
            binding = ""
            wrapIndex = 0
            For index = 1 To types.Count
                wrapped = False
                If Not MatchType(CStr(parameters(index)), CStr(types(index)), binding, registry, index = contextualIndex, wrapped, conversionPass = 1) Then valid = False
                If wrapped Then wrapIndex = index
            Next index
            If valid Then
                Set LastSignature = signature
                LastContextWrap = wrapIndex
                If signature.ReturnType = "T" Then
                    ResolveSignature = binding
                Else
                    ResolveSignature = Replace(signature.ReturnType, "<T>", "<" & binding & ">")
                End If
                Exit Function
            End If
        End If
    Next signature
    Next conversionPass
    Err.Raise vbObjectError + 2320, "ATTypeChecker", "TYPE_SIGNATURE_MISMATCH: " & primitive.Id
End Function

Private Function MatchType(ByVal expected As String, ByVal actual As String, ByRef binding As String, ByVal registry As ATRegistry, ByVal allowSingleton As Boolean, ByRef wrapped As Boolean, ByVal allowConversions As Boolean) As Boolean
    Dim marker As Long, prefix As String, suffix As String, extracted As String
    Dim ignored As Boolean
    If allowSingleton And Left$(expected, 5) = "LIST<" And Left$(actual, 5) <> "LIST<" Then
        MatchType = MatchType(Mid$(expected, 6, Len(expected) - 6), actual, binding, registry, False, ignored, allowConversions)
        wrapped = MatchType
    ElseIf expected = actual And expected <> "T" And InStr(expected, "<T>") = 0 Then
        MatchType = True
    ElseIf expected = "T" Then
        If Len(binding) = 0 Then binding = actual
        MatchType = binding = actual
    Else
        marker = InStr(expected, "<T>")
        If marker > 0 Then
            prefix = Left$(expected, marker)
            suffix = Mid$(expected, marker + 2)
            If Left$(actual, Len(prefix)) = prefix And Right$(actual, Len(suffix)) = suffix And Len(actual) > Len(prefix) + Len(suffix) Then
                extracted = Mid$(actual, Len(prefix) + 1, Len(actual) - Len(prefix) - Len(suffix))
                If Len(binding) = 0 Then binding = extracted
                MatchType = binding = extracted
            End If
        Else
            If allowConversions Then MatchType = registry.CanConvert(actual, expected)
        End If
    End If
End Function

