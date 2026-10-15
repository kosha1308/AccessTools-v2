Attribute VB_Name = "modATContracts"
Option Explicit

Public Enum ATEvalState
    atEvalOK = 0
    atEvalError = 1
    atEvalSkipped = 2
    atEvalManualReview = 3
End Enum

Public Enum ATValueState
    atValuePresent = 0
    atValueEmpty = 1
    atValueMissing = 2
    atValueInvalid = 3
End Enum

Public Function AT_CanonicalId(ByVal value As String) As String
    Dim i As Long
    Dim code As Long
    If Len(value) = 0 Then Err.Raise vbObjectError + 2140, "ATContracts", "REG_INVALID_ID: empty"
    For i = 1 To Len(value)
        code = AscW(Mid$(value, i, 1))
        If code >= 97 And code <= 122 Then code = code - 32
        If Not ((code >= 65 And code <= 90) Or (i > 1 And ((code >= 48 And code <= 57) Or code = 95))) Then
            Err.Raise vbObjectError + 2140, "ATContracts", "REG_INVALID_ID: " & value
        End If
        AT_CanonicalId = AT_CanonicalId & ChrW$(code)
    Next i
End Function

Public Function AT_MakePrimitive(ByVal idValue As String, ByVal categoryValue As String, ByVal argumentTypes As String, ByVal resultType As String, ByVal example As String) As ATPrimitive
    Dim item As New ATPrimitive
    item.Initialize idValue, categoryValue, idValue, idValue, "1.0.0", "ACTIVE", "PURE", example, "positive; negative; empty; boundary; malformed"
    item.AddSignature argumentTypes, resultType
    Set AT_MakePrimitive = item
End Function

Public Function AT_CreateRegistry() As ATRegistry
    Dim registry As New ATRegistry
    Dim item As Variant
    Dim primitive As ATPrimitive
    For Each item In Array("TEXT", "NUMBER", "BOOL", "DATE", "DATETIME", "EMAIL", "PROTOCOL", "HOSTNAME", "PORT", "PORT_SET", "IP", "IP_EXPRESSION")
        registry.Register AT_MakePrimitive(CStr(item), "DATATYPE", "", CStr(item), CStr(item))
    Next item
    For Each item In Array("ANY", "NONE", "COVER", "CONTAINS", "EXACT", "OVERLAPS")
        Set primitive = AT_MakePrimitive(CStr(item), "OPERATOR", "PORT_SET,PORT_SET", "BOOL", "@PORTS " & CStr(item) & ": 443")
        primitive.AddSignature "IP_EXPRESSION,IP_EXPRESSION", "BOOL"
        registry.Register primitive
    Next item
    registry.Register AT_MakePrimitive("LOWER", "FUNCTION", "TEXT", "TEXT", "LOWER(@NAME)")
    Set AT_CreateRegistry = registry
End Function

