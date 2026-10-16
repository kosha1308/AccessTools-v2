Attribute VB_Name = "modATRuntimeNormalize"
Option Explicit

Public Function AT_RuntimeInvalid(ByVal typeName As String, ByVal code As String, ByVal node As ATAstNode) As ATEvalResult
    Dim result As New ATEvalResult, diagnostics As New Collection
    diagnostics.Add AT_RuntimeDiagnostic(code, node)
    result.Initialize atEvalOK, AT_RuntimeState(typeName, atValueInvalid), diagnostics, Nothing
    Set AT_RuntimeInvalid = result
End Function

Public Function AT_ParseDecimal(ByVal text As String, ByVal node As ATAstNode) As ATEvalResult
    Dim canonical As String, value As Double, sign As Double
    On Error GoTo Invalid
    sign = 1#
    If Left$(text, 1) = "-" Or Left$(text, 1) = "+" Then
        If Left$(text, 1) = "-" Then sign = -1#
        text = Mid$(text, 2)
    End If
    canonical = AT_CanonicalDecimal(text)
    value = Val(canonical) * sign
    If value = 0# And canonical <> "0" Then GoTo Invalid
    Set AT_ParseDecimal = AT_RuntimeOK(AT_RuntimeScalar("NUMBER", value))
    Exit Function
Invalid:
    Set AT_ParseDecimal = AT_RuntimeInvalid("NUMBER", "RUNTIME_NUMBER_RANGE", node)
End Function

Private Function NormalizeNumericPort(ByVal typeName As String, ByVal raw As Variant, ByVal node As ATAstNode) As ATEvalResult
    Dim number As ATEvalResult, converted As ATEvalResult, budget As New ATEvalBudget, limits As New ATEvalLimits, value As New ATRuntimeValue
    If typeName <> "PORT" And typeName <> "PORT_SET" Then
        Set NormalizeNumericPort = AT_RuntimeInvalid(typeName, "DOMAIN_INVALID_IPV4", node)
        Exit Function
    End If
    Set number = AT_NormalizeScalar("NUMBER", raw, node)
    If number.Value.ValueState <> atValuePresent Then
        Set NormalizeNumericPort = AT_RuntimeInvalid(typeName, "DOMAIN_INVALID_PORT", node)
        Exit Function
    End If
    budget.Initialize limits
    Set converted = AT_ConvertPortValue(number.Value, node, budget)
    If converted.State <> atEvalOK Then
        Set NormalizeNumericPort = AT_RuntimeInvalid(typeName, "DOMAIN_INVALID_PORT", node)
    Else
        value.InitializeDomain typeName, converted.Value.Domain, converted.Value.DeclarationKeys
        Set NormalizeNumericPort = AT_RuntimeOK(value)
    End If
End Function

Public Function AT_NormalizeScalar(ByVal typeName As String, ByVal raw As Variant, ByVal node As ATAstNode) As ATEvalResult
    Dim value As Double, dateValue As Date
    On Error GoTo Invalid
    If AT_RuntimeDomainType(typeName) Then
        If VarType(raw) = vbString Then
            Set AT_NormalizeScalar = AT_ParseDomain(typeName, CStr(raw), node)
        Else
            Set AT_NormalizeScalar = NormalizeNumericPort(typeName, raw, node)
        End If
        Exit Function
    End If
    Select Case typeName
        Case "TEXT"
            If VarType(raw) <> vbString Then GoTo Invalid
        Case "BOOL"
            If VarType(raw) <> vbBoolean Then GoTo Invalid
        Case "NUMBER"
            If VarType(raw) = vbString Then
                Set AT_NormalizeScalar = AT_ParseDecimal(CStr(raw), node)
                Exit Function
            End If
            Select Case VarType(raw)
                Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal: value = CDbl(raw)
                Case Else: GoTo Invalid
            End Select
            Set AT_NormalizeScalar = AT_RuntimeOK(AT_RuntimeScalar(typeName, value))
            Exit Function
        Case "DATE"
            If VarType(raw) = vbString Then
                dateValue = AT_GregorianDate(CStr(raw))
            ElseIf VarType(raw) = vbDate Then
                value = CDbl(raw)
                If value <> Fix(value) Then GoTo Invalid
                dateValue = raw
            Else
                GoTo Invalid
            End If
            Set AT_NormalizeScalar = AT_RuntimeOK(AT_RuntimeScalar(typeName, dateValue))
            Exit Function
        Case Else
            Set AT_NormalizeScalar = AT_RuntimeFailure(atEvalError, "RUNTIME_UNSUPPORTED_NORMALIZER", node)
            Exit Function
    End Select
    Set AT_NormalizeScalar = AT_RuntimeOK(AT_RuntimeScalar(typeName, raw))
    Exit Function
Invalid:
    Set AT_NormalizeScalar = AT_RuntimeInvalid(typeName, "RUNTIME_INVALID_VALUE", node)
End Function

Public Function AT_GregorianDate(ByVal text As String) As Date
    Dim y As Long, m As Long, d As Long, days As Long, i As Long, ordinal As Long
    Dim months As Variant
    If Len(text) <> 10 Then GoTo Invalid
    If Not text Like "####-##-##" Then GoTo Invalid
    y = CLng(Left$(text, 4)): m = CLng(Mid$(text, 6, 2)): d = CLng(Right$(text, 2))
    If y < 100 Or y > 9999 Or m < 1 Or m > 12 Or d < 1 Then GoTo Invalid
    months = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    If y Mod 400 = 0 Or (y Mod 4 = 0 And y Mod 100 <> 0) Then months(1) = 29
    days = months(m - 1)
    If d > days Then GoTo Invalid
    ordinal = 365 * (y - 1) + (y - 1) \ 4 - (y - 1) \ 100 + (y - 1) \ 400 + d
    For i = 0 To m - 2
        ordinal = ordinal + months(i)
    Next i
    AT_GregorianDate = CDate(CDbl(ordinal - 693594))
    Exit Function
Invalid:
    Err.Raise vbObjectError + 2510, , "RUNTIME_INVALID_DATE"
End Function

Public Function AT_AsciiCase(ByVal text As String, ByVal upper As Boolean) As String
    Dim i As Long, code As Long, output As String
    output = text
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1))
        If upper Then
            If code >= 97 And code <= 122 Then Mid$(output, i, 1) = ChrW$(code - 32)
        Else
            If code >= 65 And code <= 90 Then Mid$(output, i, 1) = ChrW$(code + 32)
        End If
    Next i
    AT_AsciiCase = output
End Function
Public Function AT_TextKey(ByVal text As String, ByVal comparisonMode As String) As String
    Select Case comparisonMode
        Case "BINARY": AT_TextKey = text
        Case "ASCII_TRIM_CI": AT_TextKey = AT_AsciiCase(Trim$(text), False)
        Case Else: Err.Raise vbObjectError + 2511, , "CONFIG_UNSUPPORTED"
    End Select
End Function
