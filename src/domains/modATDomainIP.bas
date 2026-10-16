Attribute VB_Name = "modATDomainIP"
Option Explicit

Public Function AT_ParseIPv4(ByVal text As String) As Double
    Dim parts As Variant, i As Long, octet As Double, value As Double
    parts = Split(text, ".")
    If UBound(parts) <> 3 Then GoTo Invalid
    For i = 0 To 3
        If Not AT_Digits(CStr(parts(i)), 3) Then GoTo Invalid
        octet = Val(CStr(parts(i)))
        If octet > 255# Then GoTo Invalid
        value = value * 256# + octet
    Next i
    AT_ParseIPv4 = value
    Exit Function
Invalid:
    Err.Raise vbObjectError + 2611, , "DOMAIN_INVALID_IPV4"
End Function
Public Function AT_IPText(ByVal number As Double) As String
    Dim factor As Double, octet As Long, i As Long
    If number < 0# Or number > 4294967295# Or number <> Fix(number) Then Err.Raise vbObjectError + 2611, , "DOMAIN_INVALID_IPV4"
    factor = 16777216#
    For i = 1 To 4
        octet = CLng(Fix(number / factor))
        If i > 1 Then AT_IPText = AT_IPText & "."
        AT_IPText = AT_IPText & CStr(octet)
        number = number - octet * factor
        factor = factor / 256#
    Next i
End Function

Public Sub AT_ParseIPToken(ByVal token As String, ByVal intervals As ATIntervalSet, ByVal keys As Collection, ByVal budget As ATEvalBudget, ByRef prefix As String)
    Dim parts As Variant, low As Double, high As Double, tmp As Double, bits As Long, block As Double, base As String
    Dim i As Long, octets As Variant, canonical As String, endpoint As String
    If InStr(token, "/") > 0 Then
        parts = Split(token, "/")
        If UBound(parts) <> 1 Then GoTo Invalid
        low = AT_ParseIPv4(CStr(parts(0)))
        If Not AT_Digits(CStr(parts(1)), 2) Then GoTo Invalid
        bits = CLng(parts(1))
        If bits > 32 Then GoTo Invalid
        base = AT_IPText(low)
        block = 2# ^ (32 - bits)
        low = Fix(low / block) * block
        intervals.Add low, low + block - 1#, budget
        AT_AddDomainKey keys, "CIDR:" & base & "/" & CStr(bits), budget
        prefix = Left$(base, InStrRev(base, "."))
    ElseIf InStr(token, "-") > 0 Then
        parts = Split(token, "-")
        If UBound(parts) <> 1 Then GoTo Invalid
        endpoint = CStr(parts(0))
        If AT_Digits(endpoint, 3) And Len(prefix) > 0 Then endpoint = prefix & endpoint
        low = AT_ParseIPv4(endpoint)
        base = AT_IPText(low)
        endpoint = CStr(parts(1))
        If AT_Digits(endpoint, 3) Then endpoint = Left$(base, InStrRev(base, ".")) & endpoint
        high = AT_ParseIPv4(endpoint)
        prefix = Left$(endpoint, InStrRev(endpoint, "."))
        If low > high Then
            tmp = low: low = high: high = tmp
        End If
        AT_AddDomainRange intervals, keys, low, high, budget
    ElseIf InStr(1, token, "x", vbTextCompare) > 0 Then
        octets = Split(AT_AsciiCase(token, False), ".")
        If UBound(octets) <> 3 Then GoTo Invalid
        For i = 0 To 3
            If octets(i) <> "x" Then
                If Not AT_Digits(CStr(octets(i)), 3) Then GoTo Invalid
                If Val(CStr(octets(i))) > 255# Then GoTo Invalid
            End If
        Next i
        ExpandMask octets, 0, 0#, intervals, keys, budget
        prefix = ""
    Else
        If AT_Digits(token, 3) And Len(prefix) > 0 Then token = prefix & token
        low = AT_ParseIPv4(token)
        base = AT_IPText(low)
        prefix = Left$(base, InStrRev(base, "."))
        AT_AddDomainRange intervals, keys, low, low, budget
    End If
    Exit Sub
Invalid:
    Err.Raise vbObjectError + 2611, , "DOMAIN_INVALID_IPV4"
End Sub

Private Sub ExpandMask(ByVal octets As Variant, ByVal position As Long, ByVal prefix As Double, ByVal intervals As ATIntervalSet, ByVal keys As Collection, ByVal budget As ATEvalBudget)
    Dim allWild As Boolean, i As Long, count As Double
    budget.Step
    If position = 4 Then
        AT_AddDomainRange intervals, keys, prefix, prefix, budget
        Exit Sub
    End If
    allWild = True
    For i = position To 3
        If octets(i) <> "x" Then allWild = False
    Next i
    If allWild Then
        count = 256# ^ (4 - position)
        AT_AddDomainRange intervals, keys, prefix * count, prefix * count + count - 1#, budget
    ElseIf octets(position) = "x" Then
        For i = 0 To 255
            ExpandMask octets, position + 1, prefix * 256# + i, intervals, keys, budget
        Next i
    Else
        ExpandMask octets, position + 1, prefix * 256# + Val(CStr(octets(position))), intervals, keys, budget
    End If
End Sub

Public Function AT_IPDeclarationsEqual(ByVal a As ATRuntimeValue, ByVal b As ATRuntimeValue, ByVal budget As ATEvalBudget) As Boolean
    Dim actualSet As New ATIntervalSet, expectedSet As New ATIntervalSet, key As Variant, parts As Variant, keysA As Collection, keysB As Collection
    If a.ValueState <> atValuePresent Or b.ValueState <> atValuePresent Then Err.Raise vbObjectError + 2613, , "DOMAIN_INVALID_DECLARATION"
    If a.DataType <> "IP_EXPRESSION" Or b.DataType <> "IP_EXPRESSION" Then Err.Raise vbObjectError + 2613, , "DOMAIN_INVALID_DECLARATION"
    actualSet.Initialize 4294967295#: expectedSet.Initialize 4294967295#
    Set keysA = a.DeclarationKeys: Set keysB = b.DeclarationKeys
    For Each key In keysA
        budget.Step
        If Left$(CStr(key), 5) = "CIDR:" Then
            If Not AT_HasDomainKey(keysB, CStr(key)) Then Exit Function
        Else
            parts = Split(CStr(key), ":")
            actualSet.Add Val(CStr(parts(1))), Val(CStr(parts(2))), budget
        End If
    Next key
    For Each key In keysB
        budget.Step
        If Left$(CStr(key), 5) = "CIDR:" Then
            If Not AT_HasDomainKey(keysA, CStr(key)) Then Exit Function
        Else
            parts = Split(CStr(key), ":")
            expectedSet.Add Val(CStr(parts(1))), Val(CStr(parts(2))), budget
        End If
    Next key
    actualSet.Freeze: expectedSet.Freeze
    AT_IPDeclarationsEqual = actualSet.IsSubsetOf(expectedSet, budget) And expectedSet.IsSubsetOf(actualSet, budget)
End Function
