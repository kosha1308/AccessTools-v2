VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATInitiatorSelector"
Option Explicit
Private mReady As Boolean, mPositive As Object, mNegative As Object
Friend Sub Initialize(ByVal text As String, ByVal budget As ATEvalBudget)
    Dim raw As Variant, token As String, key As String, negative As Boolean, target As Object
    If mReady Then Err.Raise vbObjectError + 2970, , "INITIATOR_FROZEN"
    Set mPositive = CreateObject("Scripting.Dictionary"): Set mNegative = CreateObject("Scripting.Dictionary")
    For Each raw In AT_InitiatorTokens(text, budget)
        token = AT_AsciiCase(CStr(raw), False): negative = AT_InitiatorNegative(token)
        If Left$(token, 1) = "@" Then
            token = Mid$(token, 2)
            If InStr(token, "@") > 0 Then GoTo Invalid
        End If
        If InStr(token, "@") > 0 Then
            If Not AT_EmailSyntaxValid(token) Then GoTo Invalid
            If Not ValidDomain(Mid$(token, InStr(token, "@") + 1)) Then GoTo Invalid
            key = "E:" & token
        Else
            If Not AT_EmailSyntaxValid("probe@" & token) Then GoTo Invalid
            If Not ValidDomain(token) Then GoTo Invalid
            key = "D:" & token
        End If
        Set target = mPositive: If negative Then Set target = mNegative
        If Not target.Exists(key) Then target.Add key, True
    Next raw
    If mPositive.Count = 0 Then GoTo Invalid
    mReady = True
    Exit Sub
Invalid:
    Err.Raise vbObjectError + 2970, , "INITIATOR_SELECTOR_INVALID"
End Sub
Public Function Rank(ByVal email As String, ByVal budget As ATEvalBudget) As Long
    Dim domain As String
    If Not mReady Then Err.Raise vbObjectError + 2970, , "INITIATOR_ACCESS"
    budget.CheckText Len(email): budget.Step Len(email) + 1
    email = AT_AsciiCase(Trim$(email), False)
    If Not AT_EmailSyntaxValid(email) Then Exit Function
    domain = Mid$(email, InStr(email, "@") + 1)
    If mNegative.Exists("E:" & email) Or mNegative.Exists("D:" & domain) Then Exit Function
    If mPositive.Exists("E:" & email) Then Rank = 2: Exit Function
    If mPositive.Exists("D:" & domain) Then Rank = 1
End Function

Private Function ValidDomain(ByVal domain As String) As Boolean
    Dim label As Variant, i As Long, ch As String
    If Len(domain) > 253 Then Exit Function
    For Each label In Split(domain, ".")
        If Len(CStr(label)) = 0 Or Len(CStr(label)) > 63 Then Exit Function
        If Left$(CStr(label), 1) = "-" Or Right$(CStr(label), 1) = "-" Then Exit Function
        For i = 1 To Len(CStr(label))
            ch = Mid$(CStr(label), i, 1)
            If Not ((ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "-") Then Exit Function
        Next i
    Next label
    ValidDomain = True
End Function
