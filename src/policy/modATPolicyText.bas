Attribute VB_Name = "modATPolicyText"
Option Explicit

Public Function AT_PolicyBoolean(ByVal text As String) As Boolean
    Select Case AT_ExcelHeader(text)
        Case "да", "yes", "true", "1": AT_PolicyBoolean = True
        Case "нет", "no", "false", "0"
        Case Else: Err.Raise vbObjectError + 2960, , "POLICY_FLAG_INVALID"
    End Select
End Function

Public Function AT_PolicySeverity(ByVal text As String) As String
    Select Case AT_ExcelHeader(text)
        Case "ошибка", "error": AT_PolicySeverity = "ERROR"
        Case "предупреждение", "warning", "warn": AT_PolicySeverity = "WARNING"
        Case "информация", "info": AT_PolicySeverity = "INFO"
        Case Else: Err.Raise vbObjectError + 2960, , "POLICY_SEVERITY_INVALID"
    End Select
End Function

Public Function AT_PolicyKey(ByVal text As String) As String
    Dim parts As Variant, ids As Variant, id As Variant, check As String, severity As String, field As String
    On Error GoTo Invalid
    If Len(text) > 512 Then GoTo Invalid
    parts = Split(text, "|")
    If UBound(parts) <> 2 Then GoTo Invalid
    ids = Split(Trim$(parts(0)), ".")
    If UBound(ids) <> 1 Then GoTo Invalid
    For Each id In ids
        If Len(check) > 0 Then check = check & "."
        check = check & AT_CanonicalId(CStr(id))
    Next id
    severity = UCase$(Trim$(parts(1)))
    If severity <> "ERROR" And severity <> "WARNING" And severity <> "INFO" Then GoTo Invalid
    field = AT_ExcelHeader(CStr(parts(2)))
    If Len(field) = 0 Or Len(field) > 256 Then GoTo Invalid
    AT_PolicyKey = check & "|" & severity & "|" & field
    Exit Function
Invalid:
    Err.Raise vbObjectError + 2960, , "POLICY_KEY_INVALID"
End Function

Public Function AT_PolicyCell(ByVal row As ATRuntimeRow, ByVal id As String, ByVal budget As ATEvalBudget) As String
    Dim result As ATEvalResult, value As ATRuntimeValue
    budget.Step
    Set result = row.ResolveField(id, budget)
    If result.State <> atEvalOK Then GoTo Invalid
    Set value = result.Value
    If value.DataType <> "TEXT" Then GoTo Invalid
    Select Case value.ValueState
        Case atValueEmpty
        Case atValuePresent: AT_PolicyCell = CStr(value.Scalar)
        Case Else: GoTo Invalid
    End Select
    budget.Step Len(AT_PolicyCell)
    Exit Function
Invalid:
    Err.Raise vbObjectError + 2960, , "POLICY_INPUT_INVALID"
End Function

Public Function AT_PolicyExpand(ByVal template As String, ByVal issue As ATCheckIssue, ByVal field As String) As String
    Dim position As Long, closeAt As Long, token As String, output As String, piece As String, startAt As Long
    If Len(template) > 32767 Then Err.Raise vbObjectError + 2960, , "POLICY_MESSAGE_LIMIT"
    position = 1
    Do While position <= Len(template)
        startAt = position
        Do While position <= Len(template)
            If Mid$(template, position, 1) = "{" Or Mid$(template, position, 1) = "}" Then Exit Do
            position = position + 1
        Loop
        output = output & Mid$(template, startAt, position - startAt)
        If position > Len(template) Then Exit Do
        If Mid$(template, position, 1) <> "{" Then GoTo Invalid
        closeAt = InStr(position + 1, template, "}")
        If closeAt = 0 Then GoTo Invalid
        token = LCase$(Mid$(template, position + 1, closeAt - position - 1)): piece = ""
        Select Case token
            Case "field": piece = field
            Case "message", "check", "severity", "row"
                If Not issue Is Nothing Then
                    Select Case token
                        Case "message": piece = issue.Message
                        Case "check": piece = issue.CheckId
                        Case "severity": piece = issue.Severity
                        Case "row": piece = AT_PolicyRowNumber(issue.CellRef)
                    End Select
                End If
            Case Else: GoTo Invalid
        End Select
        If Len(output) + Len(piece) > 32767 Then Err.Raise vbObjectError + 2960, , "POLICY_MESSAGE_LIMIT"
        output = output & piece: position = closeAt + 1
    Loop
    If Len(output) > 32767 Then Err.Raise vbObjectError + 2960, , "POLICY_MESSAGE_LIMIT"
    AT_PolicyExpand = output
    Exit Function
Invalid:
    Err.Raise vbObjectError + 2960, , "POLICY_TEMPLATE_INVALID"
End Function

Private Function AT_PolicyRowNumber(ByVal ref As String) As String
    Dim i As Long, ch As String
    For i = Len(ref) To InStrRev(ref, "!") + 1 Step -1
        ch = Mid$(ref, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
        AT_PolicyRowNumber = ch & AT_PolicyRowNumber
    Next i
End Function
