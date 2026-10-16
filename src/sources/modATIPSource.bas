Attribute VB_Name = "modATIPSource"
Option Explicit
Public Function AT_IPSourceKey(ByVal sourceId As String, ByVal fields As Collection) As String
    Dim field As Variant, seen As Object, canonical As String
    If Len(sourceId) = 0 Or Len(sourceId) > 128 Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_SCHEMA_INVALID"
    sourceId = AT_CanonicalId(sourceId)
    If fields Is Nothing Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_SCHEMA_INVALID"
    If fields.Count = 0 Or fields.Count > 16 Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_SCHEMA_INVALID"
    Set seen = CreateObject("Scripting.Dictionary"): AT_IPSourceKey = sourceId
    For Each field In fields
        If VarType(field) <> vbString Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_SCHEMA_INVALID"
        If Len(CStr(field)) = 0 Or Len(CStr(field)) > 128 Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_SCHEMA_INVALID"
        canonical = AT_CanonicalId(CStr(field))
        If seen.Exists(canonical) Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_SCHEMA_INVALID"
        seen.Add canonical, True
        AT_IPSourceKey = AT_IPSourceKey & "|" & Len(canonical) & ":" & canonical
    Next field
End Function
Public Function AT_IPSourceFailure(ByVal description As String) As ATEvalResult
    Dim code As String
    code = AT_StableErrorCode(description, "IP_SOURCE_INPUT_INVALID")
    If Left$(code, 10) <> "IP_SOURCE_" And Left$(code, 14) <> "RUNTIME_LIMIT_" Then code = "IP_SOURCE_INPUT_INVALID"
    Set AT_IPSourceFailure = AT_RuntimeFailure(atEvalError, code, Nothing)
End Function
Public Sub AT_IPSortPositions(ByRef values() As Long, ByVal count As Long, ByVal budget As ATEvalBudget)
    Dim buffer() As Long, width As Long, first As Long, middle As Long, last As Long, left As Long, right As Long, target As Long
    Dim chooseLeft As Boolean
    If count < 2 Then Exit Sub
    ReDim buffer(1 To count): width = 1
    Do While width < count
        For first = 1 To count Step width * 2
            middle = first + width - 1: If middle > count Then middle = count
            last = first + width * 2 - 1: If last > count Then last = count
            left = first: right = middle + 1
            For target = first To last
                budget.Step
                If left > middle Then
                    chooseLeft = False
                ElseIf right > last Then
                    chooseLeft = True
                Else
                    chooseLeft = values(left) <= values(right)
                End If
                If chooseLeft Then buffer(target) = values(left): left = left + 1 Else buffer(target) = values(right): right = right + 1
            Next target
        Next first
        For target = 1 To count: budget.Step: values(target) = buffer(target): Next target
        width = width * 2
    Loop
End Sub
