VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATTechnicalNoteContext"
Option Explicit
Private mReady As Boolean, mSourceRef As String, mDestinationRef As String, mDescription As String
Private mSource As ATIntervalSet, mDestination As ATIntervalSet
Friend Sub Initialize(ByVal row As ATRuntimeRow, ByVal budget As ATEvalBudget)
    Dim fields As Variant, texts(0 To 3) As String, i As Long, valid As Boolean, parsed As ATEvalResult
    If mReady Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_FROZEN"
    fields = Array("SOURCE_RAW", "DESTINATION_RAW", "PORT_RAW", "PROTOCOL_RAW")
    mSourceRef = row.CellRef("SOURCE_RAW"): mDestinationRef = row.CellRef("DESTINATION_RAW")
    For i = 0 To 3
        texts(i) = AT_RegistryText(row.ResolveField(CStr(fields(i)), budget), valid, budget)
        If i < 2 And valid And Not AT_ExcelBlank(texts(i)) Then
            Set parsed = AT_ParseDomain("IP_EXPRESSION", Replace(Replace(texts(i), "х", "x"), "Х", "x"), Nothing, budget)
            If parsed.State <> atEvalOK Then Err.Raise vbObjectError + 3160, , parsed.Diagnostics(1).Code
            If parsed.Value.ValueState = atValuePresent Then
                If i = 0 Then Set mSource = parsed.Value.Domain Else Set mDestination = parsed.Value.Domain
            End If
        End If
        If Not valid Then texts(i) = "[ошибка Excel]"
    Next i
    mDescription = "Техническая связанность: " & texts(0) & " -> " & texts(1) & "; порт: " & texts(2) & "; протокол: " & texts(3) & ". Ячейки: " & mSourceRef & ", " & mDestinationRef
    If Len(mDescription) > 32767 Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_LIMIT"
    mReady = True
End Sub
Friend Function Includes(ByVal entry As ATNetworkRegistryEntry, ByVal budget As ATEvalBudget) As Boolean
    Dim refs As ATRuntimeValue, i As Long, ref As String
    If Not mReady Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_INPUT"
    budget.Step: Set refs = entry.References
    For i = 1 To refs.Count
        ref = AT_PolicyCell(refs.Row(i), "CELL_REF", budget)
        If ref = mSourceRef Or ref = mDestinationRef Then Includes = True: Exit Function
    Next i
    If Not mSource Is Nothing Then
        If entry.Domain.Overlaps(mSource, budget) Then Includes = True: Exit Function
    End If
    If Not mDestination Is Nothing Then Includes = entry.Domain.Overlaps(mDestination, budget)
End Function
Friend Property Get Description() As String
    If Not mReady Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_INPUT"
    Description = mDescription
End Property
