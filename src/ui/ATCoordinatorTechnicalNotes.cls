VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATCoordinatorTechnicalNotes"
Option Explicit
Public Function Run(ByVal source As ATNetworkReadResult, ByVal registry As ATNetworkRegistryResult, ByVal policy As ATNetworkPolicyReport, Optional ByVal externalBudget As ATEvalBudget = Nothing) As ATCoordinatorNotesResult
    Dim result As New ATCoordinatorNotesResult, rows As New Collection, row As ATRuntimeRow, raw As ATRuntimeValue
    Dim budget As New ATEvalBudget, limits As New ATEvalLimits, refRows As Object, contexts As Object, texts As Object, statuses As Object
    Dim context As ATTechnicalNoteContext, entry As ATNetworkRegistryEntry, entries As Collection, issue As ATCheckIssue, outcome As ATCheckOutcome
    Dim i As Long, found As Long, incomplete As Boolean, globalCount As Long, assigned As Boolean, key As String, block As String, code As String
    On Error GoTo Failed
    If source Is Nothing Or registry Is Nothing Or policy Is Nothing Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_INPUT"
    registry.ValidateRequest source
    If source.State <> atEvalOK Or registry.State <> atEvalOK Or policy.AllReport.State <> atEvalOK Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_BLOCKED"
    For Each outcome In policy.AllReport.Checks
        If outcome.CheckId = "NET.IP_EXCEPTIONS" Then
            found = found + 1: incomplete = outcome.Incomplete Or outcome.Status = "SKIPPED"
            If outcome.Status = "BLOCKED" Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_BLOCKED"
        End If
    Next outcome
    If found <> 1 Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_INPUT"
    If externalBudget Is Nothing Then budget.Initialize limits Else Set budget = externalBudget
    Set refRows = CreateObject("Scripting.Dictionary"): Set contexts = CreateObject("Scripting.Dictionary")
    Set texts = CreateObject("Scripting.Dictionary"): Set statuses = CreateObject("Scripting.Dictionary")
    Set raw = AT_InitiatorRows(source.Context.ResolveField("TECHNICAL_ROWS_RAW"), "TECHNICAL_TABLE", budget)
    For i = 1 To raw.Count
        Set row = raw.Row(i): budget.Step
        If Len(row.CellRef("SOURCE_RAW")) = 0 Or Len(row.CellRef("DESTINATION_RAW")) = 0 Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_INPUT"
        refRows.Add row.CellRef("SOURCE_RAW"), row: refRows.Add row.CellRef("DESTINATION_RAW"), row
    Next i
    Set entries = registry.Entries: budget.CheckList entries.Count
    For Each entry In entries
        budget.Step: texts.Add entry.Id, "": statuses.Add entry.Id, "PASS"
    Next entry
    For Each issue In policy.SupersetIssues
        budget.Step
        If issue.CheckId = "NET.IP_EXCEPTIONS" Then
            assigned = False
            If refRows.Exists(issue.CellRef) Then
                Set row = refRows(issue.CellRef): key = row.RowId
                If Not contexts.Exists(key) Then
                    Set context = New ATTechnicalNoteContext: context.Initialize row, budget: contexts.Add key, context
                Else
                    Set context = contexts(key)
                End If
                block = context.Description & vbLf & issue.Message
                For Each entry In entries
                    If context.Includes(entry, budget) Then
                        budget.Step Len(block)
                        key = texts(entry.Id): Append key, block: texts(entry.Id) = key
                        If issue.Severity = "ERROR" Then statuses(entry.Id) = "FAIL"
                        If issue.Severity = "WARNING" And statuses(entry.Id) = "PASS" Then statuses(entry.Id) = "WARNING"
                        assigned = True
                    End If
                Next entry
            End If
            If Not assigned Then globalCount = globalCount + 1
        End If
    Next issue
    If globalCount > 0 Then incomplete = True
    For Each entry In entries
        budget.Step: Set row = New ATRuntimeRow: row.Initialize "COORDINATOR_NOTES", entry.Id
        row.AddField "INSTRUCTION", AT_RuntimeOK(AT_RuntimeScalar("TEXT", CStr(texts(entry.Id))))
        row.AddField "STATUS", AT_RuntimeOK(AT_RuntimeScalar("TEXT", CStr(statuses(entry.Id))))
        rows.Add row
    Next entry
    result.Initialize atEvalOK, registry, rows, globalCount, incomplete: Set Run = result: Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "COORDINATOR_NOTES_INPUT")
    Set rows = New Collection: Set result = New ATCoordinatorNotesResult
    result.Initialize atEvalError, registry, rows, 0, True, code: Set Run = result
End Function
Private Sub Append(ByRef target As String, ByVal block As String)
    If Len(target) + Len(block) + 2 > 32767 Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_LIMIT"
    If Len(target) > 0 Then target = target & vbLf & vbLf
    target = target & block
End Sub
