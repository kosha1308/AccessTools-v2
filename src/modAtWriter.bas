Attribute VB_Name = "modAtWriter"
Option Explicit

Private busy As Boolean

Private Function NoteProfileSupported(ByVal workbook As Object) As Boolean
    ' PR-06 proves one legacy shape profile on this exact x64 host. New hosts
    ' must pass capability acceptance before notes are enabled, not after writing.
#If Win64 Then
    NoteProfileSupported = workbook.FileFormat = 56 And workbook.Application.Version = "16.0" And CStr(workbook.Application.Build) = "17932"
#End If
End Function

Private Function NoteTextValid(ByVal text As String) As Boolean
    Dim i As Long, c As Long
    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1)) And &HFFFF&
        If c < 32 And c <> 9 And c <> 10 And c <> 13 Then Exit Function
    Next i
    NoteTextValid = True
End Function

Public Function AtNewOwnershipJournal(ByVal workbook As Object, ByVal directory As String, Optional ByVal maximumRevisions As Long = 256) As CAtOwnershipJournal
    Dim journal As New CAtOwnershipJournal
    journal.Initialize workbook, directory, maximumRevisions
    Set AtNewOwnershipJournal = journal
End Function

Private Function Receipt(ByVal runId As String, ByVal status As String, ByVal code As String, ByVal message As String, Optional ByVal row As Object = Nothing, Optional ByVal actual As String = "") As CAtActionReceipt
    Dim r As New CAtActionReceipt
    r.RunId = runId: r.Status = status: r.Code = code: r.Message = message: r.Actual = actual
    If Not row Is Nothing Then r.ActionInstanceKey = row("Keys"): r.Target = row("Target")
    Set Receipt = r
End Function

Private Function AreaFor(ByVal workbook As Object, ByVal sheetName As String, ByVal address As String) As Object
    Dim sheet As Object
    For Each sheet In workbook.Worksheets
        If StrComp(sheet.Name, sheetName, vbTextCompare) = 0 Then
            Set AreaFor = sheet.Range(address): Exit Function
        End If
    Next sheet
    Err.Raise 5, , "Worksheet missing"
End Function

Private Function Priority(ByVal severity As String) As Long
    Select Case severity
        Case "INFO": Priority = 1
        Case "WARNING": Priority = 2
        Case "ERROR": Priority = 3
    End Select
End Function

Public Function AtPlanActions(ByVal workbook As Object, ByVal intents As Collection, ByVal runId As String, ByVal journal As CAtOwnershipJournal) As CAtActionPlan
    Dim plan As New CAtActionPlan, actions As New Collection, index As Object, keys As Object, noteSheets As Object, targets As Object
    Dim intent As CAtActionIntent, requested As Object, area As Object, row As Object, entry As Object, snapshot As Object
    Dim target As String, status As String, code As String, message As String, key As Variant, reason As String
    On Error GoTo Broken
    status = "FAILED": code = "PLAN_INVALID": message = "Invalid action, address or explicit workbook."
    If Not AtValidId(runId) Or intents.Count > 32 Then GoTo Done
    If Not journal.BoundTo(workbook) Then code = "BOOK_IDENTITY": GoTo Done
    If journal.RecoveryRequired Or Not journal.Verify() Then status = "RECOVERY_REQUIRED": code = "RECOVERY_REQUIRED": GoTo Done
    Set index = AtDictionary(): Set keys = AtDictionary()
    For Each intent In intents
        code = "PLAN_INVALID"
        If Not AtValidId(intent.ActionInstanceKey) Or Not AtValidAddress(intent.Address) Or Priority(intent.Severity) = 0 Then GoTo Done
        If keys.Exists(intent.ActionInstanceKey) Then GoTo Done
        keys.Add intent.ActionInstanceKey, True
        If Not intent.Highlight And Not intent.AddNote Then GoTo Done
        If intent.FillColor < 0 Or intent.FillColor > 16777215 Or Len(intent.Message) = 0 Then GoTo Done
        If Not NoteTextValid(intent.Message) Then code = "CAPABILITY_UNSUPPORTED": message = "Unsupported control character.": GoTo Done
        If intent.AddNote And Not NoteProfileSupported(workbook) Then code = "CAPABILITY_UNSUPPORTED": message = "Legacy note capability profile is not proven for this host/workbook format.": GoTo Done
        Set requested = AreaFor(workbook, intent.Sheet, intent.Address)
        Set area = requested.Cells(1, 1).MergeArea
        code = "CAPABILITY_UNSUPPORTED"
        If requested.Row < area.Row Or requested.Column < area.Column Or requested.Row + requested.Rows.Count > area.Row + area.Rows.Count Or requested.Column + requested.Columns.Count > area.Column + area.Columns.Count Then GoTo Done
        target = area.Parent.CodeName & "!" & area.Address
        If Not index.Exists(target) Then
            Set snapshot = AtWCapture(area, code, reason)
            Set entry = journal.Entry(target)
            If snapshot Is Nothing Then
                message = reason
                If Not entry Is Nothing Then status = "CONFLICT": code = "TARGET_CHANGED"
                GoTo Done
            End If
            If intent.AddNote And snapshot("ThreadedMode") = "UNAVAILABLE" And snapshot("FileFormat") <> "56" Then
                code = "CAPABILITY_UNSUPPORTED": message = "Threaded API unavailable: notes supported only in legacy BIFF8 compatibility workbooks.": GoTo Done
            End If
            If Not entry Is Nothing Then
                If entry("State") <> "ROLLED_BACK" Then
                    If Not AtWEqual(snapshot, AtWUnpack(entry("Expected"))) Then status = "CONFLICT": code = "TARGET_CHANGED": GoTo Done
                Else
                    Set entry = Nothing
                End If
            End If
            If entry Is Nothing Then
                If snapshot("Note") <> "0" Then code = "CAPABILITY_UNSUPPORTED": message = "Existing user notes are never edited or deleted.": GoTo Done
            End If
            Set row = AtDictionary()
            row.Add "Target", target: row.Add "Sheet", area.Parent.Name: row.Add "Address", area.Address
            row.Add "Observed", AtWPack(snapshot): row.Add "Before", AtWPack(snapshot)
            If Not entry Is Nothing Then row("Before") = entry("Before")
            row.Add "Keys", "": row.Add "RunId", runId: row.Add "Messages", "AccessTools"
            row.Add "Fill", "0": row.Add "Note", "0": row.Add "Priority", "0": row.Add "Color", "0": row.Add "Clear", "0"
            index.Add target, row: actions.Add row
        End If
        Set row = index(target)
        If row("Keys") <> "" Then row("Keys") = row("Keys") & "|"
        row("Keys") = row("Keys") & runId & "/" & intent.ActionInstanceKey & "/" & target
        row("Messages") = row("Messages") & vbLf & "[" & intent.Severity & "] " & Replace(Replace(intent.Message, vbCrLf, vbLf), vbCr, vbLf)
        If Len(row("Messages")) > 512 Then code = "CAPABILITY_UNSUPPORTED": GoTo Done
        If intent.Highlight Then
            row("Fill") = "1"
            If Priority(intent.Severity) > CLng(row("Priority")) Then row("Priority") = CStr(Priority(intent.Severity)): row("Color") = CStr(intent.FillColor)
        End If
        If intent.AddNote Then row("Note") = "1"
    Next intent
    ' Reconcile obsolete owned regions, using journal addresses, never new locators.
    For Each entry In journal.Entries
        If entry("State") <> "ROLLED_BACK" And Not index.Exists(entry("Target")) Then
            Set area = AreaFor(workbook, entry("Sheet"), entry("Address"))
            Set snapshot = AtWCapture(area, code)
            If Not AtWEqual(snapshot, AtWUnpack(entry("Expected"))) Then status = "CONFLICT": code = "TARGET_CHANGED": GoTo Done
            Set row = AtRecordCopy(entry)
            row("Observed") = AtWPack(snapshot): row("Clear") = "1": row("RunId") = runId
            actions.Add row
        End If
    Next entry
    ' Z-order is part of the guarded note shape. Do not delete/recreate one
    ' note while silently changing the Z-order fingerprint of another owned note.
    Set noteSheets = AtDictionary()
    Set targets = AtDictionary()
    For Each entry In journal.Entries: targets.Add entry("Target"), True: Next entry
    For Each row In actions
        If Not targets.Exists(row("Target")) Then targets.Add row("Target"), True
        Set snapshot = AtWUnpack(row("Observed"))
        If row("Clear") = "0" And row("Note") = "1" And Len(row("Observed")) > 32768 Then code = "CAPABILITY_UNSUPPORTED": message = "Insufficient bounded snapshot space for a new note.": GoTo Done
        If row("Note") = "1" Or snapshot("Note") = "1" Then
            If noteSheets.Exists(row("Sheet")) Then code = "CAPABILITY_UNSUPPORTED": message = "At most one owned note region per worksheet, including pending cleanup.": GoTo Done
            noteSheets.Add row("Sheet"), True
        End If
    Next row
    If targets.Count > 32 Then code = "LIMIT": message = "Journal target budget exceeded before mutation.": GoTo Done
    status = "READY": code = "": message = ""
Done:
    plan.Initialize workbook, runId, journal.SessionId, actions, status, code, message
    Set AtPlanActions = plan: Exit Function
Broken:
    status = "FAILED": code = "PLAN_INVALID": message = "Planning failed; no workbook mutation performed: " & Err.Description
    Resume Done
End Function

Private Function KnownState(ByVal snapshot As Object, ByVal entry As Object) As Boolean
    If AtWEqual(snapshot, AtWUnpack(entry("Expected"))) Then KnownState = True: Exit Function
    If entry.Exists("Next") Then
        If entry("Next") <> "" Then KnownState = AtWEqual(snapshot, AtWUnpack(entry("Next")))
    End If
End Function

Private Sub Pending(ByVal journal As CAtOwnershipJournal, ByVal entry As Object, ByVal current As Object, ByVal nextState As Object, ByVal phase As String)
    entry("State") = "PENDING": entry("Phase") = phase: entry("StepBefore") = AtWPack(current)
    entry("Next") = "": entry("Actual") = ""
    If Not nextState Is Nothing Then entry("Next") = AtWPack(nextState)
    journal.Persist entry
End Sub

Private Sub Acknowledge(ByVal journal As CAtOwnershipJournal, ByVal entry As Object, ByVal actual As Object)
    entry("Expected") = AtWPack(actual): entry("Actual") = AtWPack(actual): entry("State") = "STEP_APPLIED"
    entry("Next") = "": journal.Persist entry
End Sub

Private Sub CheckBudget(ByVal count As Long, ByVal maximum As Long)
    If maximum >= 0 And count >= maximum Then Err.Raise 5, , "Mutation budget exhausted; explicit recovery required"
End Sub

Private Function ApplyRegion(ByVal workbook As Object, ByVal row As Object, ByVal journal As CAtOwnershipJournal, ByRef mutations As Long, ByVal maximum As Long) As CAtActionReceipt
    Dim area As Object, original As Object, current As Object, desired As Object, nextState As Object, actual As Object, entry As Object
    Dim code As String, note As Object, changed As Boolean, reason As String
    Set area = AreaFor(workbook, row("Sheet"), row("Address"))
    Set current = AtWCapture(area, code): Set original = AtWUnpack(row("Before"))
    If Not AtWEqual(current, AtWUnpack(row("Observed"))) Then Err.Raise 5, , "Target changed after preflight"
    Set entry = AtRecordCopy(row)
    entry("State") = "INTENT": entry("Phase") = "INTENT": entry("Expected") = AtWPack(current)
    entry("StepBefore") = AtWPack(current): entry("Next") = "": entry("Actual") = ""
    Set desired = AtRecordCopy(original)
    If row("Clear") = "0" Then
        If row("Fill") = "1" Then desired("Pattern") = "1": desired("Color") = row("Color")
    End If
    Set nextState = AtWFillState(current, desired)
    If Not AtWEqual(current, nextState) Then
        CheckBudget mutations, maximum
        Pending journal, entry, current, nextState, "FILL"
        AtWSetFill area, desired: mutations = mutations + 1: changed = True
        Set actual = AtWCapture(area, code)
        If Not AtWEqual(actual, nextState) Then Err.Raise 5, , "Fill result not verified"
        Acknowledge journal, entry, actual: Set current = actual
    End If
    Set note = area.Cells(1, 1).Comment
    If Not note Is Nothing Then
        If row("Clear") = "1" Or row("Note") = "0" Or note.Text <> row("Messages") Then
            CheckBudget mutations, maximum
            Set nextState = AtWWithoutNote(current)
            Pending journal, entry, current, nextState, "REMOVE_NOTE"
            note.Delete: mutations = mutations + 1: changed = True
            Set actual = AtWCapture(area, code)
            If Not AtWEqual(actual, nextState) Then Err.Raise 5, , "Note removal not verified"
            Acknowledge journal, entry, actual: Set current = actual
        End If
    End If
    If row("Clear") = "0" And row("Note") = "1" And current("Note") = "0" Then
        CheckBudget mutations, maximum
        ' New note shape defaults are not predictable before creation: Next remains
        ' unknown. A crash before acknowledgement cannot confer ownership by text.
        Pending journal, entry, current, Nothing, "ADD_NOTE"
        area.Cells(1, 1).AddComment row("Messages"): mutations = mutations + 1: changed = True
        Set actual = AtWCapture(area, code, reason)
        If actual Is Nothing Then Err.Raise 5, , "New note capability not verified: " & reason
        If actual("NText") <> row("Messages") Then Err.Raise 5, , "Note text not verified"
        If Not AtWEqual(AtWWithoutNote(actual), AtWWithoutNote(current)) Then Err.Raise 5, , "Unexpected side effect while adding note"
        Acknowledge journal, entry, actual: Set current = actual
    End If
    entry("Expected") = AtWPack(current): entry("Actual") = AtWPack(current)
    entry("State") = IIf(row("Clear") = "1", "ROLLED_BACK", "APPLIED"): entry("Phase") = "COMPLETE"
    journal.Persist entry
    Set ApplyRegion = Receipt(row("RunId"), IIf(row("Clear") = "1", "ROLLED_BACK", IIf(changed, "APPLIED", "SKIPPED")), "", "Verified against actual Excel state.", row, AtWPack(current))
End Function

' maximumMutations is an explicit host interruption budget; -1 means no early
' stop. It allows deterministic interruption/recovery tests without user books.
Public Function AtApplyPlan(ByVal plan As CAtActionPlan, ByVal journal As CAtOwnershipJournal, Optional ByVal maximumMutations As Long = -1) As Collection
    Dim results As New Collection, row As Variant, area As Object, current As Object, code As String, mutations As Long
    Dim completed As Long, rest As Long
    On Error GoTo Broken
    If busy Then results.Add Receipt(plan.RunId, "FAILED", "WRITER_BUSY", "Concurrent writer rejected."): GoTo Done
    If plan.Status <> "READY" Then results.Add Receipt(plan.RunId, plan.Status, plan.Code, plan.Message): GoTo Done
    If plan.SessionId <> journal.SessionId Or Not journal.BoundTo(plan.Workbook) Then results.Add Receipt(plan.RunId, "CONFLICT", "BOOK_IDENTITY", "Plan/journal workbook mismatch."): GoTo Done
    If journal.RecoveryRequired Or Not journal.Verify() Then results.Add Receipt(plan.RunId, "RECOVERY_REQUIRED", "RECOVERY_REQUIRED", "Journal reconciliation required."): GoTo Done
    If maximumMutations < -1 Then results.Add Receipt(plan.RunId, "FAILED", "BUDGET_INVALID", "Invalid mutation budget."): GoTo Done
    ' Revalidate the entire plan before the first write.
    For Each row In plan.Actions
        Set area = AreaFor(plan.Workbook, row("Sheet"), row("Address"))
        Set current = AtWCapture(area, code)
        If Not AtWEqual(current, AtWUnpack(row("Observed"))) Then results.Add Receipt(plan.RunId, "CONFLICT", "TARGET_CHANGED", "Target changed since planning.", row): GoTo Done
    Next row
    busy = True
    For Each row In plan.Actions
        results.Add ApplyRegion(plan.Workbook, row, journal, mutations, maximumMutations)
        completed = completed + 1
    Next row
    busy = False
Done:
    Set AtApplyPlan = results: Exit Function
Broken:
    busy = False: journal.RequireRecovery
    If IsObject(row) Then
        results.Add Receipt(plan.RunId, "RECOVERY_REQUIRED", "WRITE_INTERRUPTED", "Write/journal verification failed: " & Err.Description, row)
    Else
        results.Add Receipt(plan.RunId, "RECOVERY_REQUIRED", "WRITE_INTERRUPTED", "Write/journal verification failed: " & Err.Description)
    End If
    For rest = completed + 2 To plan.Actions.Count
        results.Add Receipt(plan.RunId, "SKIPPED", "BLOCKED_RECOVERY", "Not attempted after interrupted write.", plan.Actions(rest))
    Next rest
    Resume Done
End Function

Public Function AtClearOwned(ByVal workbook As Object, ByVal journal As CAtOwnershipJournal, ByVal runId As String) As Collection
    Dim none As New Collection, plan As CAtActionPlan
    Set plan = AtPlanActions(workbook, none, runId, journal)
    Set AtClearOwned = AtApplyPlan(plan, journal)
End Function

Public Function AtRecoverOwned(ByVal workbook As Object, ByVal journal As CAtOwnershipJournal, ByVal runId As String) As Collection
    Dim results As New Collection, row As Variant, entry As Variant, current As Object, area As Object, actions As New Collection
    Dim code As String, count As Long
    On Error GoTo Broken
    If busy Then results.Add Receipt(runId, "FAILED", "WRITER_BUSY", "Concurrent recovery rejected."): GoTo Done
    If Not AtValidId(runId) Or Not journal.BoundTo(workbook) Then GoTo Broken
    If Not journal.Verify() Or journal.EntryCount = 0 Then GoTo Broken
    For Each entry In journal.Entries
        If entry("State") <> "ROLLED_BACK" Then
            Set area = AreaFor(workbook, entry("Sheet"), entry("Address"))
            Set current = AtWCapture(area, code)
            If Not KnownState(current, entry) Then
                journal.RequireRecovery
                results.Add Receipt(runId, "CONFLICT", "RECOVERY_REQUIRED", "Current state is not a fully journaled before/after state; manual inspection required.", entry): GoTo Done
            End If
            Set row = AtRecordCopy(entry): row("Observed") = AtWPack(current): row("Clear") = "1": row("RunId") = runId
            actions.Add row
        End If
    Next entry
    busy = True: journal.Recovering = True
    For Each row In actions: results.Add ApplyRegion(workbook, row, journal, count, -1): Next row
    busy = False: journal.Recovering = False: journal.Recovered
Done:
    Set AtRecoverOwned = results: Exit Function
Broken:
    busy = False: journal.Recovering = False: journal.RequireRecovery
    results.Add Receipt(runId, "RECOVERY_REQUIRED", "RECOVERY_REQUIRED", "Recovery unavailable for this journal/session or state.")
    Resume Done
End Function
