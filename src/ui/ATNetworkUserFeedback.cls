VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkUserFeedback"
Option Explicit
Private Const FIRST_MARKER As String = "[AccessToolsv2]"
Private Const LAST_MARKER As String = "[/AccessToolsv2]"
Private Const FILL_MARKER As String = "AccessToolsv2:annotation"
Private mPolicy As ATNetworkPolicyReport, mText As String, mErrors As Long, mFailures As Long
Public Sub Prepare(ByVal policy As ATNetworkPolicyReport)
    Dim item As ATNetworkPolicyDecision, text As String, errors As Long, entry As String
    If policy Is Nothing Then Err.Raise vbObjectError + 3210, , "FEEDBACK_POLICY_REQUIRED"
    Set mPolicy = Nothing: mText = "": mErrors = 0: mFailures = 0
    If policy.AllReport.State <> atEvalOK Then Err.Raise vbObjectError + 3210, , "FEEDBACK_POLICY_BLOCKED"
    For Each item In policy.Decisions
        If item.InUserReport And item.UserError And item.Issue.Severity = "ERROR" Then
            errors = errors + 1
            entry = CStr(errors) & ". " & item.Issue.Message
            If Len(item.Issue.CellRef) > 0 Then entry = entry & vbCrLf & "Ячейка: " & item.Issue.CellRef
            If Len(text) + Len(entry) + 4 > 2000000 Then Err.Raise vbObjectError + 3210, , "FEEDBACK_TEXT_LIMIT"
            If Len(text) > 0 Then text = text & vbCrLf & vbCrLf
            text = text & entry
        End If
    Next item
    Set mPolicy = policy: mText = text: mErrors = errors
End Sub
Public Property Get UserErrorsText() As String
    UserErrorsText = mText
End Property
Public Property Get ErrorCount() As Long
    ErrorCount = mErrors
End Property
Public Property Get NoteFailures() As Long
    NoteFailures = mFailures
End Property
Public Property Get UserIssues() As Collection
    Dim items As New Collection, item As ATNetworkPolicyDecision
    If Not mPolicy Is Nothing Then
        For Each item In mPolicy.Decisions
            If item.InUserReport Then items.Add item.Issue
        Next item
    End If
    Set UserIssues = items
End Property
Public Function TargetCell(ByVal request As Worksheet, ByVal reference As String) As Range
    Set TargetCell = TrustedCell(request, reference)
End Function
Public Function ClearNotes(ByVal request As Worksheet) As Boolean
    Dim i As Long, note As Comment, original As String, clean As String, previousEvents As Boolean, armed As Boolean
    On Error GoTo Failed
    mFailures = 0
    If request Is Nothing Then Exit Function
    If request.ProtectContents Or request.Comments.Count > 10000 Then GoTo Failed
    previousEvents = Application.EnableEvents: armed = True: Application.EnableEvents = False
    Dim rules As FormatConditions
    Set rules = request.UsedRange.FormatConditions
    If rules.Count > 10000 Then GoTo Failed
    For i = rules.Count To 1 Step -1
        If rules(i).Type = xlExpression Then
            If rules(i).Formula1 = "=N(""" & FILL_MARKER & """)=0" Then rules(i).Delete
        End If
    Next i
    For i = request.Comments.Count To 1 Step -1
        Set note = request.Comments(i): original = note.Text: clean = WithoutOwned(original)
        If clean <> original Then
            If Len(clean) = 0 Then note.Delete Else note.Text Text:=clean
        End If
    Next i
    Application.EnableEvents = previousEvents: ClearNotes = True: Exit Function
Failed:
    mFailures = mFailures + 1
    On Error Resume Next
    If armed Then Application.EnableEvents = previousEvents
End Function
Public Function ApplyNotes(ByVal request As Worksheet, ByVal runId As String) As Long
    Dim notes As Object, item As ATNetworkPolicyDecision, target As Range, address As Variant, text As String
    Dim missing As Long, previousEvents As Boolean, armed As Boolean, severities As Object, rank As Long
    On Error GoTo Failed
    mFailures = 0
    If mPolicy Is Nothing Or request Is Nothing Then Err.Raise vbObjectError + 3210, , "FEEDBACK_NOT_READY"
    If request.ProtectContents Then mFailures = 1: Exit Function
    If request.Comments.Count > 10000 Then mFailures = 1: Exit Function
    Set notes = CreateObject("Scripting.Dictionary")
    Set severities = CreateObject("Scripting.Dictionary")
    For Each item In mPolicy.Decisions
        If item.InRequest Or (Not item.RequestConfigured And item.InUserReport And item.UserError And item.Issue.Severity = "ERROR") Then
            Set target = TrustedCell(request, item.Issue.CellRef)
            If target Is Nothing Then
                mFailures = mFailures + 1
            Else
                address = target.Address(False, False)
                If Not notes.Exists(address) Then notes.Add address, ""
                text = notes(address)
                If Len(text) > 0 Then text = text & vbLf & vbLf
                notes(address) = text & item.Issue.Message
                rank = SeverityRank(item.Issue.Severity)
                If Not severities.Exists(address) Then severities.Add address, rank
                If rank > severities(address) Then severities(address) = rank
            End If
        End If
    Next item
    missing = mFailures
    If Not ClearNotes(request) Then mFailures = mFailures + missing: Exit Function
    mFailures = mFailures + missing
    previousEvents = Application.EnableEvents: armed = True: Application.EnableEvents = False
    For Each address In notes.Keys
        Set target = request.Range(CStr(address))
        If WriteNote(target, CStr(notes(address)), runId) Then ApplyNotes = ApplyNotes + 1 Else mFailures = mFailures + 1
        If Not Highlight(target, CLng(severities(address))) Then mFailures = mFailures + 1
    Next address
    Application.EnableEvents = previousEvents: Exit Function
Failed:
    mFailures = mFailures + 1
    On Error Resume Next
    If armed Then Application.EnableEvents = previousEvents
End Function
Private Function SeverityRank(ByVal severity As String) As Long
    Select Case severity
        Case "ERROR": SeverityRank = 3
        Case "WARNING": SeverityRank = 2
        Case Else: SeverityRank = 1
    End Select
End Function
Private Function Highlight(ByVal target As Range, ByVal rank As Long) As Boolean
    Dim rule As FormatCondition, color As Long
    On Error GoTo Failed
    color = RGB(221, 235, 247)
    If rank = 2 Then color = RGB(255, 235, 156)
    If rank = 3 Then color = RGB(255, 199, 206)
    Set rule = target.MergeArea.FormatConditions.Add(Type:=xlExpression, Formula1:="=N(""" & FILL_MARKER & """)=0")
    rule.Interior.Color = color: rule.SetFirstPriority: rule.StopIfTrue = True
    Highlight = True
Failed:
End Function
Private Function WriteNote(ByVal target As Range, ByVal message As String, ByVal runId As String) As Boolean
    Dim original As String, text As String
    On Error GoTo Failed
    If Not target.Comment Is Nothing Then original = target.Comment.Text
    ' An incomplete ownership block is treated as user text and left untouched.
    If InStr(original, FIRST_MARKER) > 0 Then Exit Function
    text = FIRST_MARKER & vbLf & "Проверка " & runId & vbLf & message & vbLf & LAST_MARKER
    If Len(original) > 0 Then text = original & vbLf & vbLf & text
    If Len(text) > 32767 Then Exit Function
    If target.Comment Is Nothing Then target.AddComment text Else target.Comment.Text Text:=text
    target.Comment.Visible = False
    WriteNote = True
Failed:
End Function
Private Function WithoutOwned(ByVal text As String) As String
    Dim first As Long, last As Long, start As Long
    Do
        first = InStr(text, FIRST_MARKER)
        If first = 0 Then Exit Do
        last = InStr(first + Len(FIRST_MARKER), text, LAST_MARKER)
        If last = 0 Then Exit Do
        start = first
        If first > 2 Then
            If Mid$(text, first - 2, 2) = vbLf & vbLf Then start = first - 2
        End If
        text = Left$(text, start - 1) & Mid$(text, last + Len(LAST_MARKER))
    Loop
    WithoutOwned = text
End Function
Private Function TrustedCell(ByVal request As Worksheet, ByVal reference As String) As Range
    Dim prefix As String, address As String, i As Long, ch As String, digits As Boolean, letters As Long, target As Range
    On Error GoTo Failed
    prefix = "'" & Replace(request.Name, "'", "''") & "'!"
    If Left$(reference, Len(prefix)) <> prefix Then Exit Function
    address = Mid$(reference, Len(prefix) + 1)
    If Len(address) < 2 Or Len(address) > 10 Then Exit Function
    For i = 1 To Len(address)
        ch = Mid$(address, i, 1)
        If ch >= "A" And ch <= "Z" And Not digits Then
            letters = letters + 1: If letters > 3 Then Exit Function
        ElseIf ch >= "0" And ch <= "9" And letters > 0 Then
            digits = True
        Else
            Exit Function
        End If
    Next i
    If Not digits Then Exit Function
    Set target = request.Range(address)
    If target.Address(False, False) <> address Then Exit Function
    Set TrustedCell = target.MergeArea.Cells(1, 1)
Failed:
End Function
