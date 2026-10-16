Attribute VB_Name = "modATTemplateGateTests"
Option Explicit

Public Sub AT_TestTemplateProfiles()
    Dim book As Workbook, ws As Worksheet, profile As ATTemplateProfile, decision As ATTemplateGateResult
    Dim saved As ATTemplateGateResult, caught As Boolean, savedNumber As Long, savedText As String
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Request"
    AT_TestTemplateSheet ws
    Set profile = AT_TestTemplateProfile()
    Set decision = profile.Verify(ws)
    AT_Assert decision.Allowed And decision.Code = "TEMPLATE_MATCH", "template reference structural profile passes"
    Set saved = decision
    ws.Range("F5").Value2 = "192.0.2.1": ws.Range("F6").Formula = "=""free text"""
    ws.Range("F5").Interior.Color = vbYellow
    Set decision = profile.Verify(ws)
    AT_Assert decision.Allowed, "user data and cosmetic style not fingerprint inputs"
    AT_Assert ws.Range("F6").Formula = "=""free text""" And ws.Range("F5").Interior.Color = vbYellow, "template verification preserves data and styles"
    ws.Range("A15").Value2 = "old lower table"
    Set decision = profile.Verify(ws)
    AT_Assert Not decision.Allowed And decision.Code = "TEMPLATE_MISMATCH", "changed mandatory header rejected"
    AT_Assert decision.Message = "Используйте актуальный шаблон заявки.", "one localized structural mismatch message"
    ws.Range("A15").Value2 = "Address"
    ws.Range("D2").Value2 = "Form v1"
    AT_Assert Not profile.Verify(ws).Allowed, "old explicit version marker rejected"
    ws.Range("D2").Value2 = "Form v2"
    ws.Range("A8:C8").UnMerge
    AT_Assert Not profile.Verify(ws).Allowed, "merge shape change rejected"
    ws.Range("A8:C8").Merge
    ws.Range("A15:D15").UnMerge: ws.Range("A15").ClearContents
    ws.Range("A20:D20").Merge: ws.Range("A20").Value2 = "Address"
    AT_Assert profile.Verify(ws).Allowed, "lower table can move down after data rows added"
    ws.Range("A20:D20").UnMerge: ws.Range("A20").ClearContents
    ws.Range("B20:E20").Merge: ws.Range("B20").Value2 = "Address"
    AT_Assert Not profile.Verify(ws).Allowed, "relative column relationship enforced"
    ws.Range("B20:E20").UnMerge: ws.Range("B20").ClearContents
    ws.Range("A15:D15").Merge: ws.Range("A15").Value2 = "Address"
    ws.Range("A4").Formula = "=""Initiator"""
    AT_Assert Not profile.Verify(ws).Allowed, "formula cannot impersonate protected literal header"
    ws.Range("A4").Value2 = "Initiator"
    ws.Range("A7").Value2 = "IP source"
    AT_Assert Not profile.Verify(ws).Allowed, "ambiguous mandatory anchor rejected"
    ws.Range("A7").ClearContents
    ws.Range("A4:B4").Locked = False
    AT_Assert Not profile.Verify(ws).Allowed, "required header locking enforced"
    ws.Range("A4:B4").Locked = True
    Set profile = AT_TestTemplateProfile(True)
    AT_Assert Not profile.Verify(ws).Allowed, "required sheet protection enforced"
    ws.Protect Password:="fixture"
    AT_Assert profile.Verify(ws).Allowed, "protected reference verifies read only"
    AT_Assert ws.ProtectContents, "verification preserves protection"
    ws.Unprotect Password:="fixture"
    ws.Name = "Other"
    AT_Assert Not profile.Verify(ws).Allowed, "wrong worksheet rejected"
    ws.Name = "Request"
    AT_Assert profile.Verify(Nothing).Code = "TEMPLATE_CHECK_ERROR", "missing worksheet fails closed"
    On Error Resume Next
    profile.AddRelation "PERSON", "LOWER", 1, 50, 0, 0
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo CleanupError
    AT_Assert caught, "sealed profile cannot acquire new relations"
    AT_TestBadTemplateProfiles ws
    book.Close False
    AT_Assert saved.Allowed And saved.Code = "TEMPLATE_MATCH", "template decision detached from later edits and source closure"
    Exit Sub
CleanupError:
    savedNumber = Err.Number: savedText = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise savedNumber, , savedText
End Sub

Public Sub AT_TestTemplateSheet(ByVal ws As Worksheet)
    ws.Range("D2:E2").Merge: ws.Range("D2").Value2 = "Form v2"
    ws.Range("A4:B4").Merge: ws.Range("A4").Value2 = "Initiator"
    ws.Range("A8:C8").Merge: ws.Range("A8").Value2 = "IP source"
    ws.Range("A15:D15").Merge: ws.Range("A15").Value2 = "Address"
End Sub
Public Function AT_TestTemplateProfile(Optional ByVal protection As Boolean = False) As ATTemplateProfile
    Dim profile As New ATTemplateProfile
    profile.Initialize "NETWORK_FIXTURE", "2", "Request", protection
    profile.AddAnchor "VERSION", "Form v2", 1, 3, 4, 5, 1, 2
    profile.AddAnchor "PERSON", "Initiator", 4, 4, 1, 2, 1, 2
    profile.AddAnchor "TECH", "IP source", 7, 9, 1, 3, 1, 3
    profile.AddAnchor "LOWER", "Address", 10, 30, 1, 6, 1, 4
    profile.AddRelation "VERSION", "PERSON", 2, 2, -3, -3
    profile.AddRelation "PERSON", "TECH", 4, 4, 0, 0
    profile.AddRelation "TECH", "LOWER", 1, 20, 0, 0
    profile.Seal
    Set AT_TestTemplateProfile = profile
End Function
Private Sub AT_TestBadTemplateProfiles(ByVal ws As Worksheet)
    Dim profile As New ATTemplateProfile, caught As Boolean
    Dim i As Long, headers As Variant, entry As Variant
    AT_Assert Not profile.Verify(ws).Allowed And profile.Verify(ws).Code = "TEMPLATE_PROFILE_INVALID", "uninitialized profile denies entry"
    profile.Initialize "T", "1", "Request"
    profile.AddAnchor "A", "Form v2", 1, 3, 4, 5, 1, 2
    On Error Resume Next
    profile.Seal
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught And Not profile.Verify(ws).Allowed, "one fact is not a complete fingerprint"
    profile.AddAnchor "B", "Form v2", 1, 3, 4, 5, 1, 2
    profile.AddRelation "A", "B", 0, 0, 0, 0
    profile.Seal
    AT_Assert Not profile.Verify(ws).Allowed, "two ids resolving to one actual fact rejected"
    Set profile = New ATTemplateProfile
    profile.Initialize "T", "1", "Request"
    profile.AddAnchor "A", "A", 1, 10000, 1, 64, 1, 1
    On Error Resume Next
    profile.AddAnchor "B", "B", 1, 10000, 1, 64, 1, 1
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught, "total profile scan budget enforced before reading"
    On Error Resume Next
    profile.AddAnchor "BAD", "B", 1, 1, 1, 1, 0, 1
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught, "invalid merge shape rejected"
    Set profile = New ATTemplateProfile
    profile.Initialize "T", "1", "Request"
    profile.AddAnchor "A", "A", 1, 1, 1, 1, 1, 1
    profile.AddAnchor "B", "B", 2, 2, 1, 1, 1, 1
    profile.AddRelation "A", "UNKNOWN", 0, 5, 0, 0
    On Error Resume Next
    profile.Seal
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught And Not profile.Verify(ws).Allowed, "unknown relation end cannot publish profile"
    Set profile = New ATTemplateProfile
    profile.Initialize "T", "1", "Request"
    For i = 1 To 256
        profile.AddAnchor "A" & CStr(i), "L" & CStr(i), 1, 1, 1, 1, 1, 1
    Next i
    On Error Resume Next
    profile.AddAnchor "OVER", "OVER", 1, 1, 1, 1, 1, 1
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught, "profile anchor count bounded"
    For i = 1 To 1024
        profile.AddRelation "A1", "A2", 0, 5, 0, 0
    Next i
    On Error Resume Next
    profile.AddRelation "A1", "A2", 0, 5, 0, 0
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught, "profile relation count bounded"
    Set profile = New ATTemplateProfile
    profile.Initialize "COPY", "1", "Request"
    profile.AddAnchor "TITLE", "Form v2", 1, 3, 4, 5, 1, 2
    headers = Array(Array("Left", 1, 1), Array("Right", 2, 1))
    profile.AddTableAnchor "TABLE", 25, 25, 1, 2, headers
    entry = headers(0): entry(0) = "Changed": headers(0) = entry
    profile.AddRelation "TITLE", "TABLE", 23, 23, -3, -3
    profile.Seal
    ws.Range("A25").Value2 = "Left": ws.Range("B25").Value2 = "Right"
    AT_Assert profile.Verify(ws).Allowed, "group metadata copied independently of caller arrays"
End Sub
