Attribute VB_Name = "modATTemplateReadTests"
Option Explicit
Public Sub AT_TestTemplateRead()
    Dim book As Workbook, ws As Worksheet, profile As ATTemplateProfile, spec As ATExcelTableSpec
    Dim layout As ATExcelTableLayout, gate As New ATTemplateTableGate, result As ATEvalResult
    Dim number As Long, description As String
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Request"
    AT_TestTemplateSheet ws: Set profile = AT_TestTemplateProfile()
    Set spec = AT_TestExcelSpec("TECH", 8, 8, 1, 3)
    spec.AddColumn "SRC_RAW", "IP source", "DIRECT": Set layout = spec.Locate(ws)
    ws.Range("A9").Value2 = "192.0.2.1"
    Set result = gate.Read(profile, ws, spec, 9)
    AT_Assert result.State = atEvalOK And result.Value.Count = 1, "valid structural profile allows table read"
    ws.Range("D2").Value2 = "Form v1"
    Set result = gate.Read(profile, ws, spec, 0)
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "TEMPLATE_MISMATCH", "template gate precedes invalid data bounds"
    AT_Assert result.Diagnostics.Count = 1 And result.Diagnostics(1).Message = "Используйте актуальный шаблон заявки.", "one localized rejection without business errors"
    ws.Range("D2").Value2 = "Form v2"
    Set result = gate.Read(profile, ws, spec, 9)
    AT_Assert result.State = atEvalOK, "guard rechecks current structure after rejection"
    Set result = gate.Read(Nothing, ws, spec, 9)
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "TEMPLATE_PROFILE_INVALID", "missing profile cannot bypass verification"
    Set result = gate.Read(profile, ws, Nothing, 9)
    AT_Assert result.Diagnostics(1).Code = "EXCEL_SPEC_REQUIRED", "missing table spec controlled after valid template"
    Set result = gate.Read(profile, ws, spec, 0)
    AT_Assert result.Diagnostics(1).Code = "EXCEL_DATA_BOUNDS", "valid template does not hide reader bounds error"
    AT_TestDynamicGuard book
    book.Close False
    Exit Sub
CleanupError:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise number, , description
End Sub
Private Sub AT_TestDynamicGuard(ByVal book As Workbook)
    Dim ws As Worksheet, profile As ATTemplateProfile, spec As ATExcelTableSpec
    Dim layout As ATExcelTableLayout, gate As New ATTemplateTableGate, result As ATEvalResult
    Set ws = book.Worksheets.Add: ws.Name = "Лист1": AT_TestNetworkReferenceFixture ws
    Set profile = AT_NetworkProfileV3()
    Set spec = AT_TestExcelSpec("ADDRESS", 15, 10000, 2, 2)
    spec.AddColumn "ADDRESS_RAW", "Адрес", "DIRECT"
    Set layout = spec.Locate(ws)
    ws.Range("B17").Value2 = "192.0.2.4"
    ws.Rows(15).Insert
    Set result = gate.Read(profile, ws, spec, 18)
    AT_Assert result.State = atEvalOK, "shifted valid form can be read after verification"
    If result.State = atEvalOK Then
        AT_Assert result.Value.Count = 1, "guard never returns moved actual header as business data"
        AT_Assert result.Value.Row(1).RowId = "18", "guard uses fresh lower table boundary"
    End If
End Sub
