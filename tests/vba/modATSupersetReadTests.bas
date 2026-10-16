Attribute VB_Name = "modATSupersetReadTests"
Option Explicit
Public Sub AT_TestSupersetRead()
    Dim book As Workbook, ws As Worksheet, layout As ATExcelTableLayout, spec As ATExcelTableSpec
    Dim enricher As New ATSupersetEnricher, result As ATEvalResult, saved As ATRuntimeValue
    Dim ids As Collection, savedNumber As Long, savedText As String, i As Long
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1)
    ws.Range("A1:F1").Value2 = Array("№ из сводки", "FQDN", "Сервис", "Полное имя ИС", "  СКУФ   ГРУППА ТП  ", " КОМАНДА  ПОДДЕРЖКИ ")
    ws.Range("A2:F2").Value2 = Array("1, 2, 1", "one.example", "svcA", "systemA", "Group A", "Team A")
    ws.Range("A3:F3").Value2 = Array("1;2", "two.example", "svcB", "systemB", "Group B", "Team B")
    ws.Range("A4:F4").Value2 = Array("bad", "hidden.example", "hidden", "hidden", "hidden", "hidden")
    ws.Range("A5:F5").Value2 = Array("2", "manual.example", "manual", "manual", "manual", "manual")
    ws.Range("A6:F6").Value2 = Array("10", "other.example", "other", "other", "other", "other")
    ws.Range("F2").Formula = "=""Team A"""
    ws.Range("F4").Value2 = CVErr(xlErrNA): ws.Range("F6").Value2 = CVErr(xlErrValue)
    ws.Range("A1:F6").AutoFilter Field:=3, Criteria1:="<>hidden"
    ws.Rows(5).Hidden = True
    Set spec = AT_TestSupersetSpec()
    Set layout = spec.Locate(ws): Set ids = AT_TestSupersetIds(Array("1", "2", "3"))
    Set result = enricher.Read(ws, layout, 6, ids)
    AT_Assert result.State = atEvalOK, "Superset visible linked reader ignores hidden and unrelated errors"
    If result.State <> atEvalOK Then GoTo Finish
    Set saved = result.Value
    AT_Assert saved.Count = 3 And saved.Row(3).RowId = "3", "all requested summaries retained in order"
    AT_Assert saved.Row(1).ResolveField("MATCH_COUNT").Value.Scalar = 2#, "duplicate numbers in one source row counted once"
    AT_Assert saved.Row(2).ResolveField("MATCH_COUNT").Value.Scalar = 2#, "one source row can relate to multiple summaries"
    AT_Assert saved.Row(3).ResolveField("MATCH_COUNT").Value.Scalar = 0#, "unmatched summary remains visible"
    AT_Assert saved.Row(1).ResolveField("SKUF_GROUP").Value.Scalar = "Group A" & vbLf & "Group B", "new SKUF field from ordinary visible source column"
    AT_Assert saved.Row(2).ResolveField("SUPPORT_TEAM").Value.Scalar = "Team A" & vbLf & "Team B", "new support field all visible linked rows"
    AT_Assert saved.Row(1).ResolveField("SERVICE").Value.Scalar = "svcA" & vbLf & "svcB", "service aggregation retained"
    AT_Assert saved.Row(1).ResolveField("IS_NAME").Value.Scalar = "systemA" & vbLf & "systemB", "full IS aggregation retained"
    AT_Assert saved.Row(1).ResolveField("HOSTNAME").Value.ValueState = atValueEmpty, "multiple source hostnames do not choose arbitrary one"
    AT_Assert saved.Row(3).ResolveField("SUPPORT_TEAM").Value.ValueState = atValueEmpty, "no match has no Superset enrichment"
    AT_Assert ws.FilterMode And ws.Rows(4).Hidden And ws.Rows(5).Hidden, "reader preserves filter and hidden state"
    AT_Assert ws.Range("F2").Formula = "=""Team A""" And ws.Range("E1").Value2 = "  СКУФ   ГРУППА ТП  ", "reader preserves formula and source headers"
    ws.Rows(3).Hidden = True
    Set result = enricher.Read(ws, layout, 6, ids)
    AT_Assert result.Value.Row(1).ResolveField("HOSTNAME").Value.Scalar = "one.example", "new visibility applied without stale cache"
    AT_Assert result.Value.Row(1).ResolveField("MATCH_COUNT").Value.Scalar = 1#, "visibility change updates row count"
    ws.Rows(3).Hidden = False
    ws.Range("F2").Value2 = CVErr(xlErrValue)
    Set result = enricher.Read(ws, layout, 6, ids)
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "EXCEL_CELL_ERROR", "relevant visible field error stops incomplete summary"
    ws.Range("F2").Value2 = "Team A"
    ws.Range("A2").Value2 = "1,bad"
    Set result = enricher.Read(ws, layout, 6, ids)
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "SUPERSET_LINK_INVALID", "malformed visible link cannot be partly salvaged"
    ws.Range("A2").Value2 = "01," & vbLf & "1;2"
    Set result = enricher.Read(ws, layout, 6, ids)
    AT_Assert result.State = atEvalOK, "adapter reusable after error and canonicalizes leading zero"
    AT_Assert result.Value.Row(1).ResolveField("MATCH_COUNT").Value.Scalar = 2#, "canonical duplicate source links deduped"
    Set result = enricher.Read(ws, layout, 6, AT_TestSupersetIds(Array("1", "01")))
    AT_Assert result.Diagnostics(1).Code = "SUPERSET_SUMMARY_IDS", "duplicate requested summary ids rejected"
    Set result = enricher.Read(ws, layout, 6, AT_TestSupersetIds(Array("1.5")))
    AT_Assert result.Diagnostics(1).Code = "SUPERSET_SUMMARY_IDS", "fractional ids rejected"
    ws.ShowAllData: ws.Rows(5).Hidden = False: ws.Range("A2:F6").ClearContents
    ws.Range("A2:F2").Value2 = Array("10", "host10", "service10", "system10", "group10", "team10")
    Set result = enricher.Read(ws, layout, 2, AT_TestSupersetIds(Array("1")))
    AT_Assert result.Value.Row(1).ResolveField("MATCH_COUNT").Value.Scalar = 0#, "summary1 is not substring of10"
    For i = 2 To 14
        ws.Cells(i, 1).Value2 = Replace(String$(16000, "x"), "x", "1,")
    Next i
    Set result = enricher.Read(ws, layout, 14, AT_TestSupersetIds(Array("1")))
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "SUPERSET_WORK_LIMIT", "link token processing bounded"
    ws.Range("A2:F14").ClearContents
    Set result = enricher.Read(ws, layout, 1, AT_TestSupersetIds(Array("1")))
    AT_Assert result.State = atEvalOK And result.Value.Count = 1, "empty source and reuse after work limit"
    Set result = enricher.Read(ws, layout, 1, AT_TestSupersetIds(Array(String$(32000, "0") & "1")))
    AT_Assert result.State = atEvalOK And result.Value.Row(1).RowId = "1", "long zero prefix parsed without number overflow"
    AT_TestSupersetBadInputs enricher, ws, layout
Finish:
    book.Close False
    If Not saved Is Nothing Then AT_Assert saved.Row(2).ResolveField("SUPPORT_TEAM").Value.Scalar = "Team A" & vbLf & "Team B", "enrichment snapshot survives source edits and closure"
    Exit Sub
CleanupError:
    savedNumber = Err.Number: savedText = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise savedNumber, , savedText
End Sub
Public Function AT_TestSupersetIds(ByVal values As Variant) As Collection
    Dim result As New Collection, item As Variant
    For Each item In values
        result.Add item
    Next item
    Set AT_TestSupersetIds = result
End Function
Private Sub AT_TestSupersetBadInputs(ByVal enricher As ATSupersetEnricher, ByVal ws As Worksheet, ByVal layout As ATExcelTableLayout)
    Dim ids As New Collection, result As ATEvalResult, item As Variant, i As Long
    Dim spec As ATExcelTableSpec, bad As ATExcelTableLayout
    For Each item In Array("0", "-1", "2147483648", "1e3", True, 1.5, String$(32768, "0"))
        Set result = enricher.Read(ws, layout, 1, AT_TestSupersetIds(Array(item)))
        AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "SUPERSET_SUMMARY_IDS", "invalid requested id fails closed"
    Next item
    For i = 1 To 10001
        ids.Add CStr(i)
    Next i
    Set result = enricher.Read(ws, layout, 1, ids)
    AT_Assert result.Diagnostics(1).Code = "SUPERSET_SUMMARY_IDS", "requested summary count bounded"
    Set spec = AT_TestExcelSpec("BAD", 1, 1, 1, 6)
    spec.AddColumn "SUMMARY_LINK", "№ из сводки", "DIRECT"
    Set bad = spec.Locate(ws)
    Set result = enricher.Read(ws, bad, 1, AT_TestSupersetIds(Array("1")))
    AT_Assert result.Diagnostics(1).Code = "SUPERSET_SCHEMA", "missing source enrichment field fails schema"
End Sub
Private Function AT_TestSupersetSpec() As ATExcelTableSpec
    Dim spec As ATExcelTableSpec
    Set spec = AT_TestExcelSpec("SUPERSET", 1, 1, 1, 6)
    spec.AddColumn "SUMMARY_LINK", "№ из сводки", "DIRECT"
    spec.AddColumn "HOSTNAME", "FQDN", "NONE"
    spec.AddColumn "SERVICE", "Сервис", "NONE"
    spec.AddColumn "IS_NAME", "Полное имя ИС", "NONE"
    spec.AddColumn "SKUF_GROUP", "СКУФ ГРУППА ТП", "NONE"
    spec.AddColumn "SUPPORT_TEAM", "КОМАНДА ПОДДЕРЖКИ", "NONE"
    Set AT_TestSupersetSpec = spec
End Function
