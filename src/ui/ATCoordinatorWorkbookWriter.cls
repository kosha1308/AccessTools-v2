VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATCoordinatorWorkbookWriter"
Option Explicit
Public Function Create(ByVal registry As ATNetworkRegistryResult) As Workbook
    Dim book As Workbook, summary As Worksheet, view As Worksheet, entries As Collection, entry As ATNetworkRegistryEntry
    Dim ids As New Collection, data As Variant, count As Long, spec As ATExcelTableSpec, enrichment As ATEvalResult
    Dim enricher As New ATSupersetEnricher, rows As ATRuntimeValue, row As ATRuntimeRow, refs As ATRuntimeValue
    Dim i As Long, j As Long, text As String, fields As Variant, status As String, reference As String
    Dim previousEvents As Boolean, armed As Boolean, code As String
    On Error GoTo Failed
    If registry Is Nothing Then Err.Raise vbObjectError + 3120, , "VIEW_INPUT_INVALID"
    If registry.State <> atEvalOK Then Err.Raise vbObjectError + 3120, , "VIEW_REGISTRY_BLOCKED"
    Set entries = registry.Entries
    If entries.Count > 10000 Then Err.Raise vbObjectError + 3120, , "VIEW_ROW_LIMIT"
    data = AT_ViewRows(entries, count)
    previousEvents = Application.EnableEvents: armed = True: Application.EnableEvents = False
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set summary = book.Worksheets(1): summary.Name = "Сводка координатора"
    Set view = book.Worksheets.Add(After:=summary): view.Name = "Superset"
    Prepare view, count + 5, 9, "Superset — найденные строки", "Снимок исходных значений. Связь со сводкой — по точным номерам в первом столбце."
    view.Range("A5:I5").Value2 = AT_ViewHeaders()
    If count > 0 Then view.Range("A6").Resize(count, 9).Value2 = data
    For Each entry In entries: ids.Add entry.Id: Next entry
    Set spec = AT_ViewSpec()
    Set enrichment = enricher.Read(view, spec.Locate(view), count + 5, ids)
    If enrichment.State <> atEvalOK Then Err.Raise vbObjectError + 3120, , enrichment.Diagnostics(1).Code
    Set rows = enrichment.Value
    Prepare summary, entries.Count + 5, 12, "Сводка координатора", "Сервисы и инструкции Rulebook ещё не применены. Ненайденные адреса и неполные результаты требуют ручной проверки."
    summary.Range("A5:L5").Value2 = Array("№", "IP/сеть из заявки", "Статус Superset", "Найдено строк", "Сервис", "Инструкция", "Имя хоста (Superset)", "Сервис (Superset)", "Полное имя ИС (Superset)", "СКУФ ГРУППА ТП (Superset)", "КОМАНДА ПОДДЕРЖКИ (Superset)", "Ячейки заявки")
    fields = Array("HOSTNAME", "SERVICE", "IS_NAME", "SKUF_GROUP", "SUPPORT_TEAM")
    For i = 1 To entries.Count
        Set entry = entries(i): Set row = rows.Row(i)
        status = "НАЙДЕН В SUPERSET"
        If entry.Status = "NOT_FOUND" Then status = "НЕ НАЙДЕН В SUPERSET"
        If entry.Status = "UNCONFIRMED" Then status = "НЕ ПОДТВЕРЖДЕНО"
        If entry.Status = "FOUND" And Not entry.SourceComplete Then status = "НАЙДЕН; ИСТОЧНИК НЕПОЛНЫЙ"
        summary.Cells(i + 5, 1).Resize(1, 4).Value2 = Array(entry.Id, entry.Address, status, row.ResolveField("MATCH_COUNT").Value.Scalar)
        For j = 0 To 4
            summary.Cells(i + 5, j + 7).Value2 = AT_ViewLiteral(AT_ViewText(row.ResolveField(CStr(fields(j)))))
        Next j
        Set refs = entry.References: text = ""
        For j = 1 To refs.Count
            reference = AT_ViewText(refs.Row(j).ResolveField("CELL_REF"))
            If Len(text) + Len(reference) + 1 > 32767 Then Err.Raise vbObjectError + 3120, , "VIEW_TEXT_LIMIT"
            If Len(text) > 0 Then text = text & vbLf
            text = text & reference
        Next j
        summary.Cells(i + 5, 12).Value2 = AT_ViewLiteral(text)
    Next i
    Finish view, count + 5, 9: Finish summary, entries.Count + 5, 12
    For i = 1 To entries.Count
        Set entry = entries(i)
        If entry.Status <> "FOUND" Or Not entry.SourceComplete Then summary.Cells(i + 5, 3).Interior.Color = RGB(255, 242, 204)
    Next i
    summary.Columns("A").ColumnWidth = 6: summary.Columns("B").ColumnWidth = 22
    summary.Columns("D").ColumnWidth = 12: summary.Columns("F").ColumnWidth = 42
    view.Columns("A").ColumnWidth = 15
    book.Names.Add Name:="ATVIEW_SUPERSET", RefersTo:="='Сводка координатора'!$A$1"
    summary.Activate: book.Windows(1).DisplayGridlines = False
    book.Windows(1).SplitRow = 5: book.Windows(1).FreezePanes = True
    summary.Range("A1").Select
    Application.EnableEvents = previousEvents
    Set Create = book: Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "VIEW_CREATE_FAILED")
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    If armed Then Application.EnableEvents = previousEvents
    On Error GoTo 0
    Err.Raise vbObjectError + 3120, "ATCoordinatorWorkbookWriter", code
End Function
Private Sub Prepare(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal columns As Long, ByVal title As String, ByVal note As String)
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, columns))
        .NumberFormat = "@": .Font.Name = "Calibri": .Font.Size = 11
        .Font.Color = RGB(34, 48, 64): .VerticalAlignment = xlTop: .WrapText = True
    End With
    ws.Columns(1).Resize(, columns).ColumnWidth = 29
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, columns))
        .Merge: .Cells(1, 1).Value2 = title: .Font.Size = 20: .Font.Bold = True: .RowHeight = 36
    End With
    With ws.Range(ws.Cells(2, 1), ws.Cells(2, columns))
        .Merge: .Cells(1, 1).Value2 = "AccessTools 2.0 • отдельный результат, исходные файлы не изменяются": .RowHeight = 24
    End With
    With ws.Range(ws.Cells(3, 1), ws.Cells(3, columns))
        .Merge: .Cells(1, 1).Value2 = note: .Interior.Color = RGB(255, 242, 204): .RowHeight = 32
    End With
End Sub
Private Sub Finish(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal columns As Long)
    Dim i As Long
    With ws.Range(ws.Cells(5, 1), ws.Cells(5, columns))
        .Interior.Color = RGB(23, 54, 77): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 48
    End With
    If lastRow > 5 Then
        ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, columns)).Rows.AutoFit
        For i = 6 To lastRow
            If i Mod 2 = 0 Then ws.Range(ws.Cells(i, 1), ws.Cells(i, columns)).Interior.Color = RGB(238, 245, 248)
        Next i
    End If
    ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, columns)).AutoFilter
    With ws.PageSetup
        .Orientation = xlLandscape: .PaperSize = xlPaperA3: .Zoom = False
        .FitToPagesWide = 1: .FitToPagesTall = False
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(IIf(lastRow > 5, lastRow, 6), columns)).Address
        .PrintTitleRows = "$1:$5"
    End With
End Sub
