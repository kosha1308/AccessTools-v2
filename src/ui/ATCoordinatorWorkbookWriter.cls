VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATCoordinatorWorkbookWriter"
Option Explicit
Public Function Create(ByVal registry As ATNetworkRegistryResult, Optional ByVal annotations As ATNetworkSupersetResult = Nothing, Optional ByVal technicalNotes As ATCoordinatorNotesResult = Nothing) As Workbook
    Dim book As Workbook, summary As Worksheet, view As Worksheet, entries As Collection, entry As ATNetworkRegistryEntry
    Dim ids As New Collection, data As Variant, count As Long, spec As ATExcelTableSpec, enrichment As ATEvalResult
    Dim enricher As New ATSupersetEnricher, rows As ATRuntimeValue, row As ATRuntimeRow, refs As ATRuntimeValue
    Dim i As Long, j As Long, text As String, fields As Variant, status As String, reference As String
    Dim previousEvents As Boolean, armed As Boolean, code As String
    Dim annotationRows As ATRuntimeValue, annotation As ATRuntimeRow, note As String
    Dim technicalRows As ATRuntimeValue, technicalText As String, baseText As String, omitted As Long, technicalStatus As String
    On Error GoTo Failed
    If registry Is Nothing Then Err.Raise vbObjectError + 3120, , "VIEW_INPUT_INVALID"
    If registry.State <> atEvalOK Then Err.Raise vbObjectError + 3120, , "VIEW_REGISTRY_BLOCKED"
    Set entries = registry.Entries
    If Not technicalNotes Is Nothing Then
        technicalNotes.ValidateRegistry registry
        If technicalNotes.State = atEvalOK Then
            Set technicalRows = technicalNotes.Rows
            If technicalRows.Count <> entries.Count Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_BINDING"
            For i = 1 To entries.Count
                If technicalRows.Row(i).RowId <> entries(i).Id Then Err.Raise vbObjectError + 3160, , "COORDINATOR_NOTES_BINDING"
            Next i
        End If
    End If
    If Not annotations Is Nothing Then
        annotations.ValidateRegistry registry
        If annotations.State = atEvalOK Then
            Set annotationRows = annotations.Annotations
            If annotationRows.Count <> entries.Count Then Err.Raise vbObjectError + 3140, , "SUPERSET_RULES_BINDING"
            For i = 1 To entries.Count
                If annotationRows.Row(i).RowId <> entries(i).Id Then Err.Raise vbObjectError + 3140, , "SUPERSET_RULES_BINDING"
            Next i
        End If
    End If
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
    note = "Сервисы и инструкции Rulebook ещё не применены. Ненайденные адреса и неполные результаты требуют ручной проверки."
    If Not annotations Is Nothing Then
        note = "Контексты Superset проверены по Rulebook. Неполные результаты требуют ручной проверки; технические порты и протоколы оцениваются отдельно."
        If annotations.State <> atEvalOK Then note = "Правила Superset не завершены. Результаты IP-поиска сохранены; требуется ручная проверка."
    End If
    If Not technicalNotes Is Nothing Then
        note = "Правила Superset ещё не применены."
        If Not annotations Is Nothing Then
            note = "Правила Superset обработаны."
            If annotations.State <> atEvalOK Then note = "Правила Superset не завершены."
        End If
        If technicalNotes.State = atEvalOK Then
            note = note & " Технические результаты перенесены с учётом маршрутизации."
            If technicalNotes.Incomplete Then note = note & " Техническая проверка неполная; см. отдельный отчёт проверок."
            If technicalNotes.GlobalCount > 0 Then note = note & " Непривязанных: " & CStr(technicalNotes.GlobalCount) & "; см. отдельный отчёт проверок."
        Else
            note = note & " Технические инструкции не перенесены; см. отдельный отчёт проверок."
        End If
    End If
    Prepare summary, entries.Count + 5, 12, "Сводка координатора", note
    summary.Range("A5:L5").Value2 = Array("№", "IP/сеть из заявки", "Статус Superset", "Найдено строк", "Сервис", "Инструкция", "Имя хоста (Superset)", "Сервис (Superset)", "Полное имя ИС (Superset)", "СКУФ ГРУППА ТП (Superset)", "КОМАНДА ПОДДЕРЖКИ (Superset)", "Ячейки заявки")
    fields = Array("HOSTNAME", "SERVICE", "IS_NAME", "SKUF_GROUP", "SUPPORT_TEAM")
    For i = 1 To entries.Count
        Set entry = entries(i): Set row = rows.Row(i)
        status = "НАЙДЕН В SUPERSET"
        If entry.Status = "NOT_FOUND" Then status = "НЕ НАЙДЕН В SUPERSET"
        If entry.Status = "UNCONFIRMED" Then status = "НЕ ПОДТВЕРЖДЕНО"
        If entry.Status = "FOUND" And Not entry.SourceComplete Then status = "НАЙДЕН; ИСТОЧНИК НЕПОЛНЫЙ"
        summary.Cells(i + 5, 1).Resize(1, 4).Value2 = Array(entry.Id, entry.Address, status, row.ResolveField("MATCH_COUNT").Value.Scalar)
        If Not annotationRows Is Nothing Then
            Set annotation = annotationRows.Row(i)
            summary.Cells(i + 5, 5).Value2 = AT_ViewLiteral(AT_ViewText(annotation.ResolveField("SERVICE_RESULT")))
            summary.Cells(i + 5, 6).Value2 = AT_ViewLiteral(AT_ViewText(annotation.ResolveField("INSTRUCTION")))
        ElseIf Not annotations Is Nothing Then
            summary.Cells(i + 5, 6).Value2 = "Правила Superset не завершены. Требуется ручная проверка."
        End If
        If Not technicalNotes Is Nothing Then
            technicalText = "Технические инструкции не перенесены. Требуется ручная проверка."
            If Not technicalRows Is Nothing Then technicalText = AT_ViewText(technicalRows.Row(i).ResolveField("INSTRUCTION"))
            If Len(technicalText) > 0 Then
                baseText = AT_ExcelText(summary.Cells(i + 5, 6).Value2)
                If Len(baseText) + Len(technicalText) + 2 > 32767 Then
                    omitted = omitted + 1
                Else
                    If Len(baseText) > 0 Then baseText = baseText & vbLf & vbLf
                    summary.Cells(i + 5, 6).Value2 = AT_ViewLiteral(baseText & technicalText)
                End If
            End If
        End If
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
    If omitted > 0 Then summary.Cells(3, 1).Value2 = note & " Не перенесено из-за длины: " & CStr(omitted) & "; см. отдельный отчёт проверок."
    Finish view, count + 5, 9: Finish summary, entries.Count + 5, 12
    For i = 1 To entries.Count
        Set entry = entries(i)
        If entry.Status <> "FOUND" Or Not entry.SourceComplete Then summary.Cells(i + 5, 3).Interior.Color = RGB(255, 242, 204)
        If Not annotationRows Is Nothing Then
            status = AT_ViewText(annotationRows.Row(i).ResolveField("STATUS"))
            If status = "SKIPPED" Then summary.Cells(i + 5, 6).Interior.Color = RGB(255, 242, 204)
            If status = "FAIL" Then summary.Cells(i + 5, 6).Interior.Color = RGB(252, 228, 228)
        End If
        If Not technicalNotes Is Nothing Then
            technicalStatus = "WARNING"
            If Not technicalRows Is Nothing Then technicalStatus = AT_ViewText(technicalRows.Row(i).ResolveField("STATUS"))
            If technicalStatus = "FAIL" Then
                summary.Cells(i + 5, 6).Interior.Color = RGB(252, 228, 228)
            ElseIf technicalStatus = "WARNING" Then
                If summary.Cells(i + 5, 6).Interior.Color <> RGB(252, 228, 228) Then summary.Cells(i + 5, 6).Interior.Color = RGB(255, 242, 204)
            End If
        End If
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
