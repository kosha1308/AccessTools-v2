VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATExcelReportWriter"
Option Explicit
Public Function Create(ByVal report As ATCheckReport, ByVal source As Worksheet) As Workbook
    Dim book As Workbook, view As Worksheet, checks As Worksheet, previousEvents As Boolean, armed As Boolean
    Dim issues As Collection, outcomes As Collection, issue As ATCheckIssue, outcome As ATCheckOutcome
    Dim i As Long, row As Long, finalRow As Long, ref As String, target As String
    On Error GoTo Failed
    If report Is Nothing Then Err.Raise vbObjectError + 2960, , "REPORT_REQUIRED"
    Set issues = report.Issues: Set outcomes = report.Checks
    previousEvents = Application.EnableEvents: armed = True: Application.EnableEvents = False
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set view = book.Worksheets(1): view.Name = "Отчёт"
    Set checks = book.Worksheets.Add(After:=view): checks.Name = "Проверки"
    finalRow = 10 + issues.Count
    Prepare view, finalRow, 6
    view.Columns("A").ColumnWidth = 5: view.Columns("B").ColumnWidth = 19
    view.Columns("C").ColumnWidth = 31: view.Columns("D").ColumnWidth = 26
    view.Columns("E").ColumnWidth = 85: view.Columns("F").ColumnWidth = 25
    MergedText view.Range("A1:F1"), "AccessTools 2.0 — результаты проверки", 18, True
    view.Rows(1).RowHeight = 34
    MergedText view.Range("A2:F2"), "Отдельный отчёт • исходная заявка не изменялась", 11, False
    target = "Проверена только часть заявки. Superset и сводка координатора не выполнялись. Пропущенные проверки перечислены на листе «Проверки»."
    If report.ScopeId = "NETWORK_ACCESS_REGISTRY_POLICY" Or report.ScopeId = "NETWORK_REGISTRY" Then target = "Это промежуточный отчёт с результатом IP-поиска. Итоговая копия Superset и сводка координатора ещё не формировались. Пропущенные проверки перечислены на листе «Проверки»."
    If report.ScopeId = "NETWORK_ACCESS_SUPERSET_POLICY" Then target = "Отчёт включает IP-поиск и правила Superset. Координаторская книга формируется отдельно. Пропуски и неполные результаты требуют ручной проверки."
    MergedText view.Range("A3:F4"), target, 11, False
    view.Range("A3:F4").Interior.Color = RGB(255, 242, 204): view.Rows("3:4").RowHeight = 23
    MergedText view.Range("A6:F6"), "Ошибок: " & CStr(report.ErrorCount) & "  •  предупреждений: " & CStr(report.WarningCount) & "  •  пропущено проверок: " & CStr(report.SkippedCount), 12, True
    view.Rows(6).RowHeight = 27
    target = "Локальные проверки полей и таблиц"
    If report.ScopeId = "NETWORK_DOCUMENT" Then target = "Проверки одиночных полей"
    If report.ScopeId = "NETWORK_TABLE_INPUT" Then target = "Проверки заполнения таблиц"
    If report.ScopeId = "NETWORK_ACCESS_REGISTRY_POLICY" Then target = "Поля, таблицы, правила и поиск IP в Superset"
    If report.ScopeId = "NETWORK_ACCESS_SUPERSET_POLICY" Then target = "Поля, таблицы, технические правила, IP и правила Superset"
    MergedText view.Range("A7:F7"), target, 11, False
    view.Range("A9:F9").Value2 = Array("№", "Важность", "Проверка", "Поле", "Сообщение", "Ячейка")
    Header view.Range("A9:F9")
    For i = 1 To issues.Count
        Set issue = issues(i): row = i + 9
        view.Cells(row, 1).Resize(1, 6).Value2 = Array(CStr(i), AT_SeverityTitle(issue.Severity), AT_CheckTitle(issue.CheckId), AT_FieldTitle(issue.FieldId), Literal(issue.Message), Literal(issue.CellRef))
        Shade view.Cells(row, 1).Resize(1, 6), i
        If issue.Severity = "ERROR" Then
            view.Cells(row, 2).Interior.Color = RGB(252, 228, 228): view.Cells(row, 2).Font.Color = RGB(156, 0, 6)
        ElseIf issue.Severity = "WARNING" Then
            view.Cells(row, 2).Interior.Color = RGB(255, 242, 204): view.Cells(row, 2).Font.Color = RGB(125, 78, 0)
        End If
        ref = TrustedAddress(source, issue.CellRef)
        If Len(ref) > 0 Then
            If Len(source.Parent.Path) > 0 Then
                view.Hyperlinks.Add Anchor:=view.Cells(row, 6), Address:=source.Parent.FullName, SubAddress:=issue.CellRef, TextToDisplay:=Literal(issue.CellRef)
            Else
                target = "'[" & Replace(source.Parent.Name, "'", "''") & "]" & Replace(source.Name, "'", "''") & "'!" & ref
                view.Hyperlinks.Add Anchor:=view.Cells(row, 6), Address:="", SubAddress:=target, TextToDisplay:=Literal(issue.CellRef)
            End If
        End If
    Next i
    If issues.Count = 0 Then MergedText view.Range("A10:F10"), "По выполненным проверкам ошибок и предупреждений нет. Остальные проверки заявки ещё не выполнялись.", 11, False
    If issues.Count > 0 Then view.Range("A10:F" & CStr(issues.Count + 9)).Rows.AutoFit
    view.Range("A9:F" & CStr(9 + issues.Count)).AutoFilter
    Prepare checks, 5 + outcomes.Count, 3
    checks.Columns("A").ColumnWidth = 34: checks.Columns("B").ColumnWidth = 21: checks.Columns("C").ColumnWidth = 105
    MergedText checks.Range("A1:C1"), "Выполненные проверки", 18, True: checks.Rows(1).RowHeight = 34
    MergedText checks.Range("A2:C2"), "Статус относится только к перечисленным проверкам, а не ко всей заявке.", 11, False
    checks.Range("A4:C4").Value2 = Array("Проверка", "Результат", "Причина / пояснение"): Header checks.Range("A4:C4")
    For i = 1 To outcomes.Count
        Set outcome = outcomes(i): row = i + 4
        checks.Cells(row, 1).Resize(1, 3).Value2 = Array(AT_CheckTitle(outcome.CheckId), AT_StatusTitle(outcome.Status), Literal(outcome.Reason))
        Shade checks.Cells(row, 1).Resize(1, 3), i
    Next i
    If outcomes.Count > 0 Then checks.Range("A5:C" & CStr(outcomes.Count + 4)).Rows.AutoFit
    PrintLayout view, 6, finalRow, "$9:$9"
    PrintLayout checks, 3, 5 + outcomes.Count, "$4:$4"
    checks.Activate: book.Windows(1).DisplayGridlines = False
    view.Activate: book.Windows(1).DisplayGridlines = False: view.Range("A1").Select
    Application.EnableEvents = previousEvents: Set Create = book
    Exit Function
Failed:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    If armed Then Application.EnableEvents = previousEvents
    On Error GoTo 0
    Err.Raise vbObjectError + 2960, "ATExcelReportWriter", "REPORT_CREATE_FAILED"
End Function
Private Function Literal(ByVal text As String) As String
    If Left$(text, 1) = "'" Then text = "'" & text
    Literal = text
End Function
Private Sub Prepare(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastColumn As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn))
        .NumberFormat = "@": .Font.Name = "Calibri": .Font.Size = 11
        .Font.Color = RGB(34, 48, 64): .VerticalAlignment = xlTop: .WrapText = True
        .RowHeight = 23
    End With
End Sub
Private Sub MergedText(ByVal area As Range, ByVal text As String, ByVal size As Long, ByVal bold As Boolean)
    area.Merge: area.Cells(1, 1).Value2 = text: area.Font.Size = size: area.Font.Bold = bold
End Sub
Private Sub Header(ByVal area As Range)
    area.Interior.Color = RGB(23, 54, 77): area.Font.Color = RGB(255, 255, 255)
    area.Font.Bold = True: area.RowHeight = 28
End Sub
Private Sub Shade(ByVal area As Range, ByVal index As Long)
    If index Mod 2 = 1 Then area.Interior.Color = RGB(238, 245, 248)
    area.Borders(xlEdgeBottom).Color = RGB(219, 229, 235)
End Sub
Private Sub PrintLayout(ByVal ws As Worksheet, ByVal lastColumn As Long, ByVal lastRow As Long, ByVal repeatRows As String)
    With ws.PageSetup
        .Orientation = xlLandscape: .PaperSize = xlPaperA4: .Zoom = False
        .FitToPagesWide = 1: .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(0.3): .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.4): .BottomMargin = Application.InchesToPoints(0.4)
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn)).Address
        .PrintTitleRows = repeatRows
    End With
End Sub
Private Function TrustedAddress(ByVal source As Worksheet, ByVal ref As String) As String
    Dim prefix As String, address As String, i As Long, ch As String, digits As Boolean, letters As Long
    On Error GoTo Invalid
    If source Is Nothing Then Exit Function
    prefix = "'" & Replace(source.Name, "'", "''") & "'!"
    If Left$(ref, Len(prefix)) <> prefix Then Exit Function
    address = Mid$(ref, Len(prefix) + 1)
    If Len(address) < 2 Or Len(address) > 10 Then Exit Function
    For i = 1 To Len(address)
        ch = Mid$(address, i, 1)
        If ch >= "A" And ch <= "Z" And Not digits Then
            letters = letters + 1
            If letters > 3 Then Exit Function
        ElseIf ch >= "0" And ch <= "9" And letters > 0 Then
            digits = True
        Else
            Exit Function
        End If
    Next i
    If Not digits Then Exit Function
    If source.Range(address).Address(False, False) <> address Then Exit Function
    TrustedAddress = address
Invalid:
End Function
