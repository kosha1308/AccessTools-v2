Attribute VB_Name = "modATReportWriterTests"
Option Explicit
Public Sub AT_TestReportWriter()
    Dim book As Workbook, ws As Worksheet, rendered As Workbook, repeated As Workbook
    Dim reader As New ATNetworkRequestReader, checks As New ATNetworkLocalChecks, writer As New ATExcelReportWriter
    Dim report As ATCheckReport, view As Worksheet, originalEvents As Boolean, caught As Boolean
    Dim issue As New ATCheckIssue, outcome As New ATCheckOutcome, malicious As New ATCheckReport
    Dim issues As New Collection, outcomes As New Collection, number As Long, description As String
    On Error GoTo CleanupError
    originalEvents = Application.EnableEvents
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    ws.Range("C7").Value2 = "Предоставить доступ": ws.Range("C8").Value2 = "Да"
    Set report = checks.Run(reader.Read(ws, "WRITER", "2026-10-16", "UTC"), Nothing)
    Set rendered = writer.Create(report, ws): Set view = rendered.Worksheets("Отчёт")
    AT_Assert Not rendered Is book And rendered.Worksheets.Count = 2, "report in separate new workbook"
    AT_Assert view.Range("A1").Value2 = "AccessTools 2.0 — результаты проверки", "Russian report title"
    AT_Assert InStr(view.Range("A3").Value2, "Superset") > 0 And InStr(view.Range("A3").Value2, "часть") > 0, "partial scope disclaimer prominent"
    AT_Assert InStr(view.Range("A6").Value2, "Ошибок: " & CStr(report.ErrorCount)) > 0, "report counts match immutable result"
    AT_Assert view.Range("B9").Value2 = "Важность" And view.Range("E9").Value2 = "Сообщение", "Russian issue headers"
    AT_Assert view.Range("C10").Value2 = "Обязательные поля инициатора" And view.Range("D10").Value2 = "Фамилия Имя Отчество", "internal check and field IDs translated"
    AT_Assert view.Hyperlinks.Count = report.Issues.Count, "trusted refs produce source hyperlinks"
    AT_Assert view.Range("F10").Value2 = report.Issues(1).CellRef, "linked source ref text preserves opening quote"
    AT_Assert view.Hyperlinks(1).Address = "" And InStr(view.Hyperlinks(1).SubAddress, "[" & book.Name & "]") > 0, "unsaved open source linked explicitly"
    view.Hyperlinks(1).Follow
    AT_Assert Application.ActiveWorkbook Is book, "hyperlink follows actual source workbook"
    AT_Assert Application.ActiveCell.Address(False, False) = "A5", "hyperlink navigates actual source cell"
    AT_Assert book.Worksheets.Count = 1 And ws.Range("C7").Value2 = "Предоставить доступ", "writer preserves source book and values"
    AT_Assert Application.EnableEvents = originalEvents, "writer restores events"
    AT_Assert rendered.Worksheets("Проверки").Range("A4").Value2 = "Проверка", "outcome sheet Russian header"
    AT_Assert InStr(rendered.Worksheets("Проверки").Range("B5").Value2, "ошибки") > 0, "outcome status localized"
    Set repeated = writer.Create(report, Nothing)
    AT_Assert Not repeated Is rendered And repeated.Worksheets("Отчёт").Hyperlinks.Count = 0, "repeat creates fresh book; no source gives plain refs"
    AT_Assert repeated.Worksheets("Отчёт").Range("F10").Value2 = report.Issues(1).CellRef, "plain source ref text preserves opening quote"
    repeated.Close False: Set repeated = Nothing
    issue.Initialize "NET.REQUEST_TYPE", "TEST", "ERROR", "REQUEST_TYPE", "'OTHER'!A1", "=1+1"
    issues.Add issue: outcome.Initialize "NET.REQUEST_TYPE", "FAIL", "=1+1": outcomes.Add outcome
    malicious.Initialize atEvalOK, "NETWORK_LOCAL", issues, outcomes
    Application.EnableEvents = False
    Set repeated = writer.Create(malicious, ws)
    AT_Assert Not Application.EnableEvents, "previously disabled events remain disabled"
    AT_Assert Not repeated.Worksheets("Отчёт").Range("E10").HasFormula And repeated.Worksheets("Отчёт").Range("E10").Value2 = "=1+1", "formula-like issue written as literal text"
    AT_Assert Not repeated.Worksheets("Проверки").Range("C5").HasFormula, "formula-like outcome reason is text"
    AT_Assert repeated.Worksheets("Отчёт").Hyperlinks.Count = 0, "foreign sheet ref never becomes link"
    repeated.Close False: Set repeated = Nothing
    On Error Resume Next
    Set repeated = writer.Create(Nothing, ws)
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo CleanupError
    AT_Assert caught And Not Application.EnableEvents, "invalid report controlled and restores prior events"
    rendered.Close False: Set rendered = Nothing
    book.Close False: Set book = Nothing
    Application.EnableEvents = originalEvents
    Exit Sub
CleanupError:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not repeated Is Nothing Then repeated.Close False
    If Not rendered Is Nothing Then rendered.Close False
    If Not book Is Nothing Then book.Close False
    Application.EnableEvents = originalEvents
    On Error GoTo 0
    Err.Raise number, , description
End Sub
Public Function AT_ExportReportSample(ByVal folder As String) As String
    Dim writer As New ATExcelReportWriter, report As New ATCheckReport, book As Workbook
    Dim issue As ATCheckIssue, outcome As ATCheckOutcome, issues As New Collection, checks As New Collection, item As Variant
    For Each item In Array(Array("NET.INITIATOR_EMAIL", "EMAIL_INVALID", "ERROR", "EMAIL", "'Лист1'!F5", "Укажите корректную электронную почту в формате имя@домен."), Array("NET.TECHNICAL_TABLE", "PORT_INVALID", "ERROR", "TECH_DEST_PORT", "'Лист1'!D15", "Порт 0 не допускается в заявке. Укажите порт от 1 до 65535."), Array("NET.ADDRESS_TABLE", "FIELD_RECOMMENDED", "WARNING", "ROLE", "'Лист1'!E17", "Добавьте роль или назначение адреса."))
        Set issue = New ATCheckIssue: issue.Initialize CStr(item(0)), CStr(item(1)), CStr(item(2)), CStr(item(3)), CStr(item(4)), CStr(item(5)): issues.Add issue
    Next item
    For Each item In Array(Array("NET.INITIATOR_EMAIL", "FAIL", ""), Array("NET.TECHNICAL_TABLE", "FAIL", ""), Array("NET.PROTOCOLS", "SKIPPED", "Нет справочника разрешённых протоколов."), Array("NET.ADDRESS_TABLE", "PASS", "Есть предупреждение по назначению адреса."))
        Set outcome = New ATCheckOutcome: outcome.Initialize CStr(item(0)), CStr(item(1)), CStr(item(2)): checks.Add outcome
    Next item
    report.Initialize atEvalOK, "NETWORK_LOCAL", issues, checks
    Set book = writer.Create(report, Nothing)
    book.Worksheets("Отчёт").Range("A2").Value2 = "Пример отчёта • вымышленные данные • не рабочая версия приложения"
    book.SaveAs folder & "\AccessTools-v2-пример-отчёта.xlsx", xlOpenXMLWorkbook
    book.ExportAsFixedFormat xlTypePDF, folder & "\report-preview.pdf"
    book.Close False
    AT_ExportReportSample = "SAVED"
End Function
Public Function AT_VerifyTableReportReference(ByVal workbookName As String) As String
    Dim reader As New ATNetworkRequestReader, checks As New ATNetworkLocalChecks, writer As New ATExcelReportWriter
    Dim report As ATCheckReport, book As Workbook, ws As Worksheet
    Set ws = Application.Workbooks(workbookName).Worksheets(1)
    Set report = checks.Run(reader.Read(ws, "REFERENCE", "2026-10-16", "UTC"), Nothing)
    Set book = writer.Create(report, ws)
    AT_VerifyTableReportReference = "STATE=" & CStr(report.State) & "|CHECKS=" & CStr(report.Checks.Count) & "|ERRORS=" & CStr(report.ErrorCount) & "|WARNINGS=" & CStr(report.WarningCount) & "|LINKS=" & CStr(book.Worksheets("Отчёт").Hyperlinks.Count)
    If book.Worksheets("Отчёт").Hyperlinks(1).Address <> ws.Parent.FullName Then Err.Raise vbObjectError + 2999, , "REFERENCE_LINK_FAILED"
    book.Close False
End Function
