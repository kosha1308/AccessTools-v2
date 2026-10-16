Attribute VB_Name = "modATNetworkCommandTests"
Option Explicit
Public Sub AT_TestNetworkCommand()
    Dim fs As Object, parent As String, first As String, second As String, book As Workbook, other As Workbook, copy As Workbook
    Dim saver As New ATNetworkOutputSaver, caught As Boolean, previousEvents As Boolean, message As String
    Dim result As ATNetworkRunResult, rulePath As String, supersetPath As String, ws As Worksheet
    Dim runId As String, runDate As String, timezoneId As String
    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject")
    parent = fs.BuildPath(fs.GetSpecialFolder(2), "AccessTools-command-" & fs.GetTempName): fs.CreateFolder parent
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set other = Application.Workbooks.Add(xlWBATWorksheet)
    book.Worksheets(1).Range("A1").Value2 = "Отчёт": other.Worksheets(1).Range("A1").Value2 = "Сводка"
    previousEvents = Application.EnableEvents: Application.EnableEvents = True
    first = saver.Save(book, other, parent)
    AT_Assert Len(first) > 0, "output saver returns new result folder"
    If Len(first) = 0 Then GoTo Cleanup
    AT_Assert fs.FileExists(first & "\Отчёт-проверок.xlsx") And fs.FileExists(first & "\Сводка-координатора.xlsx"), "both result copies exist"
    AT_Assert Application.EnableEvents And Len(book.Path) = 0 And Len(other.Path) = 0, "copy export preserves events and original output book identity"
    Set copy = Application.Workbooks.Open(first & "\Сводка-координатора.xlsx", 0, True)
    AT_Assert copy.FileFormat = 51 And copy.Worksheets(1).Range("A1").Value2 = "Сводка", "exported XLSX is readable and contains actual output"
    copy.Close False: Set copy = Nothing
    fs.CreateTextFile(first & "\user-marker.txt").Close
    Application.EnableEvents = False
    second = saver.Save(book, other, parent)
    AT_Assert StrComp(first, second, vbTextCompare) <> 0 And fs.FileExists(first & "\user-marker.txt"), "repeated export never overwrites previous results"
    AT_Assert Not Application.EnableEvents, "export preserves initially disabled events"
    Application.EnableEvents = True
    On Error Resume Next
    saver.Save book, book, parent: caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    AT_Assert caught, "same output book rejected"
    caught = False
    On Error Resume Next
    saver.Save book, other, parent & "\missing": caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    AT_Assert caught And Not fs.FolderExists(parent & "\missing"), "missing output parent never created implicitly"
    AT_Assert fs.GetFolder(parent).SubFolders.Count = 2, "rejected save leaves no partial result directories"
    AT_CommandReset
    caught = False
    On Error Resume Next
    AT_CommandSave parent: caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    AT_Assert caught, "command cannot save before a completed run"
    Set result = AT_CommandExecute(book.Worksheets(1), "missing.xlsx", "missing2.xlsx", "UI_BAD", "2026-10-17", "Europe/Moscow")
    AT_Assert Not result Is Nothing, "command returns current failed run for user diagnostics"
    If Not result Is Nothing Then
        AT_Assert result.State = atEvalError And Not result.IsComplete, "command preserves technical failure state"
        AT_TestCloseRun result: Set result = Nothing
    End If
    caught = False
    On Error Resume Next
    AT_CommandSave parent: caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    AT_Assert caught, "failed command cannot save an earlier result"
    runId = AT_CommandTimestamp(runDate, timezoneId)
    AT_Assert Len(runId) > 0 And runDate = Format$(Date, "yyyy-mm-dd") And timezoneId Like "UTC[+-]##:##", "command captures local date with explicit Windows UTC offset"
    rulePath = AT_TestRunRulebook(parent): supersetPath = AT_TestRunSuperset(parent, "Исходный сервис")
    Set ws = book.Worksheets(1): ws.Cells.Clear: ws.Name = "Лист1": AT_TestNetworkReferenceFixture ws
    ws.Range("A5").Value2 = "Тестовый Пользователь Проверочный": ws.Range("D5").Value2 = "Отдел": ws.Range("E5").Value2 = "Должность"
    ws.Range("F5").Value2 = "user@example.test": ws.Range("C7").Value2 = "Предоставить доступ": ws.Range("C8").Value2 = "Да"
    ws.Range("C11").Value2 = "Рабочая задача": ws.Range("C12").Value2 = "Бессрочно"
    ws.Range("A15:D15").Value2 = Array("192.0.2.1", "198.51.100.1", "TCP", "443"): ws.Range("E15").Value2 = "Причина"
    ws.Range("B17").Value2 = "192.0.2.1": ws.Range("C17").Value2 = "web.test": ws.Range("B18").Value2 = "198.51.100.1"
    Set result = AT_CommandExecute(ws, rulePath, supersetPath, runId, runDate, timezoneId)
    AT_Assert result.IsComplete And Not result.AllChecksPassed, "command completes pipeline without claiming all business checks passed"
    second = AT_CommandSave(parent)
    AT_Assert fs.FileExists(second & "\Отчёт-проверок.xlsx") And fs.FileExists(second & "\Сводка-координатора.xlsx"), "command saves both actual pipeline results"
    AT_CommandReset
    caught = False
    On Error Resume Next
    AT_CommandSave parent: caught = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    AT_Assert caught And Not result.CoordinatorBook Is Nothing, "cancel/reset removes save target but preserves open user results"
    AT_TestCloseRun result: Set result = Nothing
Cleanup:
    On Error Resume Next
    If Not copy Is Nothing Then copy.Close False
    If Not book Is Nothing Then book.Close False
    If Not other Is Nothing Then other.Close False
    AT_TestCloseRun result
    AT_CommandReset
    Application.EnableEvents = previousEvents
    If Len(parent) > 0 Then
        If StrComp(fs.GetParentFolderName(parent), CStr(fs.GetSpecialFolder(2)), vbTextCompare) = 0 And Left$(fs.GetFileName(parent), 20) = "AccessTools-command-" Then fs.DeleteFolder parent, True
    End If
    On Error GoTo 0
    Exit Sub
Failed:
    message = Err.Description: AT_Assert False, "output saver unexpected error: " & message: Resume Cleanup
End Sub
