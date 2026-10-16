Attribute VB_Name = "modATRunControllerTests"
Option Explicit
Public Sub AT_TestRunController()
    Dim fs As Object, folder As String, rulePath As String, supersetPath As String, sourceBook As Workbook, ws As Worksheet
    Dim controller As New ATNetworkRunController, result As ATNetworkRunResult, count As Long, number As Long, description As String
    Dim openBook As Workbook, owners As New Collection, caught As Boolean, code As String, title As String
    Dim stamp As ATRunSourceStamp, stream As Object, previousEvents As Boolean, previousSecurity As Long, stateArmed As Boolean
    Dim macroPath As String, protectedPath As String
    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = fs.BuildPath(fs.GetSpecialFolder(2), "AccessTools-run-" & fs.GetTempName): fs.CreateFolder folder
    rulePath = AT_TestRunRulebook(folder): supersetPath = AT_TestRunSuperset(folder, "Исходный сервис")
    Set sourceBook = Application.Workbooks.Add(xlWBATWorksheet): Set ws = sourceBook.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    ws.Range("A5").Value2 = "Тестовый Пользователь Проверочный": ws.Range("D5").Value2 = "Отдел": ws.Range("E5").Value2 = "Должность"
    ws.Range("F5").Value2 = "user@example.test": ws.Range("C7").Value2 = "Предоставить доступ": ws.Range("C8").Value2 = "Да"
    ws.Range("C11").Value2 = "Рабочая задача": ws.Range("C12").Value2 = "Бессрочно"
    ws.Range("A15:D15").Value2 = Array("192.0.2.1", "198.51.100.1", "TCP", "443"): ws.Range("E15").Value2 = "Причина"
    ws.Range("B17").Value2 = "192.0.2.1": ws.Range("C17").Value2 = "web.test": ws.Range("B18").Value2 = "198.51.100.1"
    count = Application.Workbooks.Count
    Set result = controller.Run(ws, rulePath, supersetPath, "RUN1", "2026-10-17", "Europe/Moscow")
    AT_Assert Not result Is Nothing, "network controller returns atomic run result"
    If result Is Nothing Then GoTo Cleanup
    AT_Assert result.State = atEvalOK And result.IsComplete, "path-based controller completes both outputs"
    If Not result.IsComplete Then GoTo Cleanup
    AT_Assert result.PolicyReport.AllReport.Checks.Count = 16 And Not result.AllChecksPassed, "execution completion does not claim missing IP checks passed"
    AT_Assert result.RequestUnsaved And Not sourceBook.Saved, "current unsaved request is captured without saving user workbook"
    AT_Assert Application.Workbooks.Count = count + 2, "controller closes own sourcebooks and retains exactly two output books"
    AT_Assert result.CoordinatorBook.Worksheets(1).Cells(6, 5).Value2 = "Сервис из правил", "controller composes Superset rule service in summary"
    AT_Assert InStr(result.CoordinatorBook.Worksheets(1).Cells(6, 6).Value2, "Техническая инструкция") > 0 And InStr(result.CoordinatorBook.Worksheets(1).Cells(6, 6).Value2, "Проверить владельца") > 0, "one run includes both instruction layers"
    AT_Assert Len(result.RulebookFingerprint) > 0 And Len(result.SupersetFingerprint) > 0, "actual source fingerprints recorded"
    AT_TestCloseRun result: Set result = Nothing
    owners.Add sourceBook
    On Error Resume Next
    AT_RunAssertNew sourceBook, owners: caught = Err.Number <> 0: code = Err.Description: Err.Clear
    On Error GoTo Failed
    AT_Assert caught And code = "SOURCE_ALREADY_OPEN", "ownership guard rejects workbook objects already owned by user even through aliases"
    supersetPath = AT_TestRunSuperset(folder, "Новый сервис после замены")
    Set result = controller.Run(ws, rulePath, supersetPath, "RUN2", "2026-10-17", "Europe/Moscow")
    AT_Assert result.IsComplete And result.CoordinatorBook.Worksheets(1).Cells(6, 8).Value2 = "Новый сервис после замены", "same filename reread on next run without stale source cache"
    AT_Assert Len(result.CoordinatorBook.Worksheets(1).Cells(6, 5).Value2) = 0, "previous compiled service match not reused after source replacement"
    AT_TestCloseRun result: Set result = Nothing
    Set openBook = Application.Workbooks.Open(supersetPath, 0, False)
    openBook.Worksheets(1).Range("G2").Value2 = "Несохранённое значение"
    Set result = controller.Run(ws, rulePath, fs.BuildPath(fs.GetParentFolderName(fs.GetFile(supersetPath).ShortPath), fs.GetFile(supersetPath).Name), "OPEN", "2026-10-17", "Europe/Moscow")
    AT_Assert result.State = atEvalError And result.ErrorCode = "SOURCE_ALREADY_OPEN" And Not result.IsComplete, "open unsaved source rejects path run without using old disk data"
    AT_Assert Not openBook.Saved And openBook.Worksheets(1).Range("G2").Value2 = "Несохранённое значение", "user source workbook remains open unsaved and unchanged"
    AT_TestCloseRun result: Set result = Nothing
    openBook.Worksheets(1).Range("B1").Value2 = "Неверный заголовок": openBook.Save: openBook.Close False: Set openBook = Nothing
    Set result = controller.Run(ws, rulePath, supersetPath, "BAD_HEADER", "2026-10-17", "Europe/Moscow")
    AT_Assert result.State = atEvalError And result.CoordinatorBook Is Nothing And result.PolicyReport Is Nothing, "bad replacement source never returns stale prior coordinator or policy"
    AT_Assert result.ReportBook.Worksheets(1).Range("A3").Value2 Like "Запуск не завершён*", "source failure visible in separate blocked report"
    AT_TestCloseRun result: Set result = Nothing
    title = ws.Range("A1").Value2: ws.Range("A1").Value2 = "Неверный шаблон"
    Set result = controller.Run(ws, "missing.xlsx", "missing2.xlsx", "BAD_GATE", "2026-10-17", "Europe/Moscow")
    AT_Assert result.State = atEvalError And result.Phase = "чтение заявки", "request gate runs before touching source paths"
    AT_TestCloseRun result: Set result = Nothing: ws.Range("A1").Value2 = title
    supersetPath = AT_TestRunSuperset(folder, "Исходный сервис")
    previousEvents = Application.EnableEvents: previousSecurity = Application.AutomationSecurity: stateArmed = True
    Application.EnableEvents = False: Application.AutomationSecurity = 2
    Set result = controller.Run(ws, rulePath, supersetPath, "RECOVERY", "2026-10-17", "Europe/Moscow")
    AT_Assert result.IsComplete And Not Application.EnableEvents And Application.AutomationSecurity = 2, "repeat after error restores original disabled events and macro setting"
    Application.EnableEvents = previousEvents: Application.AutomationSecurity = previousSecurity: stateArmed = False
    AT_TestCloseRun result: Set result = Nothing
    Set result = controller.Run(ws, rulePath, rulePath, "COLLISION", "2026-10-17", "Europe/Moscow")
    AT_Assert result.State = atEvalError And result.ErrorCode = "SOURCE_PATH_COLLISION", "same rulebook and Superset file rejected"
    AT_TestCloseRun result: Set result = Nothing
    protectedPath = folder & "\protected.xlsx"
    Set openBook = Application.Workbooks.Add(xlWBATWorksheet)
    openBook.SaveAs Filename:=protectedPath, FileFormat:=51, Password:="fixture-only": openBook.Close False: Set openBook = Nothing
    Set result = controller.Run(ws, rulePath, protectedPath, "PASSWORD", "2026-10-17", "Europe/Moscow")
    AT_Assert result.State = atEvalError And result.CoordinatorBook Is Nothing And Application.Workbooks.Count = count + 1, "password-protected source fails without leaking owned input book or prompting for password"
    AT_TestCloseRun result: Set result = Nothing
    macroPath = folder & "\unsupported.xlsm": fs.CopyFile supersetPath, macroPath
    Set stamp = New ATRunSourceStamp
    On Error Resume Next
    stamp.Initialize macroPath: caught = Err.Number <> 0: code = Err.Description: Err.Clear
    On Error GoTo Failed
    AT_Assert caught And code = "SOURCE_FORMAT_UNSUPPORTED", "initial controller accepts supplied macro-free XLSX sources only"
    Set stamp = New ATRunSourceStamp: stamp.Initialize supersetPath: stamp.VerifyCurrent
    Set stream = fs.OpenTextFile(supersetPath, 8): stream.Write "x": stream.Close: Set stream = Nothing
    On Error Resume Next
    stamp.VerifyCurrent: caught = Err.Number <> 0: code = Err.Description: Err.Clear
    On Error GoTo Failed
    AT_Assert caught And code = "SOURCE_CHANGED_DURING_RUN", "size or timestamp replacement invalidates captured source receipt"
Cleanup:
    AT_TestCloseRun result
    If Not sourceBook Is Nothing Then sourceBook.Close False
    If Not openBook Is Nothing Then openBook.Close False
    If Len(folder) > 0 Then
        If StrComp(fs.GetParentFolderName(folder), CStr(fs.GetSpecialFolder(2)), vbTextCompare) <> 0 Or Left$(fs.GetFileName(folder), 16) <> "AccessTools-run-" Then Err.Raise vbObjectError + 3180, , "TEST_FOLDER_INVALID"
        If fs.FileExists(macroPath) Then fs.DeleteFile macroPath, True
        If fs.FileExists(protectedPath) Then fs.DeleteFile protectedPath, True
        If fs.FileExists(rulePath) Then fs.DeleteFile rulePath, True
        If fs.FileExists(supersetPath) Then fs.DeleteFile supersetPath, True
        If fs.FolderExists(folder) Then fs.DeleteFolder folder, False
    End If
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If stateArmed Then Application.EnableEvents = previousEvents: Application.AutomationSecurity = previousSecurity
    If Not stream Is Nothing Then stream.Close
    If Not openBook Is Nothing Then openBook.Close False
    AT_TestCloseRun result
    If Not sourceBook Is Nothing Then sourceBook.Close False
    If Len(folder) > 0 Then
        If StrComp(fs.GetParentFolderName(folder), CStr(fs.GetSpecialFolder(2)), vbTextCompare) = 0 And Left$(fs.GetFileName(folder), 16) = "AccessTools-run-" Then
        If fs.FileExists(macroPath) Then fs.DeleteFile macroPath, True
        If fs.FileExists(protectedPath) Then fs.DeleteFile protectedPath, True
        If fs.FileExists(rulePath) Then fs.DeleteFile rulePath, True
        If fs.FileExists(supersetPath) Then fs.DeleteFile supersetPath, True
        If fs.FolderExists(folder) Then fs.DeleteFolder folder, False
        End If
    End If
    On Error GoTo 0
    Err.Raise number, "AT_TestRunController", description
End Sub
Public Sub AT_TestCloseRun(ByVal result As ATNetworkRunResult)
    If result Is Nothing Then Exit Sub
    If Not result.CoordinatorBook Is Nothing Then result.CoordinatorBook.Close False
    If Not result.ReportBook Is Nothing Then result.ReportBook.Close False
End Sub
Public Function AT_TestRunRulebook(ByVal folder As String) As String
    Dim book As Workbook, ws As Worksheet, names As Variant, headers As Variant, i As Long
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    names = Array("Инициаторы", "Правила и исключения", "Протоколы", "Проблемы", "Общий текст", "Маршрутизация", "Ограничения IP инициаторов")
    headers = Array("Активно|Домены / почты|Требуемое согласование|Уровень сообщения|Комментарий / инструкция|Служебный режим", "Активно|Название правила|IP|Область IP|Hostname/FQDN|Тип совпадения hostname|Порт назначения|Протокол назначения|Сервис — условие|Сервис — результат|Группа назначения|Требуемое согласование|Комментарий / инструкция|Уровень сообщения|Инструкция", "Активно|Протокол|Комментарий", "Поле заявки|Что проверяется|Штатное сообщение|Сообщение пользователю|Сообщение координатору|Уровень сообщения|Ошибка пользователя|Подсвечивать в заявке|Активно|Ключ", "Ключ|Значение", "Проверка|Поле / результат|Что проверяется|В «Все проблемы»|Подсветка / комментарий в заявке|Результат / комментарий в Superset|Активно|Ключ", "Активно|Email / домен|Область IP|Разрешённые IP|Запрещённые IP (НЕ)|Комментарий / инструкция|Учетная запись для портов БД")
    For i = 0 To 6
        If i = 0 Then Set ws = book.Worksheets(1) Else Set ws = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
        ws.Name = CStr(names(i)): ws.Cells(1, 1).Resize(1, UBound(Split(CStr(headers(i)), "|")) + 1).Value2 = Split(CStr(headers(i)), "|")
    Next i
    book.Worksheets("Протоколы").Range("A2:B2").Value2 = Array("Да", "TCP")
    book.Worksheets("Инициаторы").Range("A2:F2").Value2 = Array("Да", "example.test", "", "Информация", "", "")
    book.Worksheets("Правила и исключения").Range("A2:O2").Value2 = AT_TestSupersetRule("Сервисное", "192.0.2.1", "web.test", "Исходный сервис", "Сервис из правил", "Проверить владельца")
    book.Worksheets("Правила и исключения").Range("A3:O3").Value2 = Array("Да", "Техническое", "198.51.100.1", "Назначение", "", "", "443", "TCP", "", "Технический сервис", "", "", "", "Информация", "Техническая инструкция")
    AT_TestRunRulebook = folder & "\rules.xlsx": book.SaveAs AT_TestRunRulebook, 51: book.Close False
End Function
Public Function AT_TestRunSuperset(ByVal folder As String, ByVal service As String) As String
    Dim book As Workbook, ws As Worksheet
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1)
    ws.Range("B1:I1").Value2 = Array("FQDN", "SOURCE:ТЕЛЕГРАФ", "SOURCE: ЦХД", "SOURCE: AWX", "ПОЛНОЕ ИМЯ ИС", "СЕРВИС", "СКУФ ГРУППА ТП", "КОМАНДА ПОДДЕРЖКИ")
    ws.Range("B2:I2").Value2 = Array("web.test", "192.0.2.1", "", "", "Тестовая ИС", service, "Группа", "Поддержка")
    AT_TestRunSuperset = folder & "\superset.xlsx": book.SaveAs AT_TestRunSuperset, 51: book.Close False
End Function
