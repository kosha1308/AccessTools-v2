Attribute VB_Name = "modATNetworkCommand"
Option Explicit
Private mResult As ATNetworkRunResult, mBusy As Boolean
Private mFeedback As ATNetworkUserFeedback, mRequest As Worksheet, mFeedbackWarning As String
Private mRulePath As String, mSupersetPath As String, mSettingsWarning As String
Private mWindow As Object, mWorkingBook As Workbook
Private mOwnHidden As Boolean
Private mUserRun As Boolean
Private Type ATCommandSystemTime
    year As Integer
    month As Integer
    dayOfWeek As Integer
    day As Integer
    hour As Integer
    minute As Integer
    second As Integer
    milliseconds As Integer
End Type
Private Type ATCommandTimeZone
    bias As Long
    standardName(0 To 31) As Integer
    standardDate As ATCommandSystemTime
    standardBias As Long
    daylightName(0 To 31) As Integer
    daylightDate As ATCommandSystemTime
    daylightBias As Long
End Type
#If VBA7 Then
Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (ByRef information As ATCommandTimeZone) As Long
#Else
Private Declare Function GetTimeZoneInformation Lib "kernel32" (ByRef information As ATCommandTimeZone) As Long
#End If
Public Sub AT_CommandReset()
    On Error Resume Next
    If Not mWindow Is Nothing Then Unload mWindow
    Set mWindow = Nothing
    If mOwnHidden And Not mResult Is Nothing Then
        CloseHidden mResult.ReportBook
        CloseHidden mResult.CoordinatorBook
    End If
    mOwnHidden = False
    mUserRun = False
    Set mWorkingBook = Nothing
    On Error GoTo 0
    Set mResult = Nothing
    Set mFeedback = Nothing: Set mRequest = Nothing: mFeedbackWarning = ""
End Sub
Private Sub CloseHidden(ByVal book As Workbook)
    On Error Resume Next
    If Not book Is Nothing Then
        If Not book.Windows(1).Visible Then book.Close False
    End If
    On Error GoTo 0
End Sub
Public Function AT_CommandExecute(ByVal request As Worksheet, ByVal rulebookPath As String, ByVal supersetPath As String, ByVal runId As String, ByVal runDate As String, ByVal timezoneId As String) As ATNetworkRunResult
    Dim controller As New ATNetworkRunController
    AT_CommandReset
    Set mResult = controller.Run(request, rulebookPath, supersetPath, runId, runDate, timezoneId)
    Set AT_CommandExecute = mResult
End Function
Public Function AT_CommandSave(ByVal parentFolder As String) As String
    Dim saver As New ATNetworkOutputSaver
    If mResult Is Nothing Then Err.Raise vbObjectError + 3190, , "COMMAND_NO_COMPLETE_RUN"
    If Not mResult.IsComplete Then Err.Raise vbObjectError + 3190, , "COMMAND_NO_COMPLETE_RUN"
    If mUserRun And mWorkingBook Is Nothing Then Err.Raise vbObjectError + 3240, , "WORKSPACE_NOT_READY"
    If Not mWorkingBook Is Nothing Then
        AT_CommandSave = saver.Save(Nothing, mWorkingBook, parentFolder, False)
    Else
        AT_CommandSave = saver.Save(mResult.ReportBook, mResult.CoordinatorBook, parentFolder)
    End If
End Function
Public Function AT_CommandTimestamp(ByRef runDate As String, ByRef timezoneId As String) As String
    Dim information As ATCommandTimeZone, state As Long, offset As Long, captured As Date, sign As String
    captured = Now: state = GetTimeZoneInformation(information)
    If state = -1 Then Err.Raise vbObjectError + 3190, , "COMMAND_TIMEZONE_UNAVAILABLE"
    offset = information.bias
    If state = 1 Then offset = offset + information.standardBias
    If state = 2 Then offset = offset + information.daylightBias
    offset = -offset: sign = "+": If offset < 0 Then sign = "-"
    timezoneId = "UTC" & sign & Format$(Abs(offset) \ 60, "00") & ":" & Format$(Abs(offset) Mod 60, "00")
    runDate = Format$(captured, "yyyy-mm-dd")
    AT_CommandTimestamp = Format$(captured, "yyyymmdd_hhnnss") & "_" & Format$(Timer * 100, "00000000")
End Function

' Public Russian macros are visible in Excel's macro list (Alt+F8).
Public Sub ПроверитьСвязанности()
    Dim request As Worksheet, rulePath As String, supersetPath As String, result As ATNetworkRunResult
    Dim profile As ATTemplateProfile, decision As ATTemplateGateResult, previousStatus As Variant
    Dim previousScreen As Boolean, runId As String, runDate As String, timezoneId As String, message As String
    If mBusy Then Exit Sub
    If Not AT_PackageAllowCommand() Then Exit Sub
    mBusy = True: previousStatus = Application.StatusBar: previousScreen = Application.ScreenUpdating
    On Error GoTo Failed
    AT_CommandReset
    If Not TypeOf Application.ActiveSheet Is Worksheet Then
        MsgBox "Откройте лист заявки на сетевую связность и повторите команду.", vbExclamation, "AccessTools": GoTo Cleanup
    End If
    Set request = Application.ActiveSheet
    Set profile = AT_NetworkProfileV3(): Set decision = profile.Verify(request)
    If Not decision.Allowed Then
        MsgBox decision.Message & vbCrLf & "Откройте лист актуальной заявки на сетевую связность.", vbExclamation, "AccessTools": GoTo Cleanup
    End If
    If Not ResolveSources(False, rulePath, supersetPath) Then GoTo Cleanup
    Application.StatusBar = "AccessTools: проверка заявки, правил и Superset..."
    Application.ScreenUpdating = False
    runId = AT_CommandTimestamp(runDate, timezoneId)
    Dim oldNotes As New ATNetworkUserFeedback
    oldNotes.ClearNotes request
    Set result = AT_CommandExecute(request, rulePath, supersetPath, runId, runDate, timezoneId)
    mUserRun = True
    If Not result.IsComplete Then
        message = "Запуск не завершён. Этап: " & result.Phase & "." & vbCrLf & "Код: " & result.ErrorCode
        If oldNotes.NoteFailures > 0 Then message = message & vbCrLf & "Старые примечания не удалось обновить; они могут относиться к предыдущей проверке."
        request.Activate
        If Not result.ReportBook Is Nothing Then result.ReportBook.Windows(1).Visible = False: mOwnHidden = True
        MsgBox message & vbCrLf & "Сводка не сформирована. Подробности — по кнопке «Диагностика».", vbExclamation, "AccessTools": GoTo Cleanup
    End If
    AT_CommandPresent request, result, supersetPath, runId
    Application.ScreenUpdating = previousScreen: Application.StatusBar = previousStatus
    ShowCurrentResults
Cleanup:
    On Error Resume Next
    Application.ScreenUpdating = previousScreen: Application.StatusBar = previousStatus: mBusy = False
    On Error GoTo 0
    Exit Sub
Failed:
    message = Err.Description
    MsgBox "Не удалось выполнить команду. " & message & vbCrLf & "Открытые книги не закрывались.", vbExclamation, "AccessTools"
    Resume Cleanup
End Sub
Public Function AT_CommandPresent(ByVal request As Worksheet, ByVal result As ATNetworkRunResult, ByVal supersetPath As String, ByVal runId As String) As Workbook
    If mResult Is Nothing Or result Is Nothing Then Err.Raise vbObjectError + 3240, , "WORKSPACE_NO_CURRENT_RUN"
    If Not result Is mResult Then Err.Raise vbObjectError + 3240, , "WORKSPACE_STALE_RUN"
    mUserRun = True: mOwnHidden = True: Set mRequest = request
    request.Activate
    result.ReportBook.Windows(1).Visible = False
    result.CoordinatorBook.Windows(1).Visible = False
    Set mWorkingBook = Nothing
    Set mWorkingBook = AT_CommandWorkingCopy(result, supersetPath)
    AT_CommandFeedback request, result, runId
    request.Activate
    Set AT_CommandPresent = mWorkingBook
End Function

Public Sub СохранитьРезультатыСвязанностей()
    If mBusy Then Exit Sub
    If Not AT_PackageAllowCommand() Then Exit Sub
    mBusy = True
    SaveCurrent
    mBusy = False
End Sub
Public Sub AT_RibbonCheck(ByVal control As Object)
    ПроверитьСвязанности
End Sub
Public Sub AT_RibbonSave(ByVal control As Object)
    СохранитьРезультатыСвязанностей
End Sub
Public Sub AT_RibbonResults(ByVal control As Object)
    РезультатыПроверкиСвязанностей
End Sub
Public Sub AT_RibbonSources(ByVal control As Object)
    ВыбратьИсточникиСвязанностей
End Sub
Public Sub ВыбратьИсточникиСвязанностей()
    Dim rulePath As String, supersetPath As String
    If mBusy Then Exit Sub
    If Not AT_PackageAllowCommand() Then Exit Sub
    On Error GoTo Failed
    mBusy = True
    If ResolveSources(True, rulePath, supersetPath) Then
        AT_CommandReset
        MsgBox "Источники выбраны. Следующие проверки будут использовать их без повторного выбора." & vbCrLf & mSettingsWarning, vbInformation, "AccessTools"
    End If
    mBusy = False: Exit Sub
Failed:
    mBusy = False: MsgBox "Не удалось выбрать источники. " & Err.Description, vbExclamation, "AccessTools"
End Sub
Private Function ResolveSources(ByVal force As Boolean, ByRef rulePath As String, ByRef supersetPath As String) As Boolean
    Dim settings As New ATNetworkSourceSettings, fs As Object, changed As Boolean
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Len(mRulePath) = 0 Or Len(mSupersetPath) = 0 Then
        If settings.Load Then mRulePath = settings.RulebookPath: mSupersetPath = settings.SupersetPath
    End If
    rulePath = mRulePath: supersetPath = mSupersetPath
    If force Or Not fs.FileExists(rulePath) Then
        rulePath = ChooseSource("Выберите книгу правил NETWORK_ACCESS (XLSX)"): changed = True
        If Len(rulePath) = 0 Then Exit Function
    End If
    If force Or Not fs.FileExists(supersetPath) Then
        supersetPath = ChooseSource("Выберите выгрузку Superset (XLSX)"): changed = True
        If Len(supersetPath) = 0 Then Exit Function
    End If
    If StrComp(fs.GetAbsolutePathName(rulePath), fs.GetAbsolutePathName(supersetPath), vbTextCompare) = 0 Then
        MsgBox "Выбран один и тот же файл. Для правил и Superset нужны разные XLSX.", vbExclamation, "AccessTools": Exit Function
    End If
    mRulePath = rulePath: mSupersetPath = supersetPath
    If changed Then
        mSettingsWarning = ""
        On Error Resume Next
        Err.Clear
        settings.Save rulePath, supersetPath
        If Err.Number <> 0 Then mSettingsWarning = "Пути запомнены только до закрытия Excel: не удалось сохранить настройки."
        Err.Clear: On Error GoTo 0
    End If
    ResolveSources = True
End Function
Public Sub AT_CommandFeedback(ByVal request As Worksheet, ByVal result As ATNetworkRunResult, ByVal runId As String)
    On Error GoTo Failed
    Set mFeedback = Nothing: Set mRequest = request: mFeedbackWarning = ""
    If result Is Nothing Then Exit Sub
    If Not result.IsComplete Then Exit Sub
    Set mFeedback = New ATNetworkUserFeedback
    mFeedback.Prepare result.PolicyReport
    mFeedback.ApplyNotes request, runId
    Application.DisplayCommentIndicator = xlCommentIndicatorOnly
    result.ReportBook.Worksheets(1).Range("A2").Value2 = "Отдельный отчёт: значения заявки не менялись. Примечания обновлены в открытой книге без автоматического сохранения."
    If mFeedback.NoteFailures > 0 Then mFeedbackWarning = "Часть примечаний не добавлена или не обновлена. Проверьте защиту листа и отчёт; данные проверки доступны в отдельных книгах."
    Exit Sub
Failed:
    mFeedbackWarning = "Не удалось полностью подготовить список ошибок или примечания. Используйте отчёт проверок. Код: " & AT_StableErrorCode(Err.Description, "FEEDBACK_FAILED")
    Set mFeedback = Nothing
End Sub
Public Sub РезультатыПроверкиСвязанностей()
    If mBusy Then Exit Sub
    If Not AT_PackageAllowCommand() Then Exit Sub
    mBusy = True: ShowCurrentResults: mBusy = False
End Sub
Private Sub ShowCurrentResults()
    Dim message As String
    On Error GoTo Failed
    If mResult Is Nothing Then GoTo NoResult
    If Not mResult.IsComplete Then GoTo NoResult
    message = mFeedbackWarning
    If Len(mSettingsWarning) > 0 Then message = message & vbCrLf & mSettingsWarning
    If Not mWindow Is Nothing Then Unload mWindow
    Set mWindow = VBA.UserForms.Add("frmATResults")
    mWindow.LoadData mFeedback, mRequest, message
    mWindow.Show 0
Cleanup:
    Exit Sub
NoResult:
    MsgBox "Сначала выполните проверку связанностей. Готового результата последнего запуска нет.", vbInformation, "AccessTools": GoTo Cleanup
Failed:
    MsgBox "Не удалось открыть окно результатов. Используйте кнопки «Диагностика» или «Сводка и Superset».", vbExclamation, "AccessTools": Resume Cleanup
End Sub
Public Sub AT_ResultCopy()
    On Error GoTo Failed
    If mFeedback Is Nothing Then Exit Sub
    If mFeedback.ErrorCount = 0 Then Exit Sub
    AT_CopyUnicodeClipboard mFeedback.UserErrorsText
    MsgBox "Все ошибки пользователя скопированы. Вставьте текст сочетанием Ctrl+V.", vbInformation, "AccessTools": Exit Sub
Failed:
    MsgBox "Не удалось скопировать ошибки. Буфер может быть занят: повторите из меню «Результаты проверки». Код: " & AT_StableErrorCode(Err.Description, "CLIPBOARD_WRITE_FAILED"), vbExclamation, "AccessTools"
End Sub
Public Sub AT_ResultReport()
    On Error GoTo Failed
    If mResult Is Nothing Then MsgBox "Сначала выполните проверку заявки.", vbInformation, "AccessTools 2.0": Exit Sub
    mResult.ReportBook.Windows(1).Visible = True
    mResult.ReportBook.Activate
    Exit Sub
Failed:
    MsgBox "Книга отчёта уже закрыта. Повторите проверку.", vbInformation, "AccessTools"
End Sub
Public Sub AT_ResultCoordinator()
    On Error GoTo Failed
    If Not mWorkingBook Is Nothing Then
        mWorkingBook.Windows(1).Visible = True: mWorkingBook.Activate
    ElseIf mUserRun Then
        MsgBox "Рабочая копия Superset не сформирована. Повторите проверку. Подробности — по кнопке «Диагностика».", vbExclamation, "AccessTools 2.0"
    ElseIf Not mResult Is Nothing Then
        If mResult.IsComplete Then mResult.CoordinatorBook.Windows(1).Visible = True: mResult.CoordinatorBook.Activate
    End If
    Exit Sub
Failed:
    MsgBox "Книга сводки уже закрыта. Повторите проверку.", vbInformation, "AccessTools"
End Sub
Public Sub AT_ResultRepeat()
    If mBusy Then Exit Sub
    On Error GoTo Failed
    If mRequest Is Nothing Then Exit Sub
    mRequest.Activate
    ПроверитьСвязанности
    Exit Sub
Failed:
    MsgBox "Заявка уже закрыта. Откройте её и запустите проверку.", vbInformation, "AccessTools 2.0"
End Sub
Public Sub AT_ResultClear()
    If mBusy Then Exit Sub
    On Error GoTo Failed
    If mRequest Is Nothing Then Exit Sub
    Dim feedback As New ATNetworkUserFeedback
    If feedback.ClearNotes(mRequest) Then
        MsgBox "Отметки AccessTools 2.0 удалены. Значения заявки не менялись.", vbInformation, "AccessTools 2.0"
    Else
        MsgBox "Часть отметок не удалось удалить. Проверьте защиту листа.", vbExclamation, "AccessTools 2.0"
    End If
    Exit Sub
Failed:
    MsgBox "Не удалось очистить отметки. Возможно, заявка закрыта.", vbInformation, "AccessTools 2.0"
End Sub
Public Sub AT_OpenNetworkRules()
    Dim rulePath As String, supersetPath As String
    If mBusy Or Not AT_PackageAllowCommand() Then Exit Sub
    On Error GoTo Failed
    If Not ResolveSources(False, rulePath, supersetPath) Then Exit Sub
    Application.Workbooks.Open Filename:=rulePath, UpdateLinks:=0, AddToMru:=False
    Exit Sub
Failed:
    MsgBox "Не удалось открыть книгу правил. " & Err.Description, vbExclamation, "AccessTools 2.0"
End Sub
Public Function AT_CommandWorkingCopy(ByVal result As ATNetworkRunResult, ByVal supersetPath As String) As Workbook
    Dim stamp As New ATRunSourceStamp, source As Workbook, writer As New ATCoordinatorWorkbookWriter, book As Workbook
    Dim events As Boolean, security As Long, armed As Boolean, code As String
    On Error GoTo Failed
    If result Is Nothing Then Err.Raise vbObjectError + 3240, , "WORKSPACE_NO_RESULT"
    If Not result.IsComplete Then Err.Raise vbObjectError + 3240, , "WORKSPACE_NO_RESULT"
    stamp.Initialize supersetPath
    If stamp.Fingerprint <> result.SupersetFingerprint Then Err.Raise vbObjectError + 3240, , "WORKSPACE_SOURCE_CHANGED"
    events = Application.EnableEvents: security = Application.AutomationSecurity: armed = True
    Application.EnableEvents = False: Application.AutomationSecurity = 3
    Set source = AT_RunOpenOwned(supersetPath)
    Set book = writer.CreateWorkingCopy(source, result.CoordinatorBook)
    stamp.VerifyCurrent
    source.Close False: Set source = Nothing
    stamp.VerifyCurrent
    Application.EnableEvents = events: Application.AutomationSecurity = security: armed = False
    Set AT_CommandWorkingCopy = book: Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "WORKSPACE_FAILED")
    On Error Resume Next
    If Not source Is Nothing Then source.Close False
    If Not book Is Nothing Then book.Close False
    If armed Then Application.EnableEvents = events: Application.AutomationSecurity = security
    On Error GoTo 0
    Err.Raise vbObjectError + 3240, "AT_CommandWorkingCopy", code
End Function
Public Sub AT_ResultRequest()
    On Error GoTo Failed
    If Not mRequest Is Nothing Then mRequest.Activate
    Exit Sub
Failed:
    MsgBox "Заявка уже закрыта.", vbInformation, "AccessTools"
End Sub
Public Sub AT_ResultSave()
    SaveCurrent
End Sub
Private Function ChooseSource(ByVal title As String) As String
    Dim choice As Variant
    choice = Application.GetOpenFilename("Книги Excel (*.xlsx),*.xlsx", 1, title, , False)
    If VarType(choice) <> vbBoolean Then ChooseSource = CStr(choice)
End Function
Private Sub SaveCurrent()
    Dim dialog As Object, folder As String, output As String
    On Error GoTo Failed
    If mResult Is Nothing Then GoTo NoResult
    If Not mResult.IsComplete Then GoTo NoResult
    Set dialog = Application.FileDialog(4)
    dialog.Title = "Выберите папку: результаты будут сохранены в новой подпапке"
    dialog.AllowMultiSelect = False
    If dialog.Show <> -1 Then Exit Sub
    folder = dialog.SelectedItems(1)
    output = AT_CommandSave(folder)
    MsgBox "Результаты сохранены в новую папку:" & vbCrLf & output, vbInformation, "AccessTools"
    Exit Sub
NoResult:
    MsgBox "Сначала выполните команду «ПроверитьСвязанности». Последний запуск должен завершиться формированием обеих книг.", vbExclamation, "AccessTools"
    Exit Sub
Failed:
    MsgBox "Сохранение не завершено. Открытые результаты доступны для просмотра и повторного сохранения." & vbCrLf & "Код: " & AT_StableErrorCode(Err.Description, "OUTPUT_SAVE_FAILED"), vbExclamation, "AccessTools"
End Sub
