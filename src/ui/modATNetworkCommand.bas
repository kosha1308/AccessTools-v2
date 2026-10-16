Attribute VB_Name = "modATNetworkCommand"
Option Explicit
Private mResult As ATNetworkRunResult, mBusy As Boolean
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
    Set mResult = Nothing
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
    AT_CommandSave = saver.Save(mResult.ReportBook, mResult.CoordinatorBook, parentFolder)
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
    rulePath = ChooseSource("Выберите книгу правил NETWORK_ACCESS (XLSX)")
    If Len(rulePath) = 0 Then GoTo Cleanup
    supersetPath = ChooseSource("Выберите выгрузку Superset (XLSX)")
    If Len(supersetPath) = 0 Then GoTo Cleanup
    Application.StatusBar = "AccessTools: проверка заявки, правил и Superset..."
    Application.ScreenUpdating = False
    runId = AT_CommandTimestamp(runDate, timezoneId)
    Set result = AT_CommandExecute(request, rulePath, supersetPath, runId, runDate, timezoneId)
    Application.ScreenUpdating = previousScreen: Application.StatusBar = previousStatus
    If Not result.IsComplete Then
        message = "Запуск не завершён. Этап: " & result.Phase & "." & vbCrLf & "Код: " & result.ErrorCode
        If Not result.ReportBook Is Nothing Then result.ReportBook.Activate
        MsgBox message & vbCrLf & "Сводка не сформирована. Подробности — в отчёте, если он создан.", vbExclamation, "AccessTools": GoTo Cleanup
    End If
    result.CoordinatorBook.Activate
    message = "Обработка завершена. Открыты отчёт проверок и сводка координатора."
    If result.AllChecksPassed Then
        message = message & vbCrLf & "Все 16 проверок пройдены."
    Else
        message = message & vbCrLf & "Есть замечания или неполные проверки. Изучите обе книги."
    End If
    If result.RequestUnsaved Then message = message & vbCrLf & "Использованы текущие несохранённые значения заявки."
    If MsgBox(message & vbCrLf & vbCrLf & "Сохранить отдельные копии результатов?", vbYesNo + vbQuestion, "AccessTools") = vbYes Then SaveCurrent
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

Public Sub СохранитьРезультатыСвязанностей()
    If mBusy Then Exit Sub
    mBusy = True
    SaveCurrent
    mBusy = False
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
    MsgBox "Обе книги сохранены в новую папку:" & vbCrLf & output, vbInformation, "AccessTools"
    Exit Sub
NoResult:
    MsgBox "Сначала выполните команду «ПроверитьСвязанности». Последний запуск должен завершиться формированием обеих книг.", vbExclamation, "AccessTools"
    Exit Sub
Failed:
    MsgBox "Сохранение не завершено. Открытые результаты доступны для просмотра и повторного сохранения." & vbCrLf & "Код: " & AT_StableErrorCode(Err.Description, "OUTPUT_SAVE_FAILED"), vbExclamation, "AccessTools"
End Sub
