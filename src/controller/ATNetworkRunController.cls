VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkRunController"
Option Explicit
Public Function Run(ByVal request As Worksheet, ByVal rulebookPath As String, ByVal supersetPath As String, ByVal runId As String, ByVal runDate As String, ByVal timezoneId As String) As ATNetworkRunResult
    Dim result As New ATNetworkRunResult, reader As New ATNetworkRequestReader, source As ATNetworkReadResult
    Dim ruleStamp As New ATRunSourceStamp, supersetStamp As New ATRunSourceStamp, ruleBook As Workbook, supersetBook As Workbook
    Dim catalogCompiler As New ATNetworkPolicyCompiler, actorCompiler As New ATInitiatorPolicyCompiler, ruleCompiler As New ATNetworkRulebookCompiler
    Dim catalog As ATNetworkPolicyResult, actor As ATInitiatorPolicyResult, rules As ATNetworkRulebookResult
    Dim indexCompiler As New ATIPSourceCompiler, index As ATIPSourceResult, addressFields As New Collection
    Dim planner As New ATNetworkRegistryPlanner, registry As ATNetworkRegistryResult, evaluator As New ATNetworkSupersetRules, annotations As ATNetworkSupersetResult
    Dim checks As New ATNetworkAccessChecks, policy As ATNetworkPolicyReport, notesRunner As New ATCoordinatorTechnicalNotes, notes As ATCoordinatorNotesResult
    Dim reportWriter As New ATExcelReportWriter, coordinatorWriter As New ATCoordinatorWorkbookWriter, report As ATCheckReport
    Dim reportBook As Workbook, coordinatorBook As Workbook, previousEvents As Boolean, previousSecurity As Long, armed As Boolean
    Dim phase As String, code As String, ruleFingerprint As String, supersetFingerprint As String, unsaved As Boolean
    On Error GoTo Failed
    phase = "чтение заявки"
    If request Is Nothing Then Err.Raise vbObjectError + 3180, , "RUN_REQUEST_REQUIRED"
    unsaved = Not request.Parent.Saved
    Set source = reader.Read(request, runId, runDate, timezoneId)
    If source.State <> atEvalOK Then Err.Raise vbObjectError + 3180, , source.Diagnostics(1).Code
    phase = "подготовка файлов"
    ruleStamp.Initialize rulebookPath: supersetStamp.Initialize supersetPath
    ruleFingerprint = ruleStamp.Fingerprint: supersetFingerprint = supersetStamp.Fingerprint
    If StrComp(ruleStamp.Path, supersetStamp.Path, vbTextCompare) = 0 Then Err.Raise vbObjectError + 3180, , "SOURCE_PATH_COLLISION"
    If StrComp(request.Parent.FullName, ruleStamp.Path, vbTextCompare) = 0 Or StrComp(request.Parent.FullName, supersetStamp.Path, vbTextCompare) = 0 Then Err.Raise vbObjectError + 3180, , "SOURCE_PATH_COLLISION"
    AT_RunRejectOpen ruleStamp.Path: AT_RunRejectOpen supersetStamp.Path
    previousEvents = Application.EnableEvents: previousSecurity = Application.AutomationSecurity
    armed = True: Application.EnableEvents = False: Application.AutomationSecurity = 3
    phase = "открытие источников"
    Set ruleBook = AT_RunOpenOwned(ruleStamp.Path)
    Set supersetBook = AT_RunOpenOwned(supersetStamp.Path)
    If Not ruleBook.ReadOnly Or Not supersetBook.ReadOnly Then Err.Raise vbObjectError + 3180, , "SOURCE_READONLY_REQUIRED"
    phase = "чтение справочников"
    Set catalog = catalogCompiler.Compile(AT_RunReadSource(ruleBook, "Протоколы", "PROTOCOLS"), AT_RunReadSource(ruleBook, "Проблемы", "PROBLEMS"), AT_RunReadSource(ruleBook, "Маршрутизация", "ROUTING"), AT_RunReadSource(ruleBook, "Общий текст", "COMMON_TEXT"))
    If catalog.State <> atEvalOK Then Err.Raise vbObjectError + 3180, , catalog.Diagnostics(1).Code
    Set actor = actorCompiler.Compile(AT_RunReadSource(ruleBook, "Инициаторы", "INITIATORS"), AT_RunReadSource(ruleBook, "Ограничения IP инициаторов", "INITIATOR_IP"), AT_NETWORK_NO_IP_POLICY_ACTION)
    Set rules = ruleCompiler.Compile(AT_RunReadSource(ruleBook, "Правила и исключения", "RULES"), catalog.Policy.Protocols)
    addressFields.Add "SOURCE_TELEGRAPH": addressFields.Add "SOURCE_CXD": addressFields.Add "SOURCE_AWX"
    Set index = indexCompiler.Compile(AT_RunReadSource(supersetBook, supersetBook.Worksheets(1).Name, "SUPERSET"), "SUPERSET", addressFields, supersetFingerprint)
    If index.State <> atEvalOK Then Err.Raise vbObjectError + 3180, , index.Diagnostics(1).Code
    ruleBook.Close False: Set ruleBook = Nothing
    supersetBook.Close False: Set supersetBook = Nothing
    Application.AutomationSecurity = previousSecurity
    ruleStamp.VerifyCurrent: supersetStamp.VerifyCurrent
    phase = "проверки и поиск"
    Set registry = planner.Run(source, index)
    Set annotations = evaluator.Run(registry, rules)
    Set policy = checks.ApplySupersetRules(checks.ApplyRegistry(checks.RunWithRules(source, catalog, actor, rules), registry, catalog), annotations, catalog)
    Set report = policy.AllReport
    If report.State <> atEvalOK Then Err.Raise vbObjectError + 3180, , "RUN_CHECKS_BLOCKED"
    Set notes = notesRunner.Run(source, registry, policy)
    ruleStamp.VerifyCurrent: supersetStamp.VerifyCurrent
    phase = "формирование результатов"
    Set reportBook = reportWriter.Create(report, request)
    Set coordinatorBook = coordinatorWriter.Create(registry, annotations, notes)
    ruleStamp.VerifyCurrent: supersetStamp.VerifyCurrent
    Application.AutomationSecurity = previousSecurity: Application.EnableEvents = previousEvents: armed = False
    result.Initialize atEvalOK, "", "завершено", report, policy, reportBook, coordinatorBook, ruleFingerprint, supersetFingerprint, unsaved
    Set Run = result: Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "RUN_FAILED")
    On Error Resume Next
    If Not ruleBook Is Nothing Then ruleBook.Close False
    If Not supersetBook Is Nothing Then supersetBook.Close False
    If Not coordinatorBook Is Nothing Then coordinatorBook.Close False
    If Not reportBook Is Nothing Then reportBook.Close False
    Set coordinatorBook = Nothing: Set reportBook = Nothing
    If report Is Nothing Then Set report = AT_RunFailure(code, phase)
    Set reportBook = reportWriter.Create(report, request)
    If Not reportBook Is Nothing Then reportBook.Worksheets(1).Cells(3, 1).Value2 = "Запуск не завершён; координаторский результат не сформирован. " & AT_RunFailure(code, phase).Issues(1).Message
    If armed Then Application.AutomationSecurity = previousSecurity: Application.EnableEvents = previousEvents
    On Error GoTo 0
    Set result = New ATNetworkRunResult
    result.Initialize atEvalError, code, phase, report, policy, reportBook, coordinatorBook, ruleFingerprint, supersetFingerprint, unsaved
    Set Run = result
End Function
