VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkDocumentChecks"
Option Explicit
Public Function Run(ByVal source As ATNetworkReadResult) As ATCheckReport
    Dim adapter As New ATNetworkRequestNormalizer, normalized As ATNetworkReadResult, context As ATEvaluationContext
    Dim issues As New Collection, checks As New Collection, report As New ATCheckReport
    Dim entry As Variant, value As ATEvalResult, before As Long, typeKey As String, diagnostic As ATDiagnostic
    On Error GoTo Failed
    Set normalized = adapter.Normalize(source, "DMY_TEXT", "Бессрочно")
    If normalized.State <> atEvalOK Then
        Set diagnostic = normalized.Diagnostics(1)
        If diagnostic.Code = "TEMPLATE_MISMATCH" Then
            Set Run = Blocked(diagnostic.Code, diagnostic.Message)
        Else
            Set Run = Blocked(diagnostic.Code, "Не удалось подготовить данные заявки. Проверки остановлены.")
        End If
        Exit Function
    End If
    Set context = normalized.Context
    ValidateContext context
    For Each entry In Array(Array("INITIATOR_FULL_NAME", "Фамилия Имя Отчество"), Array("DEPARTMENT", "Подразделение"), Array("POSITION", "Должность"), Array("EMAIL", "Электронная почта"))
        Set value = context.ResolveField(CStr(entry(0)))
        If Not CellError(value, "NET.REQUIRED_FIELDS", CStr(entry(0)), context, issues) Then
            If Blank(value) Then AddIssue issues, "NET.REQUIRED_FIELDS", "FIELD_REQUIRED", "ERROR", CStr(entry(0)), context, "Заполните поле «" & CStr(entry(1)) & "»."
        End If
    Next entry
    AddOutcome checks, issues, "NET.REQUIRED_FIELDS", 0
    before = issues.Count: Set value = context.ResolveField("EMAIL")
    If value.State <> atEvalOK Then
        AddSkip checks, "NET.INITIATOR_EMAIL", "Ошибка ячейки почты уже отмечена в обязательных полях."
    ElseIf Blank(value) Then
        AddSkip checks, "NET.INITIATOR_EMAIL", "Пустая почта уже отмечена в обязательных полях."
    Else
        If Not AT_EmailSyntaxValid(CStr(value.Value.Scalar)) Then AddIssue issues, "NET.INITIATOR_EMAIL", "EMAIL_INVALID", "ERROR", "EMAIL", context, "Укажите корректную электронную почту в формате имя@домен."
        AddOutcome checks, issues, "NET.INITIATOR_EMAIL", before
    End If
    before = issues.Count: Set value = context.ResolveField("REQUEST_TYPE")
    If Not CellError(value, "NET.REQUEST_TYPE", "REQUEST_TYPE", context, issues) Then
        If Not Blank(value) Then typeKey = AT_ExcelHeader(CStr(value.Value.Scalar))
        If Not IsGrant(typeKey) And Not IsRevoke(typeKey) Then AddIssue issues, "NET.REQUEST_TYPE", "REQUEST_TYPE_INVALID", "ERROR", "REQUEST_TYPE", context, "Выберите тип заявки: «Предоставить доступ» или «Отозвать доступ»."
    End If
    AddOutcome checks, issues, "NET.REQUEST_TYPE", before
    before = issues.Count: Set value = context.ResolveField("SYSTEM_IEP")
    If Not CellError(value, "NET.SYSTEM_IEP", "SYSTEM_IEP", context, issues) Then
        If Blank(value) Then
            AddIssue issues, "NET.SYSTEM_IEP", "SYSTEM_IEP_INVALID", "ERROR", "SYSTEM_IEP", context, "В поле «Система в ИЭП?» должно быть выбрано «Да»."
        ElseIf AT_ExcelHeader(CStr(value.Value.Scalar)) <> "да" Then
            AddIssue issues, "NET.SYSTEM_IEP", "SYSTEM_IEP_INVALID", "ERROR", "SYSTEM_IEP", context, "В поле «Система в ИЭП?» должно быть выбрано «Да»."
        End If
    End If
    AddOutcome checks, issues, "NET.SYSTEM_IEP", before
    CheckValidity context, typeKey, issues, checks
    before = issues.Count: Set value = context.ResolveField("JUSTIFICATION")
    If Not CellError(value, "NET.TOP_JUSTIFICATION", "JUSTIFICATION", context, issues) Then
        If Blank(value) Then AddIssue issues, "NET.TOP_JUSTIFICATION", "JUSTIFICATION_EMPTY", "WARNING", "JUSTIFICATION", context, "Добавьте обоснование заявки. Это предупреждение не блокирует исполнение."
    End If
    AddOutcome checks, issues, "NET.TOP_JUSTIFICATION", before
    report.Initialize atEvalOK, "NETWORK_DOCUMENT", issues, checks: Set Run = report
    Exit Function
Failed:
    Set Run = Blocked("NETWORK_CHECK_INPUT_INVALID", "Не удалось проверить поля заявки. Проверьте исходные данные и повторите запуск.")
End Function
Private Sub ValidateContext(ByVal context As ATEvaluationContext)
    Dim id As Variant, value As ATEvalResult
    context.Validate
    For Each id In Array("INITIATOR_FULL_NAME", "DEPARTMENT", "POSITION", "EMAIL", "REQUEST_TYPE", "SYSTEM_IEP", "JUSTIFICATION")
        Set value = context.ResolveField(CStr(id))
        If Len(context.CellRef(CStr(id))) = 0 Then BadInput
        If value.State = atEvalOK Then
            If value.Value.DataType <> "TEXT" Then BadInput
            If value.Value.ValueState <> atValuePresent And value.Value.ValueState <> atValueEmpty Then BadInput
        Else
            If value.State <> atEvalError Or value.Diagnostics.Count <> 1 Then BadInput
            If value.Diagnostics(1).Code <> "EXCEL_CELL_ERROR" Then BadInput
        End If
    Next id
    If Len(context.CellRef("VALIDITY_RAW")) = 0 Then BadInput
End Sub
Private Sub CheckValidity(ByVal context As ATEvaluationContext, ByVal typeKey As String, ByVal issues As Collection, ByVal checks As Collection)
    Dim before As Long, value As ATEvalResult, unlimited As ATEvalResult
    before = issues.Count
    If IsRevoke(typeKey) Then
        AddSkip checks, "NET.VALIDITY", "При отзыве доступа срок действия не проверяется.": Exit Sub
    End If
    If Not IsGrant(typeKey) Then
        AddSkip checks, "NET.VALIDITY", "Сначала укажите корректный тип заявки.": Exit Sub
    End If
    Set value = context.ResolveField("VALIDITY_DATE")
    If Not CellError(value, "NET.VALIDITY", "VALIDITY_RAW", context, issues) Then
        Set unlimited = context.ResolveField("VALIDITY_IS_UNLIMITED")
        If unlimited.State <> atEvalOK Then BadInput
        If Not CBool(unlimited.Value.Scalar) Then
            Select Case value.Value.ValueState
                Case atValuePresent
                    If value.Value.Scalar < AT_GregorianDate(context.RunDate) Then AddIssue issues, "NET.VALIDITY", "VALIDITY_EXPIRED", "ERROR", "VALIDITY_RAW", context, "Срок действия истёк. Укажите сегодняшнюю или будущую дату."
                Case Else
                    AddIssue issues, "NET.VALIDITY", "VALIDITY_INVALID", "ERROR", "VALIDITY_RAW", context, "Укажите «Бессрочно» или корректную дату окончания в формате ДД.ММ.ГГГГ."
            End Select
        End If
    End If
    AddOutcome checks, issues, "NET.VALIDITY", before
End Sub
Private Function Blank(ByVal value As ATEvalResult) As Boolean
    If value.Value.ValueState = atValueEmpty Then
        Blank = True
    Else
        Blank = AT_ExcelBlank(value.Value.Scalar)
    End If
End Function
Private Function IsGrant(ByVal value As String) As Boolean
    IsGrant = (value = "предоставить доступ" Or value = "предоставить")
End Function
Private Function IsRevoke(ByVal value As String) As Boolean
    IsRevoke = (value = "отозвать доступ" Or value = "отозвать")
End Function
Private Function CellError(ByVal value As ATEvalResult, ByVal checkId As String, ByVal fieldId As String, ByVal context As ATEvaluationContext, ByVal issues As Collection) As Boolean
    If value.State = atEvalOK Then Exit Function
    If value.State <> atEvalError Or value.Diagnostics.Count <> 1 Then BadInput
    If value.Diagnostics(1).Code <> "EXCEL_CELL_ERROR" Then BadInput
    AddIssue issues, checkId, "CELL_ERROR", "ERROR", fieldId, context, "В ячейке ошибка Excel. Исправьте значение и повторите проверку."
    CellError = True
End Function
Private Sub AddIssue(ByVal issues As Collection, ByVal checkId As String, ByVal code As String, ByVal severity As String, ByVal fieldId As String, ByVal context As ATEvaluationContext, ByVal message As String)
    Dim issue As New ATCheckIssue
    issue.Initialize checkId, code, severity, fieldId, context.CellRef(fieldId), message: issues.Add issue
End Sub
Private Sub AddOutcome(ByVal checks As Collection, ByVal issues As Collection, ByVal id As String, ByVal before As Long)
    Dim outcome As New ATCheckOutcome, i As Long, status As String
    status = "PASS"
    For i = before + 1 To issues.Count
        If issues(i).Severity = "ERROR" Then status = "FAIL"
    Next i
    outcome.Initialize id, status, "": checks.Add outcome
End Sub
Private Sub AddSkip(ByVal checks As Collection, ByVal id As String, ByVal reason As String)
    Dim outcome As New ATCheckOutcome
    outcome.Initialize id, "SKIPPED", reason: checks.Add outcome
End Sub
Private Function Blocked(ByVal code As String, ByVal message As String) As ATCheckReport
    Dim issues As New Collection, checks As New Collection, issue As New ATCheckIssue, outcome As New ATCheckOutcome, report As New ATCheckReport
    issue.Initialize "NET.INPUT", code, "ERROR", "", "", message: issues.Add issue
    outcome.Initialize "NET.INPUT", "BLOCKED", message: checks.Add outcome
    report.Initialize atEvalError, "NETWORK_DOCUMENT", issues, checks: Set Blocked = report
End Function
Private Sub BadInput()
    Err.Raise vbObjectError + 2941, , "NETWORK_CHECK_INPUT_INVALID"
End Sub
