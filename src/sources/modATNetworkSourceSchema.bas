Attribute VB_Name = "modATNetworkSourceSchema"
Option Explicit

Public Function AT_NetworkSourceSpec(ByVal sourceId As String) As ATExcelTableSpec
    Dim spec As New ATExcelTableSpec, fields As Variant, headers As Variant, i As Long
    Select Case sourceId
        Case "SUPERSET"
            fields = Array("HOSTNAME", "SOURCE_TELEGRAPH", "SOURCE_CXD", "SOURCE_AWX", "IS_NAME", "SERVICE", "SKUF_GROUP", "SUPPORT_TEAM")
            headers = Array("FQDN", "SOURCE:ТЕЛЕГРАФ", "SOURCE: ЦХД", "SOURCE: AWX", "ПОЛНОЕ ИМЯ ИС", "СЕРВИС", "СКУФ ГРУППА ТП", "КОМАНДА ПОДДЕРЖКИ")
        Case "INITIATORS"
            fields = Array("ACTIVE", "EMAIL_DOMAINS", "APPROVAL", "SEVERITY", "COMMENT", "SERVICE_MODE")
            headers = Array("Активно", "Домены / почты", "Требуемое согласование", "Уровень сообщения", "Комментарий / инструкция", "Служебный режим")
        Case "RULES"
            fields = Array("ACTIVE", "RULE_NAME", "IP", "IP_SCOPE", "HOSTNAME", "HOSTNAME_MATCH", "DEST_PORT", "DEST_PROTOCOL", "SERVICE_CONDITION", "SERVICE_RESULT", "DEST_GROUP", "APPROVAL", "COMMENT", "SEVERITY", "INSTRUCTION")
            headers = Array("Активно", "Название правила", "IP", "Область IP", "Hostname/FQDN", "Тип совпадения hostname", "Порт назначения", "Протокол назначения", "Сервис — условие", "Сервис — результат", "Группа назначения", "Требуемое согласование", "Комментарий / инструкция", "Уровень сообщения", "Инструкция")
        Case "PROTOCOLS"
            fields = Array("ACTIVE", "PROTOCOL", "COMMENT")
            headers = Array("Активно", "Протокол", "Комментарий")
        Case "PROBLEMS"
            fields = Array("REQUEST_FIELD", "CHECK_DESCRIPTION", "DEFAULT_MESSAGE", "USER_MESSAGE", "COORDINATOR_MESSAGE", "SEVERITY", "USER_ERROR", "HIGHLIGHT", "ACTIVE", "KEY")
            headers = Array("Поле заявки", "Что проверяется", "Штатное сообщение", "Сообщение пользователю", "Сообщение координатору", "Уровень сообщения", "Ошибка пользователя", "Подсвечивать в заявке", "Активно", "Ключ")
        Case "COMMON_TEXT"
            fields = Array("KEY", "VALUE")
            headers = Array("Ключ", "Значение")
        Case "ROUTING"
            fields = Array("CHECK_ID", "FIELD_RESULT", "CHECK_DESCRIPTION", "IN_USER_REPORT", "IN_REQUEST", "IN_SUPERSET", "ACTIVE", "KEY")
            headers = Array("Проверка", "Поле / результат", "Что проверяется", "В «Все проблемы»", "Подсветка / комментарий в заявке", "Результат / комментарий в Superset", "Активно", "Ключ")
        Case "INITIATOR_IP"
            fields = Array("ACTIVE", "EMAIL_DOMAIN", "IP_SCOPE", "ALLOWED_IP", "DENIED_IP", "COMMENT", "DB_ACCOUNT")
            headers = Array("Активно", "Email / домен", "Область IP", "Разрешённые IP", "Запрещённые IP (НЕ)", "Комментарий / инструкция", "Учетная запись для портов БД")
        Case Else: Err.Raise vbObjectError + 2860, , "SOURCE_SCHEMA_UNKNOWN"
    End Select
    spec.Initialize sourceId & "_RAW", 1, 20, 1, 64
    For i = 0 To UBound(fields)
        spec.AddColumn CStr(fields(i)), CStr(headers(i)), "DIRECT"
    Next i
    Set AT_NetworkSourceSpec = spec
End Function
