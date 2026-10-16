Attribute VB_Name = "modATReportLabels"
Option Explicit
Public Function AT_CheckTitle(ByVal id As String) As String
    Select Case id
        Case "NET.REQUIRED_FIELDS": AT_CheckTitle = "Обязательные поля инициатора"
        Case "NET.INITIATOR_EMAIL": AT_CheckTitle = "Формат электронной почты"
        Case "NET.REQUEST_TYPE": AT_CheckTitle = "Тип заявки"
        Case "NET.SYSTEM_IEP": AT_CheckTitle = "Система в ИЭП"
        Case "NET.VALIDITY": AT_CheckTitle = "Срок действия"
        Case "NET.TOP_JUSTIFICATION": AT_CheckTitle = "Обоснование заявки"
        Case "NET.TECHNICAL_TABLE": AT_CheckTitle = "Техническая таблица"
        Case "NET.PROTOCOLS": AT_CheckTitle = "Разрешённые протоколы"
        Case "NET.ADDRESS_TABLE": AT_CheckTitle = "Таблица адресов"
        Case "NET.ADDRESS_MATCH": AT_CheckTitle = "Соответствие адресов двух таблиц"
        Case "NET.INPUT", "NET.TABLE_INPUT": AT_CheckTitle = "Подготовка заявки"
        Case Else: AT_CheckTitle = "Проверка заявки"
    End Select
End Function
Public Function AT_FieldTitle(ByVal id As String) As String
    Select Case id
        Case "INITIATOR_FULL_NAME": AT_FieldTitle = "Фамилия Имя Отчество"
        Case "DEPARTMENT": AT_FieldTitle = "Подразделение"
        Case "POSITION": AT_FieldTitle = "Должность"
        Case "EMAIL": AT_FieldTitle = "Электронная почта"
        Case "REQUEST_TYPE": AT_FieldTitle = "Тип заявки"
        Case "SYSTEM_IEP": AT_FieldTitle = "Система в ИЭП?"
        Case "VALIDITY", "VALIDITY_RAW": AT_FieldTitle = "Срок действия"
        Case "JUSTIFICATION", "TECH_REASON": AT_FieldTitle = "Обоснование"
        Case "TECH_SOURCE_IP": AT_FieldTitle = "IP источника"
        Case "TECH_DEST_IP": AT_FieldTitle = "IP назначения"
        Case "TECH_PROTOCOL": AT_FieldTitle = "Протокол назначения"
        Case "TECH_DEST_PORT": AT_FieldTitle = "Порт назначения"
        Case "SYSTEM": AT_FieldTitle = "Система"
        Case "ADDRESS_IP": AT_FieldTitle = "Адрес"
        Case "HOSTNAME": AT_FieldTitle = "Имя хоста"
        Case "ROLE": AT_FieldTitle = "Роль / назначение"
        Case "TECHNICAL_ROWS": AT_FieldTitle = "Техническая таблица"
        Case "ADDRESS_ROWS": AT_FieldTitle = "Таблица адресов"
        Case "": AT_FieldTitle = "—"
        Case Else: AT_FieldTitle = "Поле заявки"
    End Select
End Function
Public Function AT_StatusTitle(ByVal status As String) As String
    Select Case status
        Case "PASS": AT_StatusTitle = "Ошибок нет"
        Case "FAIL": AT_StatusTitle = "Есть ошибки"
        Case "SKIPPED": AT_StatusTitle = "Пропущена"
        Case Else: AT_StatusTitle = "Остановлена"
    End Select
End Function
Public Function AT_SeverityTitle(ByVal severity As String) As String
    Select Case severity
        Case "ERROR": AT_SeverityTitle = "Ошибка"
        Case "WARNING": AT_SeverityTitle = "Предупреждение"
        Case Else: AT_SeverityTitle = "Информация"
    End Select
End Function
