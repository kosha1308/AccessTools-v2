Attribute VB_Name = "modATPolicyBindings"
Option Explicit

' Explicit compatibility metadata, not another implementation of the checks.
Public Function AT_NetworkIssuePolicyKey(ByVal issue As ATCheckIssue) As String
    Dim field As String, check As String
    check = issue.CheckId
    Select Case check
        Case "NET.INITIATOR_ALLOWED"
            If issue.Code = "INITIATOR_NOT_ALLOWED" Then field = "Электронная почта"
        Case "NET.INITIATOR_APPROVAL_RULES"
            If issue.Code = "INITIATOR_APPROVAL_REQUIRED" Then field = "Дополнительное согласование"
        Case "NET.INITIATOR_IP_ACCESS"
            If issue.Code = "INITIATOR_IP_DENIED" Then field = AT_FieldTitle(issue.FieldId)
            If issue.Code = "INITIATOR_DB_ACCOUNT_FORBIDDEN" Then field = "Учетная запись"
        Case "NET.REQUIRED_FIELDS"
            If issue.Code = "FIELD_REQUIRED" Then
                Select Case issue.FieldId
                    Case "INITIATOR_FULL_NAME": field = "Фамилия Имя Отчество"
                    Case "DEPARTMENT": field = "Подразделение"
                    Case "POSITION": field = "Должность"
                    Case "EMAIL": field = "Электронная почта"
                End Select
            End If
        Case "NET.INITIATOR_EMAIL"
            If issue.Code = "EMAIL_INVALID" Then field = "Электронная почта"
        Case "NET.REQUEST_TYPE"
            If issue.Code = "REQUEST_TYPE_INVALID" Then field = "Тип заявки"
        Case "NET.SYSTEM_IEP"
            If issue.Code = "SYSTEM_IEP_INVALID" Then field = "Система в ИЭП?"
        Case "NET.TOP_JUSTIFICATION"
            If issue.Code = "JUSTIFICATION_EMPTY" Then field = "Обоснование"
        Case "NET.VALIDITY"
            If issue.Code = "VALIDITY_INVALID" Or issue.Code = "VALIDITY_EXPIRED" Then field = "Срок действия"
        Case "NET.ADDRESS_MATCH"
            If issue.Code = "ADDRESS_EXTRA" Then field = "Адрес"
            If issue.Code = "ADDRESS_MISSING" Then field = "Техническая таблица"
        Case "NET.PROTOCOLS"
            If issue.Code = "PROTOCOL_INVALID" Then check = "NET.TECHNICAL_TABLE": field = "Протокол назначения"
        Case "NET.TECHNICAL_TABLE"
            Select Case issue.Code
                Case "ROW_DUPLICATE": field = "Дубликат связанности"
                Case "FIELD_RECOMMENDED"
                    If issue.FieldId = "TECH_REASON" Then field = "Обоснование связанности"
                Case "FIELD_REQUIRED", "IP_INVALID", "PORT_INVALID"
                    Select Case issue.FieldId
                        Case "TECH_SOURCE_IP": field = "IP источника"
                        Case "TECH_DEST_IP": field = "IP назначения"
                        Case "TECH_DEST_PORT": field = "Порт назначения"
                        Case "TECH_PROTOCOL": field = "Протокол назначения"
                    End Select
            End Select
        Case "NET.ADDRESS_TABLE"
            Select Case issue.Code
                Case "FIELD_REQUIRED", "IP_INVALID"
                    If issue.FieldId = "ADDRESS_IP" Then field = "Адрес"
                    If issue.FieldId = "HOSTNAME" Then field = "Имя хоста"
                Case "FIELD_RECOMMENDED"
                    If issue.FieldId = "SYSTEM" Then field = "Система"
                    If issue.FieldId = "ROLE" Then field = "Роль/Назначение"
            End Select
    End Select
    If Len(field) > 0 Then AT_NetworkIssuePolicyKey = check & "|" & issue.Severity & "|" & field
End Function
