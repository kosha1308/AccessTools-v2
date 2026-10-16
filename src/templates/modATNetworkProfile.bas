Attribute VB_Name = "modATNetworkProfile"
Option Explicit

' Published structural description of the user-supplied empty v3.0 form.
' User-entered values and validation-list helper cells are intentionally excluded.
Public Function AT_NetworkProfileV3() As ATTemplateProfile
    Dim profile As New ATTemplateProfile, item As Variant
    profile.Initialize "NETWORK_ACCESS", "3.0-profile1", "Лист1", False
    NetworkAnchor profile, "TITLE", "Заявка на сетевую связность (v.3.0)", 1, 1, 7, 2
    NetworkAnchor profile, "INITIATOR_SECTION", "Инициатор заявки", 3, 1, 7
    For Each item In NetworkDocumentMetadata()
        NetworkAnchor profile, Left$(CStr(item(0)), Len(CStr(item(0))) - 4), CStr(item(1)), CLng(item(2)), CLng(item(3)), CLng(item(4))
    Next item
    NetworkAnchor profile, "ACCESS_SECTION", "Описание доступа", 6, 1, 7
    NetworkAnchor profile, "DB_SECTION", "Поля обязательные для портов БД", 9, 1, 2, 2
    NetworkAnchor profile, "TECH_SECTION", "Техническая спецификация доступа (обязательно для заполнения)", 13, 1, 7
    profile.AddTableAnchor "TECH_TABLE", 14, 14, 1, 7, NetworkTechnicalHeaders()
    profile.AddTableAnchor "ADDRESS_TABLE", 15, 10000, 1, 7, NetworkAddressHeaders()
    profile.AddRelation "TITLE", "TECH_TABLE", 13, 13, 0, 0
    profile.AddRelation "TECH_TABLE", "ADDRESS_TABLE", 1, 9986, 0, 0
    profile.Seal
    Set AT_NetworkProfileV3 = profile
End Function

Private Sub NetworkAnchor(ByVal profile As ATTemplateProfile, ByVal id As String, ByVal text As String, ByVal row As Long, ByVal col As Long, ByVal width As Long, Optional ByVal height As Long = 1)
    profile.AddAnchor id, text, row, row + height - 1, col, col + width - 1, height, width, False
End Sub

Private Function NetworkTechnicalHeaders() As Variant
    NetworkTechnicalHeaders = Array(Array("IP источника", 1, 1), Array("IP назначения", 2, 1), _
        Array("Протокол назначения", 3, 1), Array("Порт назначения", 4, 1), Array("Обоснование", 5, 3))
End Function
Private Function NetworkAddressHeaders() As Variant
    NetworkAddressHeaders = Array(Array("Система", 1, 1), Array("Адрес", 2, 1), Array("Имя хоста", 3, 2), _
        Array("Роль/Назначение (В случае использования виртуализации (трансляции) указывается реальный и виртуализируемый (транслируемый) адреса)", 5, 3))
End Function
Private Function NetworkDocumentMetadata() As Variant
    NetworkDocumentMetadata = Array( _
        Array("INITIATOR_FULL_NAME_RAW", "Фамилия Имя Отчество", 4, 1, 3, 1, 0), _
        Array("DEPARTMENT_RAW", "Подразделение", 4, 4, 1, 1, 0), _
        Array("POSITION_RAW", "Должность", 4, 5, 1, 1, 0), _
        Array("EMAIL_RAW", "Электронная почта", 4, 6, 1, 1, 0), _
        Array("CONTACT_RAW", "Телефон / @telegram", 4, 7, 1, 1, 0), _
        Array("REQUEST_TYPE_RAW", "Тип заявки", 7, 1, 2, 0, 1), _
        Array("SYSTEM_IEP_RAW", "Система в ИЭП?", 8, 1, 2, 0, 1), _
        Array("DB_ACCOUNT_RAW", "Учетная запись:", 9, 3, 4, 1, 0), _
        Array("PERSONAL_DATA_RAW", "Планируется передача персональнах данных по заказываемой связанности?", 9, 7, 1, 1, 0), _
        Array("JUSTIFICATION_RAW", "Обоснование", 11, 1, 2, 0, 1), _
        Array("VALIDITY_RAW", "Срок действия (указывается либо бессрочно, либо дата, до которой будет действовать связность)", 12, 1, 2, 0, 1))
End Function
Public Function AT_NetworkDocumentFieldsV3() As Collection
    Dim result As New Collection, item As Variant, spec As ATExcelFieldSpec
    For Each item In NetworkDocumentMetadata()
        Set spec = New ATExcelFieldSpec
        spec.Initialize CStr(item(0)), CStr(item(1)), CLng(item(2)), CLng(item(2)), CLng(item(3)), CLng(item(3) + item(4) - 1), CLng(item(5)), CLng(item(6))
        result.Add spec
    Next item
    Set AT_NetworkDocumentFieldsV3 = result
End Function
Public Function AT_NetworkTechnicalSpecV3() As ATExcelTableSpec
    Dim spec As New ATExcelTableSpec, item As Variant, ids As Variant, i As Long, usage As String
    spec.Initialize "TECHNICAL_TABLE_RAW", 14, 14, 1, 7
    ids = Array("SOURCE_RAW", "DESTINATION_RAW", "PROTOCOL_RAW", "PORT_RAW", "REASON_RAW")
    For Each item In NetworkTechnicalHeaders()
        usage = "INHERITED": If i = 4 Then usage = "NONE"
        spec.AddColumn CStr(ids(i)), CStr(item(0)), usage: i = i + 1
    Next item
    Set AT_NetworkTechnicalSpecV3 = spec
End Function
Public Function AT_NetworkAddressSpecV3() As ATExcelTableSpec
    Dim spec As New ATExcelTableSpec, item As Variant, ids As Variant, i As Long
    spec.Initialize "ADDRESS_TABLE_RAW", 15, 10000, 1, 7
    ids = Array("SYSTEM_RAW", "ADDRESS_RAW", "HOSTNAME_RAW", "ROLE_RAW")
    For Each item In NetworkAddressHeaders()
        spec.AddColumn CStr(ids(i)), CStr(item(0)), "DIRECT": i = i + 1
    Next item
    Set AT_NetworkAddressSpecV3 = spec
End Function
