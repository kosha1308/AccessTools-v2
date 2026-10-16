Attribute VB_Name = "modATNetworkProfile"
Option Explicit

' Published structural description of the user-supplied empty v3.0 form.
' User-entered values and validation-list helper cells are intentionally excluded.
Public Function AT_NetworkProfileV3() As ATTemplateProfile
    Dim profile As New ATTemplateProfile
    profile.Initialize "NETWORK_ACCESS", "3.0-profile1", "Лист1", False
    NetworkAnchor profile, "TITLE", "Заявка на сетевую связность (v.3.0)", 1, 1, 7, 2
    NetworkAnchor profile, "INITIATOR_SECTION", "Инициатор заявки", 3, 1, 7
    NetworkAnchor profile, "INITIATOR_NAME", "Фамилия Имя Отчество", 4, 1, 3
    NetworkAnchor profile, "DEPARTMENT", "Подразделение", 4, 4, 1
    NetworkAnchor profile, "POSITION", "Должность", 4, 5, 1
    NetworkAnchor profile, "EMAIL", "Электронная почта", 4, 6, 1
    NetworkAnchor profile, "CONTACT", "Телефон / @telegram", 4, 7, 1
    NetworkAnchor profile, "ACCESS_SECTION", "Описание доступа", 6, 1, 7
    NetworkAnchor profile, "REQUEST_TYPE", "Тип заявки", 7, 1, 2
    NetworkAnchor profile, "SYSTEM_IEP", "Система в ИЭП?", 8, 1, 2
    NetworkAnchor profile, "DB_SECTION", "Поля обязательные для портов БД", 9, 1, 2, 2
    NetworkAnchor profile, "DB_ACCOUNT", "Учетная запись:", 9, 3, 4
    NetworkAnchor profile, "PERSONAL_DATA", "Планируется передача персональнах данных по заказываемой связанности?", 9, 7, 1
    NetworkAnchor profile, "JUSTIFICATION", "Обоснование", 11, 1, 2
    NetworkAnchor profile, "VALIDITY", "Срок действия (указывается либо бессрочно, либо дата, до которой будет действовать связность)", 12, 1, 2
    NetworkAnchor profile, "TECH_SECTION", "Техническая спецификация доступа (обязательно для заполнения)", 13, 1, 7
    profile.AddTableAnchor "TECH_TABLE", 14, 14, 1, 7, Array( _
        Array("IP источника", 1, 1), Array("IP назначения", 2, 1), _
        Array("Протокол назначения", 3, 1), Array("Порт назначения", 4, 1), Array("Обоснование", 5, 3))
    profile.AddTableAnchor "ADDRESS_TABLE", 15, 10000, 1, 7, Array( _
        Array("Система", 1, 1), Array("Адрес", 2, 1), Array("Имя хоста", 3, 2), _
        Array("Роль/Назначение (В случае использования виртуализации (трансляции) указывается реальный и виртуализируемый (транслируемый) адреса)", 5, 3))
    profile.AddRelation "TITLE", "TECH_TABLE", 13, 13, 0, 0
    profile.AddRelation "TECH_TABLE", "ADDRESS_TABLE", 1, 9986, 0, 0
    profile.Seal
    Set AT_NetworkProfileV3 = profile
End Function

Private Sub NetworkAnchor(ByVal profile As ATTemplateProfile, ByVal id As String, ByVal text As String, ByVal row As Long, ByVal col As Long, ByVal width As Long, Optional ByVal height As Long = 1)
    profile.AddAnchor id, text, row, row + height - 1, col, col + width - 1, height, width, False
End Sub
