Attribute VB_Name = "modATNetworkProfileTests"
Option Explicit
Public Sub AT_TestNetworkProfile()
    Dim book As Workbook, ws As Worksheet, profile As ATTemplateProfile, decision As ATTemplateGateResult
    Dim number As Long, description As String
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    Set profile = AT_NetworkProfileV3()
    AT_Assert profile.Verify(ws).Allowed, "real form structure synthetic fixture passes NETWORK_ACCESS profile"
    ws.Range("C7").Value2 = "Отозвать доступ": ws.Range("C8").Value2 = "Нет": ws.Range("C12").Value2 = "2030-12-31"
    ws.Range("A18").Value2 = "Система"
    AT_Assert profile.Verify(ws).Allowed, "input/default edits and isolated repeated lower header do not change fingerprint"
    ws.Rows(15).Insert
    AT_Assert profile.Verify(ws).Allowed, "real form supports inserted technical data row"
    ws.Range("D14").Value2 = "Old port column"
    Set decision = profile.Verify(ws)
    AT_Assert Not decision.Allowed, "whole technical header group required"
    ws.Range("D14").Value2 = vbLf & "Порт назначения"
    ws.Range("C17:D17").UnMerge
    AT_Assert Not profile.Verify(ws).Allowed, "individual grouped column merge shape enforced"
    ws.Range("C17:D17").Merge
    AT_Assert profile.Verify(ws).Allowed, "profile reusable after grouped merge repaired"
    book.Close False
    Exit Sub
CleanupError:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise number, , description
End Sub
Public Sub AT_TestNetworkReferenceFixture(ByVal ws As Worksheet)
    Dim item As Variant
    For Each item In Array("A1:G2", "A3:G3", "A4:C4", "A5:C5", "A6:G6", "A7:B7", "C7:G7", "A8:B8", "C8:G8", "A9:B10", "C9:F9", "C10:F10", "A11:B11", "C11:G11", "A12:B12", "C12:G12", "A13:G13", "E14:G14", "E15:G15", "C16:D16", "E16:G16", "C17:D17", "E17:G17", "C18:D18", "E18:G18")
        ws.Range(CStr(item)).Merge
    Next item
    ws.Range("A1").Value2 = "Заявка на сетевую связность (v.3.0)"
    ws.Range("A3").Value2 = "Инициатор заявки"
    ws.Range("A4").Value2 = "Фамилия Имя Отчество": ws.Range("D4").Value2 = "Подразделение"
    ws.Range("E4").Value2 = "Должность": ws.Range("F4").Value2 = "Электронная почта": ws.Range("G4").Value2 = "Телефон / @telegram"
    ws.Range("A6").Value2 = "Описание доступа"
    ws.Range("A7").Value2 = "Тип заявки": ws.Range("A8").Value2 = "Система в ИЭП?"
    ws.Range("A9").Value2 = "Поля обязательные для портов БД": ws.Range("C9").Value2 = "Учетная запись:"
    ws.Range("G9").Value2 = "Планируется передача персональнах данных по заказываемой связанности?"
    ws.Range("A11").Value2 = "Обоснование"
    ws.Range("A12").Value2 = "Срок действия (указывается либо бессрочно, либо дата, до которой будет действовать связность)"
    ws.Range("A13").Value2 = "Техническая спецификация доступа (обязательно для заполнения)"
    ws.Range("A14:D14").Value2 = Array("IP источника", "IP назначения", "Протокол назначения", vbLf & "Порт назначения")
    ws.Range("E14").Value2 = "Обоснование "
    ws.Range("A16").Value2 = "Система": ws.Range("B16").Value2 = "Адрес": ws.Range("C16").Value2 = "Имя хоста"
    ws.Range("E16").Value2 = "Роль/Назначение (В случае использования виртуализации (трансляции) указывается реальный и виртуализируемый (транслируемый) адреса)"
End Sub
Public Function AT_VerifyNetworkReference(ByVal workbookName As String) As String
    Dim profile As ATTemplateProfile, result As ATTemplateGateResult
    Set profile = AT_NetworkProfileV3()
    Set result = profile.Verify(Application.Workbooks(workbookName).Worksheets(1))
    AT_VerifyNetworkReference = CStr(result.Allowed) & "|" & result.Code & "|" & result.Detail
End Function
