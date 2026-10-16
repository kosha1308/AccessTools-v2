Attribute VB_Name = "modATNetworkSourceReadTests"
Option Explicit

Public Sub AT_TestNetworkSourceRead()
    Dim book As Workbook, ws As Worksheet, reader As New ATNetworkSourceReader
    Dim result As ATEvalResult, saved As ATRuntimeValue, kind As Variant, data() As Variant
    Dim i As Long, spec As ATExcelTableSpec, layout As ATExcelTableLayout
    Dim savedNumber As Long, savedText As String, stage As String
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1)
    ws.Name = "Source ' test"
    For Each kind In Array("SUPERSET", "INITIATORS", "RULES", "PROTOCOLS", "PROBLEMS", "COMMON_TEXT", "ROUTING", "INITIATOR_IP")
        ws.Cells.Clear
        AT_TestSourceHeaders ws, CStr(kind)
        Set result = reader.Read(ws, CStr(kind))
        AT_Assert result.State = atEvalOK, "published empty schema " & kind
        If result.State = atEvalOK Then
            AT_Assert result.Value.Count = 0 And result.Value.RowTableId = kind & "_RAW", "empty source has typed zero rows " & kind
        End If
    Next kind
    ws.Cells.Clear: AT_TestSourceHeaders ws, "SUPERSET"
    ws.Range("A2:A10185").Value2 = 0
    Set result = reader.Read(ws, "SUPERSET")
    AT_Assert result.State = atEvalOK And result.Value.Count = 0, "unnamed export index is not Superset record"
    ws.Range("B2:I2").Value2 = Array("one.example", "10.0.0.1", "10.0.0.2/32", "10.0.0.3-10.0.0.4", "IS one", "service one", "group one", "team one")
    ws.Range("I2").Formula = "=""team one"""
    ws.Range("B3").Value2 = "hidden.example": ws.Rows(3).Hidden = True
    ws.Columns(8).Hidden = True
    Set result = reader.Read(ws, "SUPERSET")
    AT_Assert result.State = atEvalOK And result.Value.Count = 2, "raw source includes hidden rows, actual count=" & CStr(result.Value.Count)
    Set saved = result.Value
    stage = "raw source fields and failure contracts"
    AT_Assert saved.Row(1).ResolveField("SOURCE_TELEGRAPH").Value.Scalar = "10.0.0.1", "actual SOURCE telegraph mapping"
    AT_Assert saved.Row(1).ResolveField("SOURCE_CXD").Value.Scalar = "10.0.0.2/32", "actual SOURCE CXD mapping"
    AT_Assert saved.Row(1).ResolveField("SOURCE_AWX").Value.Scalar = "10.0.0.3-10.0.0.4", "actual SOURCE AWX mapping"
    AT_Assert saved.Row(1).ResolveField("SKUF_GROUP").Value.Scalar = "group one", "hidden SKUF column retained"
    AT_Assert saved.Row(1).ResolveField("SUPPORT_TEAM").Value.Scalar = "team one", "source formula read by Value2"
    AT_Assert saved.Row(1).CellRef("SOURCE_TELEGRAPH") = "'Source '' test'!C2", "source trace preserves escaped sheet name"
    AT_Assert ws.Range("I2").Formula = "=""team one""" And ws.Rows(3).Hidden And ws.Columns(8).Hidden, "read preserves source formula/visibility"
    ws.Range("B1:I3").AutoFilter Field:=1, Criteria1:="one.example"
    Set result = reader.Read(ws, "SUPERSET")
    AT_Assert result.Value.Count = 2 And ws.FilterMode, "raw snapshot includes filtered rows without changing filter"
    ws.ShowAllData: ws.AutoFilterMode = False
    ws.Range("C2").Value2 = CVErr(xlErrNA)
    Set result = reader.Read(ws, "SUPERSET")
    AT_Assert result.State = atEvalOK And result.Value.Row(1).ResolveField("SOURCE_TELEGRAPH").State = atEvalError, "source cell error retained explicitly"
    AT_Assert saved.Row(1).ResolveField("SOURCE_TELEGRAPH").Value.Scalar = "10.0.0.1", "prior snapshot independent of edits"
    ws.Cells.Clear: AT_TestSourceHeaders ws, "SUPERSET"
    ws.Range("B1").Value2 = "  КОМАНДА   ПОДДЕРЖКИ  ": ws.Range("I1").Value2 = "FQDN"
    ws.Range("B2").Value2 = "team reordered": ws.Range("I2").Value2 = "host reordered"
    Set result = reader.Read(ws, "SUPERSET")
    AT_Assert result.State = atEvalOK And result.Value.Row(1).ResolveField("HOSTNAME").Value.Scalar = "host reordered", "reordered whitespace source headers"
    AT_Assert result.Value.Row(1).ResolveField("SUPPORT_TEAM").Value.Scalar = "team reordered", "reordered support mapping"
    ws.Range("J1").Value2 = "FQDN"
    AT_TestSourceFailure reader.Read(ws, "SUPERSET"), "EXCEL_HEADER_DUPLICATE"
    ws.Range("J1").ClearContents: ws.Range("I1").ClearContents
    AT_TestSourceFailure reader.Read(ws, "SUPERSET"), "EXCEL_HEADER_MISSING"
    AT_TestSourceFailure reader.Read(ws, "UNKNOWN"), "SOURCE_SCHEMA_UNKNOWN"
    ws.Cells.Clear: AT_TestSourceHeaders ws, "SUPERSET"
    ws.Range("B5:I5").Value2 = ws.Range("B1:I1").Value2
    AT_TestSourceFailure reader.Read(ws, "SUPERSET"), "EXCEL_HEADER_AMBIGUOUS"
    ws.Range("B5:I5").ClearContents: ws.Range("B1:B2").Merge
    AT_TestSourceFailure reader.Read(ws, "SUPERSET"), "SOURCE_MERGED_TABLE"
    ws.Cells.UnMerge
    ws.Cells.Clear: AT_TestSourceHeaders ws, "PROTOCOLS"
    ws.Range("B10").Formula = "="""""
    Set result = reader.Read(ws, "PROTOCOLS")
    AT_Assert result.Value.Count = 0, "formula displaying blank does not create source row"
    ws.Range("B10").ClearContents
    ws.Range("A2:C2").Value2 = Array(0, "TCP", "fictional")
    Set result = reader.Read(ws, "PROTOCOLS")
    AT_Assert result.Value.Row(1).ResolveField("ACTIVE").Value.Scalar = "0", "numeric zero is retained, not empty or guessed active"
    AT_Assert result.Value.Row(1).ResolveField("PROTOCOL").Value.Scalar = "TCP", "protocol body column maps by header"
    ws.Range("A2:C2").Value2 = Array("Да", "TCP", "fictional")
    ws.Range("B2:B3").Merge
    AT_TestSourceFailure reader.Read(ws, "PROTOCOLS"), "SOURCE_MERGED_TABLE"
    ws.Cells.UnMerge: ws.Cells.Clear: AT_TestSourceHeaders ws, "PROTOCOLS"
    ws.Range("A100002").Value2 = "Да"
    AT_TestSourceFailure reader.Read(ws, "PROTOCOLS"), "SOURCE_ROW_LIMIT"
    ws.Cells.Clear: AT_TestSourceHeaders ws, "PROTOCOLS"
    ws.Range("BC1").Value2 = ws.Range("C1").Value2: ws.Range("C1").ClearContents
    ws.Range("A20001").Value2 = "Да"
    AT_TestSourceFailure reader.Read(ws, "PROTOCOLS"), "SOURCE_CELL_LIMIT"
    ws.Cells.Clear: AT_TestSourceHeaders ws, "SUPERSET"
    ReDim data(1 To 10001, 1 To 1)
    For i = 1 To 10001: data(i, 1) = "10.0.0.1": Next i
    ws.Range("C2:C10002").Value2 = data
    Set result = reader.Read(ws, "SUPERSET")
    AT_Assert result.State = atEvalOK, "Superset exceeds request ten-thousand-row limit"
    If result.State = atEvalOK Then
        AT_Assert result.Value.Count = 10001 And result.Value.Row(10001).RowId = "10002", "large source retains final row"
    End If
    book.Close False: Set book = Nothing
    AT_Assert saved.Row(1).ResolveField("SERVICE").Value.Scalar = "service one", "snapshot survives source closure"
    AT_TestSourceFailure reader.Read(Nothing, "SUPERSET"), "SOURCE_READ_FAILED"
    Exit Sub
CleanupError:
    savedNumber = Err.Number: savedText = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise savedNumber, "AT_TestNetworkSourceRead", stage & ": " & savedText
End Sub

Private Sub AT_TestSourceFailure(ByVal result As ATEvalResult, ByVal code As String)
    Dim partial As ATRuntimeValue, number As Long
    AT_Assert result.State = atEvalError, "source failure state " & code
    AT_Assert result.Diagnostics(1).Code = code, "source failure code " & code
    On Error Resume Next
    Set partial = result.Value: number = Err.Number
    On Error GoTo 0
    AT_Assert number <> 0 And partial Is Nothing, "no partial source rows " & code
End Sub

Private Sub AT_TestSourceHeaders(ByVal ws As Worksheet, ByVal kind As String)
    Dim headers As Variant, i As Long, firstCol As Long
    firstCol = 1
    Select Case kind
        Case "SUPERSET"
            firstCol = 2
            headers = Array("FQDN", "SOURCE:ТЕЛЕГРАФ", "SOURCE: ЦХД", "SOURCE: AWX", "ПОЛНОЕ ИМЯ ИС", "СЕРВИС", "СКУФ ГРУППА ТП", "КОМАНДА ПОДДЕРЖКИ")
        Case "INITIATORS": headers = Array("Активно", "Домены / почты", "Требуемое согласование", "Уровень сообщения", "Комментарий / инструкция", "Служебный режим")
        Case "RULES": headers = Array("Активно", "Название правила", "IP", "Область IP", "Hostname/FQDN", "Тип совпадения hostname", "Порт назначения", "Протокол назначения", "Сервис — условие", "Сервис — результат", "Группа назначения", "Требуемое согласование", "Комментарий / инструкция", "Уровень сообщения", "Инструкция")
        Case "PROTOCOLS": headers = Array("Активно", "Протокол", "Комментарий")
        Case "PROBLEMS": headers = Array("Поле заявки", "Что проверяется", "Штатное сообщение", "Сообщение пользователю", "Сообщение координатору", "Уровень сообщения", "Ошибка пользователя", "Подсвечивать в заявке", "Активно", "Ключ")
        Case "COMMON_TEXT": headers = Array("Ключ", "Значение")
        Case "ROUTING": headers = Array("Проверка", "Поле / результат", "Что проверяется", "В «Все проблемы»", "Подсветка / комментарий в заявке", "Результат / комментарий в Superset", "Активно", "Ключ")
        Case "INITIATOR_IP": headers = Array("Активно", "Email / домен", "Область IP", "Разрешённые IP", "Запрещённые IP (НЕ)", "Комментарий / инструкция", "Учетная запись для портов БД")
    End Select
    For i = 0 To UBound(headers): ws.Cells(1, firstCol + i).Value2 = headers(i): Next i
End Sub
