Attribute VB_Name = "modATNetworkReadTests"
Option Explicit
Public Sub AT_TestNetworkMetadata()
    Dim book As Workbook, ws As Worksheet, specs As Collection, spec As ATExcelFieldSpec
    Dim reader As New ATExcelFieldReader, field As ATExcelFieldValue, located As Boolean
    Dim tableSpec As ATExcelTableSpec, layout As ATExcelTableLayout, tableReader As New ATExcelTableReader
    Dim result As ATEvalResult, number As Long, description As String
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    ws.Range("C10").Value2 = "db_test"
    Set specs = AT_NetworkDocumentFieldsV3()
    AT_Assert specs.Count = 11, "all eleven actual document fields described"
    For Each spec In specs
        If spec.FieldId = "DB_ACCOUNT_RAW" Then
            Set field = reader.Read(ws, spec): located = True
            AT_Assert field.Result.Value.Scalar = "db_test" And field.CellRef = "'Лист1'!C10", "DB account metadata selects actual lower value cell"
        End If
    Next spec
    AT_Assert located, "DB account included in document field catalog"
    Set tableSpec = AT_NetworkTechnicalSpecV3(): Set layout = tableSpec.Locate(ws)
    ws.Range("A15:D15").Value2 = Array("192.0.2.1", "198.51.100.1", "TCP", 443)
    Set result = tableReader.Read(ws, layout, 15)
    AT_Assert result.State = atEvalOK, "shared technical metadata locates actual header"
    AT_Assert result.Value.Row(1).ResolveField("PORT_RAW").Value.Scalar = "443", "technical field ids and raw values usable"
    Set tableSpec = AT_NetworkAddressSpecV3(): Set layout = tableSpec.Locate(ws)
    ws.Range("B17").Value2 = "192.0.2.1"
    Set result = tableReader.Read(ws, layout, 18)
    AT_Assert result.Value.Count = 1 And result.Value.Row(1).ResolveField("ADDRESS_RAW").Value.Scalar = "192.0.2.1", "shared address metadata skips trailing empty physical row"
    book.Close False
    Exit Sub
CleanupError:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise number, , description
End Sub
Public Sub AT_TestNetworkRead()
    Dim book As Workbook, ws As Worksheet, reader As New ATNetworkRequestReader
    Dim result As ATNetworkReadResult, context As ATEvaluationContext, snapshot As ATEvaluationContext
    Dim field As ATEvalResult, rows As ATRuntimeValue, item As Variant, expected As Variant
    Dim caught As Boolean, number As Long, description As String
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1): ws.Name = "Лист1"
    AT_TestNetworkReferenceFixture ws
    Set result = reader.Read(ws, "RUN1", "2026-10-16", "Europe/Moscow")
    AT_Assert result.State = atEvalOK, "blank actual form can be extracted without business validation"
    Set context = result.Context
    AT_Assert context.ResolveField("TECHNICAL_ROWS_RAW").Value.Count = 0 And context.ResolveField("ADDRESS_ROWS_RAW").Value.Count = 0, "blank form has empty typed raw collections"
    ws.Range("A5").Value2 = "Тестовый пользователь": ws.Range("D5").Value2 = "Отдел"
    ws.Range("E5").Value2 = "Роль": ws.Range("F5").Value2 = "user@example.test": ws.Range("G5").Value2 = "@test"
    ws.Range("C7").Value2 = "Предоставить доступ": ws.Range("C8").Value2 = "Да"
    ws.Range("C10").Value2 = "db_test": ws.Range("G10").Value2 = "Нет"
    ws.Range("C11").Value2 = "line1" & vbLf & "line2": ws.Range("C12").Value2 = 46100#
    ws.Rows(15).Resize(2).Insert
    ws.Range("A15:D15").Value2 = Array("192.0.2.1", "198.51.100.1", "TCP", 443)
    ws.Range("A17:D17").Value2 = Array("192.0.2.2", "198.51.100.2", "TCP", "80")
    ws.Range("E15:G17").Merge: ws.Range("E15").Value2 = "shared reason"
    ws.Range("A19:A21").Merge: ws.Range("A19").Value2 = "test system"
    ws.Range("B19").Value2 = "192.0.2.1": ws.Range("C19").Value2 = "host.example.test"
    ws.Range("B21").Value2 = "192.0.2.2"
    ws.Range("E19:G21").Merge: ws.Range("E19").Value2 = "shared role"
    ws.Range("B1048576").Interior.Color = vbYellow
    Set result = reader.Read(ws, "RUN2", "2026-10-16", "Europe/Moscow")
    AT_Assert result.State = atEvalOK, "complete merged/spacer form extracted in one run"
    If result.State <> atEvalOK Then GoTo Finish
    Set context = result.Context: Set snapshot = context.Snapshot
    For Each item In Array(Array("INITIATOR_FULL_NAME_RAW", "Тестовый пользователь", "A5"), Array("DEPARTMENT_RAW", "Отдел", "D5"), Array("POSITION_RAW", "Роль", "E5"), Array("EMAIL_RAW", "user@example.test", "F5"), Array("CONTACT_RAW", "@test", "G5"), Array("REQUEST_TYPE_RAW", "Предоставить доступ", "C7"), Array("SYSTEM_IEP_RAW", "Да", "C8"), Array("DB_ACCOUNT_RAW", "db_test", "C10"), Array("PERSONAL_DATA_RAW", "Нет", "G10"), Array("JUSTIFICATION_RAW", "line1" & vbLf & "line2", "C11"), Array("VALIDITY_RAW", "46100", "C12"))
        Set field = context.ResolveField(CStr(item(0)))
        AT_Assert field.State = atEvalOK And field.Value.Scalar = item(1), "document raw extraction " & CStr(item(0))
        AT_Assert context.CellRef(CStr(item(0))) = "'Лист1'!" & CStr(item(2)), "document actual source cell " & CStr(item(0))
    Next item
    AT_Assert context.ResolveField("VALIDITY_VALUE2_KIND").Value.Scalar = "NUMBER" And context.ResolveField("VALIDITY_SERIAL").Value.Scalar = 46100#, "numeric validity retains serial separate from raw text"
    Set rows = context.ResolveField("TECHNICAL_ROWS_RAW").Value
    AT_Assert rows.Count = 2 And rows.Row(2).RowId = "17", "upper blank under merged reason not connectivity"
    AT_Assert rows.Row(2).CellRef("REASON_RAW") = "'Лист1'!E15", "upper merged source metadata retained"
    Set rows = context.ResolveField("ADDRESS_ROWS_RAW").Value
    AT_Assert rows.Count = 2 And rows.Row(2).RowId = "21", "lower blank with merged auxiliary cells skipped"
    AT_Assert rows.Row(2).ResolveField("SYSTEM_RAW").Value.Scalar = "test system", "lower auxiliary inherited only on actual rows"
    AT_TestNetworkExtractedDSL context
    On Error Resume Next
    context.AddField "LATER", AT_RuntimeOK(AT_RuntimeScalar("TEXT", "x"))
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo CleanupError
    AT_Assert caught, "returned extracted context immutable"
    ws.Range("C10").Value2 = CVErr(xlErrValue)
    Set result = reader.Read(ws, "RUN3", "2026-10-16", "Europe/Moscow")
    AT_Assert result.State = atEvalOK And result.Context.ResolveField("DB_ACCOUNT_RAW").State = atEvalError, "data cell error retained for subsequent checks without structural rejection"
    ws.Range("C10").Value2 = "changed": ws.Range("C12").NumberFormat = "@": ws.Range("C12").Value2 = "46100"
    book.Date1904 = True
    Set result = reader.Read(ws, "RUN4", "2026-10-16", "Europe/Moscow")
    Set context = result.Context
    AT_Assert context.ResolveField("WORKBOOK_DATE1904").Value.Scalar, "workbook date epoch retained explicitly"
    AT_Assert context.ResolveField("VALIDITY_VALUE2_KIND").Value.Scalar = "TEXT" And context.ResolveField("VALIDITY_SERIAL").Value.ValueState = atValueMissing, "string digits never guessed to be numeric Excel date"
    ws.Range("A1").Value2 = "old template"
    Set result = reader.Read(ws, "", "bad-date", "")
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "TEMPLATE_MISMATCH", "template denial precedes run-context and data extraction errors"
    AT_Assert result.Diagnostics.Count = 1 And result.Diagnostics(1).Message = "Используйте актуальный шаблон заявки.", "one localized rejection for structural mismatch"
    On Error Resume Next
    Set context = result.Context
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo CleanupError
    AT_Assert caught, "failed extraction cannot expose partial context"
    ws.Range("A1").Value2 = "Заявка на сетевую связность (v.3.0)"
    Set result = reader.Read(ws, "", "bad-date", "")
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "CONFIG_CONTEXT_REQUIRED", "invalid explicit run context controlled"
    AT_Assert ws.Range("C11").Value2 = "line1" & vbLf & "line2" And ws.Range("E15:G17").MergeCells, "full request reader preserves values and merges"
Finish:
    book.Close False
    If Not snapshot Is Nothing Then AT_Assert snapshot.ResolveField("DB_ACCOUNT_RAW").Value.Scalar = "db_test", "whole request snapshot detached from edits and closed source"
    Exit Sub
CleanupError:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise number, , description
End Sub
Private Sub AT_TestNetworkExtractedDSL(ByVal context As ATEvaluationContext)
    Dim schema As New ATCompileContext, compiler As New ATCompiler, evaluator As New ATEvaluator, result As ATEvalResult
    Dim entry As ATTraceEntry, located As Boolean
    schema.AddField "TECHNICAL_ROWS_RAW", "COLLECTION<ROW<TECHNICAL_TABLE_RAW>>"
    schema.AddField "PORT_RAW", "TEXT", "TECHNICAL_TABLE_RAW": schema.AddField "DB_ACCOUNT_RAW", "TEXT"
    Set result = evaluator.Evaluate(compiler.Compile("COUNT(@TECHNICAL_ROWS_RAW) = 2 AND ALLROWS(@TECHNICAL_ROWS_RAW AS t, LEN(@t.PORT_RAW) > 0 AND @DB_ACCOUNT_RAW = ""db_test"")", schema), context)
    AT_Assert result.State = atEvalOK, "compiled DSL consumes unified extraction" & AT_RuntimeTestDetails(result)
    If result.State = atEvalOK Then AT_Assert result.Value.Scalar, "real extracted row and document predicates true"
    For Each entry In result.Trace
        If entry.FieldId = "DB_ACCOUNT_RAW" And entry.CellRef = "'Лист1'!C10" Then located = True
    Next entry
    AT_Assert located, "unified extraction supplies actual cell trace to runtime"
End Sub
Public Function AT_VerifyNetworkReadReference(ByVal workbookName As String) As String
    Dim reader As New ATNetworkRequestReader, result As ATNetworkReadResult, context As ATEvaluationContext
    Set result = reader.Read(Application.Workbooks(workbookName).Worksheets(1), "REFERENCE", "2026-10-16", "Europe/Moscow")
    If result.State <> atEvalOK Then
        AT_VerifyNetworkReadReference = "ERROR|" & result.Diagnostics(1).Code
    Else
        Set context = result.Context
        AT_VerifyNetworkReadReference = "OK|TECH=" & CStr(context.ResolveField("TECHNICAL_ROWS_RAW").Value.Count) & "|ADDR=" & CStr(context.ResolveField("ADDRESS_ROWS_RAW").Value.Count) & "|DB=" & context.CellRef("DB_ACCOUNT_RAW")
    End If
End Function
