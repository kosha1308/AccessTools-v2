VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkRequestReader"
Option Explicit
Public Function Read(ByVal ws As Worksheet, ByVal runId As String, ByVal runDate As String, ByVal timezoneId As String) As ATNetworkReadResult
    Dim profile As ATTemplateProfile, decision As ATTemplateGateResult
    Dim context As New ATEvaluationContext, result As New ATNetworkReadResult
    Dim spec As ATExcelTableSpec, upper As ATExcelTableLayout, lower As ATExcelTableLayout
    Dim tables As New ATExcelTableReader, fields As New ATExcelFieldReader
    Dim specs As Collection, fieldSpec As ATExcelFieldSpec, field As ATExcelFieldValue
    Dim data As ATEvalResult, raw As Variant, kind As String, code As String
    On Error GoTo Failed
    Set profile = AT_NetworkProfileV3(): Set decision = profile.Verify(ws)
    If Not decision.Allowed Then
        Set Read = Rejected(decision.Code, decision.Message): Exit Function
    End If
    context.Initialize runId, runDate, timezoneId, "INVARIANT", "BINARY": context.Validate
    Set spec = AT_NetworkTechnicalSpecV3(): Set upper = spec.Locate(ws)
    Set spec = AT_NetworkAddressSpecV3(): Set lower = spec.Locate(ws)
    Set data = tables.Read(ws, upper, lower.HeaderFirstRow - 1)
    If data.State <> atEvalOK Then
        result.Initialize atEvalError, Nothing, data.Diagnostics: Set Read = result: Exit Function
    End If
    context.AddField "TECHNICAL_ROWS_RAW", data
    Set data = tables.Read(ws, lower, tables.LastValueRow(ws, lower))
    If data.State <> atEvalOK Then
        result.Initialize atEvalError, Nothing, data.Diagnostics: Set Read = result: Exit Function
    End If
    context.AddField "ADDRESS_ROWS_RAW", data
    Set specs = AT_NetworkDocumentFieldsV3()
    For Each fieldSpec In specs
        Set field = fields.Read(ws, fieldSpec)
        If Not field.Found Then
            Set data = field.Result
            If data.State = atEvalError Then
                result.Initialize atEvalError, Nothing, data.Diagnostics: Set Read = result: Exit Function
            End If
            Set Read = Rejected("NETWORK_FIELD_MISSING", "Не найдено поле заявки."): Exit Function
        End If
        context.AddField field.FieldId, field.Result, field.CellRef
        If field.FieldId = "VALIDITY_RAW" Then
            raw = field.RawValue2
            Select Case VarType(raw)
                Case vbEmpty: kind = "EMPTY"
                Case vbString: kind = "TEXT"
                Case vbBoolean: kind = "BOOL"
                Case vbError: kind = "ERROR"
                Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal: kind = "NUMBER"
                Case Else: Err.Raise vbObjectError + 2921, , "NETWORK_SCALAR_TYPE"
            End Select
            context.AddField "VALIDITY_VALUE2_KIND", AT_RuntimeOK(AT_RuntimeScalar("TEXT", kind)), field.CellRef
            If kind = "NUMBER" Then
                Set data = AT_RuntimeOK(AT_RuntimeScalar("NUMBER", CDbl(raw)))
            ElseIf kind = "ERROR" Then
                Set data = field.Result
            Else
                Set data = AT_RuntimeOK(AT_RuntimeState("NUMBER", atValueMissing))
            End If
            context.AddField "VALIDITY_SERIAL", data, field.CellRef
        End If
    Next fieldSpec
    context.AddField "WORKBOOK_DATE1904", AT_RuntimeOK(AT_RuntimeScalar("BOOL", CBool(ws.Parent.Date1904)))
    result.Initialize atEvalOK, context, Nothing: Set Read = result
    Exit Function
Failed:
    code = Err.Description
    If Left$(code, 6) <> "EXCEL_" And Left$(code, 8) <> "NETWORK_" And Left$(code, 7) <> "CONFIG_" And Left$(code, 9) <> "TEMPLATE_" Then code = "NETWORK_READ_FAILED"
    Set Read = Rejected(code, "Не удалось прочитать заявку.")
End Function
Private Function Rejected(ByVal code As String, ByVal message As String) As ATNetworkReadResult
    Dim diagnostics As New Collection, diagnostic As New ATDiagnostic, result As New ATNetworkReadResult
    diagnostic.Initialize code, "NETWORK", "ERROR", 1, 0, message
    diagnostics.Add diagnostic
    result.Initialize atEvalError, Nothing, diagnostics
    Set Rejected = result
End Function
