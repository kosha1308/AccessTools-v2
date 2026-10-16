VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkRequestNormalizer"
Option Explicit
Public Function Normalize(ByVal source As ATNetworkReadResult, ByVal dateTextFormat As String, ByVal unlimitedKeyword As String) As ATNetworkReadResult
    Dim sourceContext As ATEvaluationContext, targetContext As New ATEvaluationContext, result As New ATNetworkReadResult
    Dim dates As New ATExcelDateAdapter, check As ATEvalResult, value As ATEvalResult
    Dim limits As New ATEvalLimits, budget As New ATEvalBudget
    Dim specs As Collection, spec As ATExcelFieldSpec, id As Variant, code As String
    On Error GoTo Failed
    If source Is Nothing Then BadInput
    If source.State <> atEvalOK Then
        result.Initialize atEvalError, Nothing, source.Diagnostics: Set Normalize = result: Exit Function
    End If
    Set check = dates.Normalize(Empty, False, dateTextFormat)
    If check.State <> atEvalOK Then
        result.Initialize atEvalError, Nothing, check.Diagnostics: Set Normalize = result: Exit Function
    End If
    If Len(unlimitedKeyword) > 1024 Then Err.Raise vbObjectError + 2930, , "CONFIG_UNLIMITED_KEYWORD"
    If Len(AT_ExcelHeader(unlimitedKeyword)) = 0 Then Err.Raise vbObjectError + 2930, , "CONFIG_UNLIMITED_KEYWORD"
    If AT_Digits(Trim$(unlimitedKeyword), Len(unlimitedKeyword)) Then Err.Raise vbObjectError + 2930, , "CONFIG_UNLIMITED_KEYWORD"
    Set check = dates.Normalize(unlimitedKeyword, False, dateTextFormat)
    If check.Value.ValueState = atValuePresent Then Err.Raise vbObjectError + 2930, , "CONFIG_UNLIMITED_KEYWORD"
    Set sourceContext = source.Context: sourceContext.Validate: budget.Initialize limits
    targetContext.Initialize sourceContext.RunId, sourceContext.RunDate, sourceContext.TimezoneId, sourceContext.LocaleId, sourceContext.ComparisonMode
    Set specs = AT_NetworkDocumentFieldsV3()
    For Each spec In specs
        Set value = RequiredField(sourceContext, spec.FieldId, "TEXT", budget)
        targetContext.AddField spec.FieldId, value, sourceContext.CellRef(spec.FieldId)
        targetContext.AddField Left$(spec.FieldId, Len(spec.FieldId) - 4), value, sourceContext.CellRef(spec.FieldId)
    Next spec
    For Each id In Array("TECHNICAL_ROWS_RAW", "ADDRESS_ROWS_RAW")
        If CStr(id) = "TECHNICAL_ROWS_RAW" Then code = "TECHNICAL_TABLE_RAW" Else code = "ADDRESS_TABLE_RAW"
        Set value = RequiredField(sourceContext, CStr(id), "COLLECTION<ROW<" & code & ">>", budget)
        If value.State <> atEvalOK Then BadInput
        If value.Value.ValueState <> atValuePresent Then BadInput
        targetContext.AddField CStr(id), value, sourceContext.CellRef(CStr(id))
        If CStr(id) = "TECHNICAL_ROWS_RAW" Then
            targetContext.AddField "TECHNICAL_ROWS", NormalizeRows(value.Value, "TECHNICAL_TABLE", TechnicalFields(), budget)
        Else
            targetContext.AddField "ADDRESS_ROWS", NormalizeRows(value.Value, "ADDRESS_TABLE", AddressFields(), budget)
        End If
    Next id
    Set value = RequiredField(sourceContext, "WORKBOOK_DATE1904", "BOOL", budget)
    If value.State <> atEvalOK Then BadInput
    If value.Value.ValueState <> atValuePresent Then BadInput
    targetContext.AddField "WORKBOOK_DATE1904", value
    Set value = RequiredField(sourceContext, "VALIDITY_VALUE2_KIND", "TEXT", budget)
    If value.State <> atEvalOK Then BadInput
    If value.Value.ValueState <> atValuePresent Then BadInput
    targetContext.AddField "VALIDITY_VALUE2_KIND", value, sourceContext.CellRef("VALIDITY_VALUE2_KIND")
    Set value = RequiredField(sourceContext, "VALIDITY_SERIAL", "NUMBER", budget)
    targetContext.AddField "VALIDITY_SERIAL", value, sourceContext.CellRef("VALIDITY_SERIAL")
    NormalizeValidity sourceContext, targetContext, dateTextFormat, unlimitedKeyword
    result.Initialize atEvalOK, targetContext, Nothing: Set Normalize = result
    Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "NETWORK_INPUT_INVALID")
    If Left$(code, 7) <> "CONFIG_" And Left$(code, 14) <> "RUNTIME_LIMIT_" Then code = "NETWORK_INPUT_INVALID"
    result.Initialize atEvalError, Nothing, AT_RuntimeFailure(atEvalError, code, Nothing).Diagnostics
    Set Normalize = result
End Function
Private Function RequiredField(ByVal context As ATEvaluationContext, ByVal id As String, ByVal expectedType As String, ByVal budget As ATEvalBudget) As ATEvalResult
    Dim value As ATEvalResult
    Set value = context.ResolveField(id)
    ValidateRaw value, expectedType, budget
    Set RequiredField = value
End Function
Private Sub ValidateRaw(ByVal value As ATEvalResult, ByVal expectedType As String, ByVal budget As ATEvalBudget)
    Dim diagnostic As ATDiagnostic
    budget.Step
    If value.State = atEvalOK Then
        If value.Value.DataType <> expectedType Then BadInput
        value.ValidateBudget budget
    Else
        For Each diagnostic In value.Diagnostics
            If diagnostic.Code = "RUNTIME_FIELD_MISSING" Or diagnostic.Code = "RUNTIME_ROW_FIELD_MISSING" Then BadInput
        Next diagnostic
    End If
End Sub
Private Function TechnicalFields() As Variant
    TechnicalFields = Array(Array("SOURCE_RAW", "TECH_SOURCE_IP", "IP_EXPRESSION"), Array("DESTINATION_RAW", "TECH_DEST_IP", "IP_EXPRESSION"), Array("PROTOCOL_RAW", "TECH_PROTOCOL", "PROTOCOL_TEXT"), Array("PORT_RAW", "TECH_DEST_PORT", "PORT_SET"), Array("REASON_RAW", "TECH_REASON", "TEXT"))
End Function
Private Function AddressFields() As Variant
    AddressFields = Array(Array("SYSTEM_RAW", "SYSTEM", "TEXT"), Array("ADDRESS_RAW", "ADDRESS_IP", "IP_EXPRESSION"), Array("HOSTNAME_RAW", "HOSTNAME", "TEXT"), Array("ROLE_RAW", "ROLE", "TEXT"))
End Function
Private Function NormalizeRows(ByVal raw As ATRuntimeValue, ByVal tableId As String, ByVal fields As Variant, ByVal budget As ATEvalBudget) As ATEvalResult
    Dim rows As New Collection, source As ATRuntimeRow, row As ATRuntimeRow
    Dim value As New ATRuntimeValue, field As Variant, data As ATEvalResult, i As Long
    budget.CheckList raw.Count
    For i = 1 To raw.Count
        Set source = raw.Row(i): Set row = New ATRuntimeRow
        row.Initialize tableId, source.RowId
        For Each field In fields
            Set data = source.ResolveField(CStr(field(0)))
            ValidateRaw data, "TEXT", budget
            row.AddField CStr(field(1)), NormalizeTextField(data, CStr(field(2)), budget), source.CellRef(CStr(field(0)))
        Next field
        rows.Add row
    Next i
    value.InitializeRows tableId, rows
    Set NormalizeRows = AT_RuntimeOK(value)
End Function
Private Function NormalizeTextField(ByVal raw As ATEvalResult, ByVal typeName As String, ByVal budget As ATEvalBudget) As ATEvalResult
    Dim converted As ATEvalResult, result As New ATEvalResult, state As ATValueState
    If raw.State <> atEvalOK Or typeName = "TEXT" Then
        Set NormalizeTextField = raw.Clone: Exit Function
    End If
    If typeName = "PROTOCOL_TEXT" Then typeName = "TEXT"
    state = raw.Value.ValueState
    If state <> atValuePresent Then
        result.Initialize atEvalOK, AT_RuntimeState(typeName, state), raw.Diagnostics, raw.Trace
        Set NormalizeTextField = result: Exit Function
    End If
    If typeName = "TEXT" Then
        Set NormalizeTextField = AT_RuntimeOK(AT_RuntimeScalar("TEXT", AT_AsciiCase(Trim$(CStr(raw.Value.Scalar)), True)))
    Else
        Set converted = AT_ParseDomain(typeName, CStr(raw.Value.Scalar), Nothing, budget)
        If converted.State <> atEvalOK Then Err.Raise vbObjectError + 2930, , converted.Diagnostics(1).Code
        Set NormalizeTextField = converted
    End If
End Function
Private Sub NormalizeValidity(ByVal sourceContext As ATEvaluationContext, ByVal targetContext As ATEvaluationContext, ByVal dateTextFormat As String, ByVal unlimitedKeyword As String)
    Dim raw As ATEvalResult, serial As ATEvalResult, dates As New ATExcelDateAdapter
    Dim dateResult As ATEvalResult, unlimited As ATEvalResult, kind As String, epoch As Boolean
    Set raw = sourceContext.ResolveField("VALIDITY_RAW"): Set serial = sourceContext.ResolveField("VALIDITY_SERIAL")
    kind = CStr(sourceContext.ResolveField("VALIDITY_VALUE2_KIND").Value.Scalar)
    epoch = CBool(sourceContext.ResolveField("WORKBOOK_DATE1904").Value.Scalar)
    Set unlimited = AT_RuntimeOK(AT_RuntimeScalar("BOOL", False))
    Select Case kind
        Case "ERROR"
            If raw.State <> atEvalError Or serial.State <> atEvalError Then BadInput
            Set dateResult = raw.Clone: Set unlimited = raw.Clone
        Case "NUMBER"
            If raw.State <> atEvalOK Or serial.State <> atEvalOK Then BadInput
            If raw.Value.ValueState <> atValuePresent Or serial.Value.ValueState <> atValuePresent Then BadInput
            Set dateResult = dates.Normalize(serial.Value.Scalar, epoch, dateTextFormat)
        Case "TEXT", "EMPTY", "BOOL"
            If raw.State <> atEvalOK Or serial.State <> atEvalOK Then BadInput
            If serial.Value.ValueState <> atValueMissing Then BadInput
            If kind = "EMPTY" And raw.Value.ValueState <> atValueEmpty Then BadInput
            If kind = "BOOL" Then
                Set dateResult = AT_RuntimeInvalid("DATE", "EXCEL_INVALID_DATE", Nothing)
            ElseIf raw.Value.ValueState <> atValuePresent Then
                Set dateResult = AT_RuntimeOK(AT_RuntimeState("DATE", raw.Value.ValueState))
            ElseIf AT_ExcelHeader(CStr(raw.Value.Scalar)) = AT_ExcelHeader(unlimitedKeyword) Then
                Set unlimited = AT_RuntimeOK(AT_RuntimeScalar("BOOL", True))
                Set dateResult = AT_RuntimeOK(AT_RuntimeState("DATE", atValueMissing))
            Else
                Set dateResult = dates.Normalize(raw.Value.Scalar, epoch, dateTextFormat)
            End If
        Case Else: BadInput
    End Select
    targetContext.AddField "VALIDITY_DATE", dateResult, sourceContext.CellRef("VALIDITY_RAW")
    targetContext.AddField "VALIDITY_IS_UNLIMITED", unlimited, sourceContext.CellRef("VALIDITY_RAW")
End Sub
Private Sub BadInput()
    Err.Raise vbObjectError + 2930, , "NETWORK_INPUT_INVALID"
End Sub
