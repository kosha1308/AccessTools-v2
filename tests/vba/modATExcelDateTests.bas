Attribute VB_Name = "modATExcelDateTests"
Option Explicit
Public Sub AT_TestExcelDates()
    Dim adapter As New ATExcelDateAdapter, result As ATEvalResult, item As Variant
    Dim emptyValue As Variant, objectValue As New Collection, dateValue As Date, i As Long
    ' A wrong epoch offset, permissive CDate, or serial-text coercion must break these assertions.
    For Each item In Array(Array(1#, False, "1900-01-01"), Array(59#, False, "1900-02-28"), Array(61#, False, "1900-03-01"), Array(1462#, False, "1904-01-01"), Array(0#, True, "1904-01-01"), Array(59#, True, "1904-02-29"), Array(2958465#, False, "9999-12-31"), Array(2957003#, True, "9999-12-31"))
        Set result = adapter.Normalize(item(0), CBool(item(1)), "ISO")
        AT_Assert result.State = atEvalOK, "Excel date state " & CStr(item(0))
        If result.State = atEvalOK Then
            AT_Assert result.Value.ValueState = atValuePresent And result.Value.DataType = "DATE", "Excel date typed present"
            If result.Value.ValueState = atValuePresent Then AT_Assert CDbl(result.Value.Scalar) = CDbl(AT_GregorianDate(CStr(item(2)))), "Excel date calendar " & CStr(item(2))
        End If
    Next item
    For Each item In Array(-1#, 0#, 60#, 1.5, 2958466#, 1E+100, True, "46100", "1900-02-29", "2025-02-29", "2026-13-01", "2026-01-32", "16/10/2026", "2026-10-16 00:00:00", "2026-1-01", "2026-10-1", "01.02.2026")
        Set result = adapter.Normalize(item, False, "ISO")
        AssertDateState result, atValueInvalid, "invalid ISO/raw date " & CStr(item)
    Next item
    For Each item In Array(-1#, 0.5, 2957004#)
        AssertDateState adapter.Normalize(item, True, "DMY_DOT"), atValueInvalid, "invalid 1904 date"
    Next item
    For Each item In Array(Array("2024-02-29", "ISO", "2024-02-29"), Array(" 2026-10-16 ", "ISO", "2026-10-16"), Array("29.02.2024", "DMY_DOT", "2024-02-29"), Array("01.02.2026", "DMY_DOT", "2026-02-01"), Array(" 16.10.2026 ", "DMY_DOT", "2026-10-16"))
        Set result = adapter.Normalize(item(0), False, CStr(item(1)))
        AssertDateState result, atValuePresent, "explicit date text " & CStr(item(0))
        If result.State = atEvalOK Then
            If result.Value.ValueState = atValuePresent Then AT_Assert result.Value.Scalar = AT_GregorianDate(CStr(item(2))), "unambiguous date components"
        End If
    Next item
    For Each item In Array("29.02.2025", "00.10.2026", "01.00.2026", "1.02.2026", "2026-02-01", "46100", "16.10.26", "16. 10.2026", "16.10.2026junk")
        AssertDateState adapter.Normalize(item, False, "DMY_DOT"), atValueInvalid, "invalid explicit DMY text"
    Next item
    AssertDateState adapter.Normalize(emptyValue, False, "ISO"), atValueEmpty, "native empty date"
    AssertDateState adapter.Normalize("   ", False, "ISO"), atValueEmpty, "empty text date"
    AssertDateState adapter.Normalize(Null, False, "ISO"), atValueMissing, "null missing date"
    Set result = adapter.Normalize(CVErr(xlErrValue), False, "ISO")
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "EXCEL_CELL_ERROR", "Excel CVErr date remains error"
    Set result = adapter.Normalize("2026-10-16", False, "AUTO")
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "CONFIG_DATE_FORMAT", "date policy never guesses locale"
    Set result = adapter.Normalize(emptyValue, False, "")
    AT_Assert result.State = atEvalError And result.Diagnostics(1).Code = "CONFIG_DATE_FORMAT", "invalid format rejected even for blank"
    AssertDateState adapter.Normalize(Array(1), False, "ISO"), atValueInvalid, "array not date"
    AssertDateState adapter.Normalize(objectValue, False, "ISO"), atValueInvalid, "object not date"
    dateValue = AT_GregorianDate("2026-10-16")
    AssertDateState adapter.Normalize(dateValue, False, "ISO"), atValueInvalid, "Date subtype not Value2 serial"
    For i = 61 To 2200 Step 37
        Set result = adapter.Normalize(CDbl(i), False, "ISO")
        AT_Assert CDbl(result.Value.Scalar) = CDbl(i), "post leap-bug serial identity"
        Set result = adapter.Normalize(CDbl(i), True, "ISO")
        AT_Assert CDbl(result.Value.Scalar) = CDbl(i + 1462), "1904 serial offset"
    Next i
End Sub
Private Sub AssertDateState(ByVal result As ATEvalResult, ByVal expected As ATValueState, ByVal label As String)
    AT_Assert result.State = atEvalOK, label & " eval state"
    If result.State = atEvalOK Then AT_Assert result.Value.DataType = "DATE" And result.Value.ValueState = expected, label & " value state"
End Sub
