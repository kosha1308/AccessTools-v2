VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATExcelDateAdapter"
Option Explicit
Public Function Normalize(ByVal raw As Variant, ByVal date1904 As Boolean, ByVal textFormat As String) As ATEvalResult
    Dim serial As Double, text As String, parsed As Date
    On Error GoTo Invalid
    If textFormat <> "ISO" And textFormat <> "DMY_DOT" Then
        Set Normalize = AT_RuntimeFailure(atEvalError, "CONFIG_DATE_FORMAT", Nothing): Exit Function
    End If
    If IsObject(raw) Or IsArray(raw) Then GoTo Invalid
    Select Case VarType(raw)
        Case vbEmpty
            Set Normalize = AT_RuntimeOK(AT_RuntimeState("DATE", atValueEmpty)): Exit Function
        Case vbNull
            Set Normalize = AT_RuntimeOK(AT_RuntimeState("DATE", atValueMissing)): Exit Function
        Case vbError
            Set Normalize = AT_RuntimeFailure(atEvalError, "EXCEL_CELL_ERROR", Nothing): Exit Function
        Case vbString
            text = Trim$(CStr(raw))
            If Len(text) = 0 Then
                Set Normalize = AT_RuntimeOK(AT_RuntimeState("DATE", atValueEmpty)): Exit Function
            End If
            If textFormat = "DMY_DOT" Then
                If Len(text) <> 10 Then GoTo Invalid
                If Not text Like "##.##.####" Then GoTo Invalid
                text = Right$(text, 4) & "-" & Mid$(text, 4, 2) & "-" & Left$(text, 2)
            End If
            parsed = AT_GregorianDate(text)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            serial = CDbl(raw)
            If serial <> Fix(serial) Then GoTo Invalid
            If date1904 Then
                If serial < 0# Or serial > 2957003# Then GoTo Invalid
                serial = serial + 1462#
            Else
                If serial < 1# Or serial > 2958465# Or serial = 60# Then GoTo Invalid
                ' Excel has a fictitious 1900-02-29; VBA/OLE does not.
                If serial < 60# Then serial = serial + 1#
            End If
            parsed = CDate(serial)
        Case Else: GoTo Invalid
    End Select
    Set Normalize = AT_RuntimeOK(AT_RuntimeScalar("DATE", parsed))
    Exit Function
Invalid:
    Set Normalize = AT_RuntimeInvalid("DATE", "EXCEL_INVALID_DATE", Nothing)
End Function
