VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATExcelDateAdapter"
Option Explicit
Public Function Normalize(ByVal raw As Variant, ByVal date1904 As Boolean, ByVal textFormat As String) As ATEvalResult
    Dim serial As Double, text As String, parsed As Date
    On Error GoTo Invalid
    If textFormat <> "ISO" And textFormat <> "DMY_DOT" And textFormat <> "DMY_TEXT" Then
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
            If textFormat = "DMY_TEXT" And Len(raw) > 65536 Then
                Set Normalize = AT_RuntimeFailure(atEvalError, "RUNTIME_LIMIT_TEXT", Nothing): Exit Function
            End If
            text = Trim$(CStr(raw))
            If Len(text) = 0 Then
                Set Normalize = AT_RuntimeOK(AT_RuntimeState("DATE", atValueEmpty)): Exit Function
            End If
            If textFormat = "DMY_DOT" Then
                If Len(text) <> 10 Then GoTo Invalid
                If Not text Like "##.##.####" Then GoTo Invalid
                text = Right$(text, 4) & "-" & Mid$(text, 4, 2) & "-" & Left$(text, 2)
            End If
            If textFormat = "DMY_TEXT" Then
                If Not TryTextDate(text, parsed) Then GoTo Invalid
            Else
                parsed = AT_GregorianDate(text)
            End If
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
Private Function TryTextDate(ByVal text As String, ByRef parsed As Date) As Boolean
    Dim i As Long, candidate As Date
    If TryGregorian(text, parsed) Then
        TryTextDate = True: Exit Function
    End If
    For i = 1 To Len(text)
        If TryDateAt(text, i, candidate) Then
            parsed = candidate: TryTextDate = True
        End If
    Next i
End Function
Private Function TryDateAt(ByVal text As String, ByVal start As Long, ByRef parsed As Date) As Boolean
    Dim position As Long, dayText As String, monthText As String, yearText As String
    If start > 1 Then
        If IsDigit(Mid$(text, start - 1, 1)) Then Exit Function
    End If
    position = start
    dayText = TakeDigits(text, position)
    If Len(dayText) = 0 Then Exit Function
    If Not TakeSeparator(text, position) Then Exit Function
    monthText = TakeDigits(text, position)
    If Len(monthText) = 0 Then Exit Function
    If Not TakeSeparator(text, position) Then Exit Function
    yearText = Mid$(text, position, 4)
    If Len(yearText) <> 4 Then Exit Function
    If Not AT_Digits(yearText, 4) Then Exit Function
    If CLng(yearText) < 1900 Then Exit Function
    If IsDigit(Mid$(text, position + 4, 1)) Then Exit Function
    TryDateAt = TryGregorian(yearText & "-" & Right$("0" & monthText, 2) & "-" & Right$("0" & dayText, 2), parsed)
End Function
Private Function TakeDigits(ByVal text As String, ByRef position As Long) As String
    Dim i As Long, ch As String
    For i = 1 To 2
        ch = Mid$(text, position, 1)
        If Not IsDigit(ch) Then Exit For
        TakeDigits = TakeDigits & ch: position = position + 1
    Next i
End Function
Private Function TakeSeparator(ByVal text As String, ByRef position As Long) As Boolean
    Select Case Mid$(text, position, 1)
        Case ".", "/", "-": position = position + 1: TakeSeparator = True
    End Select
End Function
Private Function IsDigit(ByVal ch As String) As Boolean
    If Len(ch) = 1 Then IsDigit = (ch >= "0" And ch <= "9")
End Function
Private Function TryGregorian(ByVal text As String, ByRef parsed As Date) As Boolean
    On Error GoTo Invalid
    parsed = AT_GregorianDate(text): TryGregorian = True
Invalid:
End Function
