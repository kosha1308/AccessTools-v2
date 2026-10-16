Attribute VB_Name = "modATExcelText"
Option Explicit

Public Function AT_ExcelBlank(ByVal raw As Variant) As Boolean
    If IsError(raw) Then Exit Function
    If IsEmpty(raw) Or IsNull(raw) Then
        AT_ExcelBlank = True
    ElseIf VarType(raw) = vbString Then
        AT_ExcelBlank = Len(AT_ExcelHeader(CStr(raw))) = 0
    End If
End Function

Public Function AT_ExcelHeader(ByVal text As String) As String
    Dim item As Variant
    For Each item In Array(vbTab, vbCr, vbLf, ChrW(160))
        text = Replace(text, CStr(item), " ")
    Next item
    For Each item In Array(ChrW(&H200B), ChrW(&H200C), ChrW(&H200D), ChrW(&HFEFF))
        text = Replace(text, CStr(item), "")
    Next item
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    AT_ExcelHeader = LCase$(Trim$(text))
End Function

Public Function AT_ExcelText(ByVal raw As Variant) As String
    If IsEmpty(raw) Or IsNull(raw) Then Exit Function
    Select Case VarType(raw)
        Case vbString: AT_ExcelText = CStr(raw)
        Case vbBoolean
            If raw Then AT_ExcelText = "TRUE" Else AT_ExcelText = "FALSE"
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            AT_ExcelText = Trim$(Str$(raw))
        Case Else: Err.Raise vbObjectError + 2800, , "EXCEL_CELL_TYPE"
    End Select
End Function

Public Function AT_ExcelBlock(ByVal area As Range) As Variant
    Dim data As Variant, oneCell(1 To 1, 1 To 1) As Variant
    data = area.Value2
    If area.CountLarge = 1 Then
        oneCell(1, 1) = data
        AT_ExcelBlock = oneCell
    Else
        AT_ExcelBlock = data
    End If
End Function

Public Function AT_ExcelCellReference(ByVal sheetName As String, ByVal rowNumber As Long, ByVal columnNumber As Long) As String
    Dim letters As String, remainder As Long
    Do While columnNumber > 0
        remainder = (columnNumber - 1) Mod 26
        letters = Chr$(65 + remainder) & letters
        columnNumber = (columnNumber - 1) \ 26
    Loop
    AT_ExcelCellReference = "'" & Replace(sheetName, "'", "''") & "'!" & letters & CStr(rowNumber)
End Function
