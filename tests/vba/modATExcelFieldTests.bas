Attribute VB_Name = "modATExcelFieldTests"
Option Explicit

Public Sub AT_TestExcelFields()
    Dim book As Workbook, ws As Worksheet, reader As New ATExcelFieldReader
    Dim spec As ATExcelFieldSpec, value As ATExcelFieldValue, saved As ATExcelFieldValue
    Dim raw As ATEvalResult, direction As Variant, expected As Variant, item As Variant, i As Long
    Dim savedNumber As Long, savedText As String
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = book.Worksheets(1): ws.Name = "AT's form"
    ws.Range("D4:F5").Merge: ws.Range("D4").Value2 = " Email" & ChrW(160) & vbLf & "address "
    ws.Range("G4:H5").Merge: ws.Range("G4").Value2 = "right"
    ws.Range("B4:C5").Merge: ws.Range("B4").Value2 = "left"
    ws.Range("D6:F7").Merge: ws.Range("D6").Value2 = "below"
    ws.Range("D2:F3").Merge: ws.Range("D2").Value2 = "above"
    direction = Array(Array(0, 1), Array(0, -1), Array(1, 0), Array(-1, 0))
    expected = Array(Array("right", "G4"), Array("left", "B4"), Array("below", "D6"), Array("above", "D2"))
    For i = 0 To 3
        item = direction(i)
        Set spec = AT_TestExcelFieldSpec("EMAIL_RAW", "email address", 1, 10, 1, 10, CLng(item(0)), CLng(item(1)))
        Set value = reader.Read(ws, spec)
        item = expected(i)
        AT_Assert value.Found And value.Result.State = atEvalOK, "field found with merged boundary direction" & CStr(i)
        AT_Assert value.Result.Value.Scalar = item(0), "field offset from normalized merge boundary" & CStr(i)
        AT_Assert value.CellRef = "'AT''s form'!" & CStr(item(1)), "field target merged anchor address" & CStr(i)
        AT_Assert value.AnchorRef = "'AT''s form'!D4" And value.FieldId = "EMAIL_RAW", "field identity and source anchor"
    Next i
    ws.Range("H7").Value2 = "exact offset"
    Set spec = AT_TestExcelFieldSpec("EXACT_RAW", "email address", 4, 5, 4, 6, 2, 2)
    Set value = reader.Read(ws, spec)
    AT_Assert value.Result.Value.Scalar = "exact offset" And value.CellRef = "'AT''s form'!H7", "diagonal target outside anchor search region"
    Set saved = value
    ws.Range("H7").Value2 = "edited"
    AT_Assert saved.Result.Value.Scalar = "exact offset", "field snapshot detached from source edit"
    Set spec = AT_TestExcelFieldSpec("MISSING_RAW", "absent label", 1, 10, 1, 10, 0, 1)
    Set value = reader.Read(ws, spec)
    AT_Assert Not value.Found And value.Result.State = atEvalOK, "missing anchor is controlled value state"
    AT_Assert value.Result.Value.ValueState = atValueMissing And value.CellRef = "", "missing anchor distinct from blank value"
    ws.Range("J10").Value2 = "EMAIL ADDRESS"
    Set spec = AT_TestExcelFieldSpec("EMAIL_RAW", "email address", 1, 10, 1, 10, 0, 1)
    Set value = reader.Read(ws, spec)
    AT_Assert Not value.Found And value.Result.State = atEvalError, "duplicate normalized anchors are not guessed"
    AT_Assert value.Result.Diagnostics(1).Code = "EXCEL_ANCHOR_AMBIGUOUS", "duplicate anchor stable code"
    ws.Range("J10").ClearContents
    ws.Range("G4").Formula = "="""""
    Set value = reader.Read(ws, spec)
    AT_Assert value.Found And value.Result.Value.ValueState = atValueEmpty, "empty formula found but EMPTY"
    ws.Range("G4").Value2 = ChrW(160) & vbTab & ChrW(&H200B)
    Set value = reader.Read(ws, spec)
    AT_Assert value.Found And value.Result.Value.ValueState = atValueEmpty, "visual blanks found but EMPTY"
    ws.Range("G4").Value2 = CVErr(xlErrValue)
    Set value = reader.Read(ws, spec)
    AT_Assert value.Found And value.Result.State = atEvalError, "Excel cell error retains found location"
    AT_Assert IsError(value.RawValue2), "native Excel error retained alongside controlled result"
    AT_Assert value.Result.Diagnostics(1).Code = "EXCEL_CELL_ERROR" And value.CellRef = "'AT''s form'!G4", "Excel cell error stable code and location"
    Set raw = value.Result: raw.Diagnostics.Remove 1
    AT_Assert value.Result.Diagnostics.Count = 1, "field result diagnostic getter isolated"
    ws.Range("G4").Value2 = 0
    Set value = reader.Read(ws, spec)
    AT_Assert value.Result.Value.Scalar = "0", "numeric zero not treated as blank field"
    AT_Assert VarType(value.RawValue2) = vbDouble And value.RawValue2 = 0#, "native numeric zero distinct from raw TEXT"
    ws.Range("G4").Value2 = "broken-prefix 192.0.2.1" & vbLf & "raw suffix"
    Set value = reader.Read(ws, spec)
    AT_Assert value.Result.Value.Scalar = "broken-prefix 192.0.2.1" & vbLf & "raw suffix", "field raw multiline content retained whole"
    ws.Range("G4").Formula = "=""user value"""
    ws.Range("G4").Interior.Color = RGB(10, 20, 30)
    Set value = reader.Read(ws, spec)
    AT_Assert ws.Range("G4").Formula = "=""user value""" And ws.Range("G4:H5").MergeCells, "field read preserves formula and merge"
    AT_Assert ws.Range("G4").Interior.Color = RGB(10, 20, 30), "field read preserves source formatting"
    Set spec = AT_TestExcelFieldSpec("BOUNDS_RAW", "email address", 4, 4, 4, 6, 1, 0)
    Set value = reader.Read(ws, spec)
    AT_Assert value.Result.Diagnostics(1).Code = "EXCEL_MERGE_BOUNDS", "anchor merge cannot escape search bounds"
    Set spec = AT_TestExcelFieldSpec("BOUNDS_RAW", "email address", 4, 5, 5, 6, 0, 1)
    Set value = reader.Read(ws, spec)
    AT_Assert value.Result.State = atEvalError, "left edge inside anchor merge is not missing"
    If value.Result.State = atEvalError Then AT_Assert value.Result.Diagnostics(1).Code = "EXCEL_MERGE_BOUNDS", "left partial anchor merge stable error"
    Set spec = AT_TestExcelFieldSpec("BOUNDS_RAW", "email address", 5, 5, 4, 6, 0, 1)
    Set value = reader.Read(ws, spec)
    AT_Assert value.Result.State = atEvalError, "top edge inside anchor merge is not missing"
    If value.Result.State = atEvalError Then AT_Assert value.Result.Diagnostics(1).Code = "EXCEL_MERGE_BOUNDS", "top partial anchor merge stable error"
    Set spec = AT_TestExcelFieldSpec("OTHER_RAW", "unrelated label", 5, 5, 4, 6, 0, 1)
    Set value = reader.Read(ws, spec)
    AT_Assert value.Result.State = atEvalOK And value.Result.Value.ValueState = atValueMissing, "unrelated partial merge does not supply an absent anchor"
    ws.Range("A20").Value2 = "Edge"
    Set spec = AT_TestExcelFieldSpec("EDGE_RAW", "Edge", 20, 20, 1, 1, 0, -1)
    Set value = reader.Read(ws, spec)
    AT_Assert value.Result.Diagnostics(1).Code = "EXCEL_FIELD_TARGET_BOUNDS", "target before first column rejected"
    ws.Cells(1048576, 1).Value2 = "Bottom"
    Set spec = AT_TestExcelFieldSpec("EDGE_RAW", "Bottom", 1048576, 1048576, 1, 1, 1, 0)
    Set value = reader.Read(ws, spec)
    AT_Assert value.Result.Diagnostics(1).Code = "EXCEL_FIELD_TARGET_BOUNDS", "target after last row rejected"
    AT_TestExcelBadFieldSpecs reader, ws
    book.Close False
    AT_Assert saved.SheetName = "AT's form" And saved.Result.Value.Scalar = "exact offset", "field snapshot survives closing workbook"
    AT_Assert saved.RawValue2 = "exact offset", "native scalar detached from later edits and source closure"
    Exit Sub
CleanupError:
    savedNumber = Err.Number: savedText = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise savedNumber, , savedText
End Sub

Public Function AT_TestExcelFieldSpec(ByVal id As String, ByVal label As String, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal rowOffset As Long, ByVal colOffset As Long) As ATExcelFieldSpec
    Dim spec As New ATExcelFieldSpec
    spec.Initialize id, label, firstRow, lastRow, firstCol, lastCol, rowOffset, colOffset
    Set AT_TestExcelFieldSpec = spec
End Function

Private Sub AT_TestExcelBadFieldSpecs(ByVal reader As ATExcelFieldReader, ByVal ws As Worksheet)
    Dim item As Variant, spec As ATExcelFieldSpec, caught As Boolean, value As ATExcelFieldValue
    For Each item In Array(Array(0, 1, 1, 10, 0, 1), Array(1, 100000, 1, 64, 0, 1), Array(1, 10, 1, 10, 0, 0), Array(1, 10, 1, 10, 0, 16384))
        On Error Resume Next
        Set spec = AT_TestExcelFieldSpec("BAD", "Label", CLng(item(0)), CLng(item(1)), CLng(item(2)), CLng(item(3)), CLng(item(4)), CLng(item(5)))
        caught = Err.Number <> 0: Err.Clear
        On Error GoTo 0
        AT_Assert caught, "field spec rejects unsafe bounds or offsets"
    Next item
    Set spec = New ATExcelFieldSpec
    Set value = reader.Read(ws, spec)
    AT_Assert value.Result.State = atEvalError, "uninitialized spec controlled error"
    Set value = reader.Read(ws, Nothing)
    AT_Assert value.Result.State = atEvalError, "missing spec controlled error"
    ws.Range("A15000:A15001").Merge
    Set spec = AT_TestExcelFieldSpec("LIMIT_RAW", "absent", 100, 20100, 1, 2, 0, 1)
    Set value = reader.Read(ws, spec)
    AT_Assert value.Result.State = atEvalError, "large mixed merge edge bounded before full scan"
    If value.Result.State = atEvalError Then AT_Assert value.Result.Diagnostics(1).Code = "EXCEL_FIELD_MERGE_LIMIT", "merge metadata budget stable code"
    ws.Range("A15000:A15001").UnMerge
    Set value = reader.Read(ws, spec)
    AT_Assert value.Result.State = atEvalOK And value.Result.Value.ValueState = atValueMissing, "reader reuse and unmerged long edges avoid cell probes"
    On Error Resume Next
    spec.Initialize "NEW", "Other", 1, 1, 1, 1, 0, 1
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught And spec.FieldId = "LIMIT_RAW", "field spec cannot be changed after initialization"
End Sub
