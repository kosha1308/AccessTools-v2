Attribute VB_Name = "modATExcelLocatorTests"
Option Explicit

Public Sub AT_TestExcelLocator()
    Dim book As Workbook, ws As Worksheet, other As Worksheet
    Dim spec As ATExcelTableSpec, layout As ATExcelTableLayout, caught As String
    Dim savedNumber As Long, savedText As String
    On Error GoTo CleanupError
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = book.Worksheets(1): ws.Name = "AT's form"
    Set other = book.Worksheets.Add
    Set spec = AT_TestExcelSpec("T", 1, 20, 1, 8)
    spec.AddColumn "SRC_RAW", "IP source", "INHERITED"
    spec.AddColumn "DST_RAW", "IP destination", "DIRECT"
    ws.Range("D4:E5").Merge: ws.Range("D4").Value2 = "IP source"
    ws.Range("B4:B5").Merge: ws.Range("B4").Value2 = "  IP" & vbLf & "destination" & ChrW(160)
    Set layout = spec.Locate(ws)
    AT_Assert layout.HeaderFirstRow = 4 And layout.HeaderLastRow = 5, "locator merged header bottom"
    AT_Assert layout.ColumnNumber(1) = 4 And layout.ColumnNumber(2) = 2, "locator reordered columns"
    AT_Assert layout.TableId = "T" And layout.FieldId(1) = "SRC_RAW", "locator column identity"
    AT_Assert layout.FirstColumn = 2 And layout.LastColumn = 5, "locator includes horizontal merged span"
    On Error Resume Next
    layout.ValidateSheet other
    caught = Err.Description: Err.Clear
    On Error GoTo CleanupError
    AT_Assert caught = "EXCEL_LAYOUT_SHEET_MISMATCH", "locator rejects another sheet"
    spec.AddColumn "EXTRA_RAW", "Extra", "NONE"
    AT_Assert layout.Count = 2 And layout.Usage(1) = "INHERITED", "layout detached from mutable spec"
    AT_Assert AT_TestLocateError(spec, ws) = "EXCEL_HEADER_MISSING", "locator missing required column"
    ws.Range("F4").Value2 = "Extra"
    ws.Range("G4").Value2 = "IP source"
    AT_Assert AT_TestLocateError(spec, ws) = "EXCEL_HEADER_DUPLICATE", "locator rejects duplicate complete header"
    ws.Range("G4").ClearContents
    ws.Range("B10").Value2 = "IP destination": ws.Range("D10").Value2 = "IP source": ws.Range("F10").Value2 = "Extra"
    AT_Assert AT_TestLocateError(spec, ws) = "EXCEL_HEADER_AMBIGUOUS", "locator rejects two complete headers"
    Set spec = AT_TestExcelSpec("T", 1, 4, 1, 8)
    spec.AddColumn "SRC_RAW", "IP source", "INHERITED"
    AT_Assert AT_TestLocateError(spec, ws) = "EXCEL_MERGE_BOUNDS", "header merge cannot escape search bounds"
    AT_TestBadExcelSpecs
    book.Close False
    Exit Sub
CleanupError:
    savedNumber = Err.Number: savedText = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Err.Raise savedNumber, , savedText
End Sub

Public Function AT_TestExcelSpec(ByVal id As String, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long) As ATExcelTableSpec
    Dim spec As New ATExcelTableSpec
    spec.Initialize id, firstRow, lastRow, firstCol, lastCol
    Set AT_TestExcelSpec = spec
End Function

Private Function AT_TestLocateError(ByVal spec As ATExcelTableSpec, ByVal ws As Worksheet) As String
    Dim layout As ATExcelTableLayout
    On Error GoTo Failed
    Set layout = spec.Locate(ws)
    Exit Function
Failed:
    AT_TestLocateError = Err.Description
End Function

Private Sub AT_TestBadExcelSpecs()
    Dim spec As ATExcelTableSpec, caught As Boolean
    On Error Resume Next
    Set spec = AT_TestExcelSpec("T", 0, 10, 1, 8)
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught, "profile rejects row zero"
    On Error Resume Next
    Set spec = AT_TestExcelSpec("T", 1, 100000, 1, 64)
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught, "profile rejects huge header scan before reading"
    Set spec = AT_TestExcelSpec("T", 1, 20, 1, 8)
    spec.AddColumn "A", "Header", "DIRECT"
    On Error Resume Next
    spec.AddColumn "A", "Other", "NONE"
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught, "profile rejects duplicate field id"
    On Error Resume Next
    spec.AddColumn "B", " header ", "NONE"
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught, "profile rejects equivalent header labels"
    On Error Resume Next
    spec.AddColumn "C", "Third", "GUESS"
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught, "profile rejects unknown row participation mode"
End Sub
