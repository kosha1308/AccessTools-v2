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
