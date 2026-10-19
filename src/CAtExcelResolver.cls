VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtExcelResolver"
Option Explicit

Private mPackage As CAtPackage, mTemplate As String, mIndex As Object, mOffsets As Object
Private mRemaining As Long
Private Const MAX_CELLS As Long = 100000

Friend Sub Initialize(ByVal package As CAtPackage, ByVal templateId As String, ByVal columnOffsets As Object, ByVal catalog As Object)
    If Not mPackage Is Nothing Then Err.Raise 5, , "Resolver configuration is immutable"
    Dim issues As Collection, d As CAtDefinition, key As Variant, registry As New CAtOperationRegistry
    If catalog Is Nothing Then Set catalog = registry.Catalog
    Set mPackage = AtFromXml(AtToXml(package), catalog, issues)
    If mPackage Is Nothing Then Err.Raise 5, , "Invalid Resolver package"
    If AtHasErrors(issues) Then Set mPackage = Nothing: Err.Raise 5, , "Invalid Resolver package"
    Set mIndex = AtDictionary(): Set mOffsets = AtDictionary()
    For Each d In mPackage.Definitions: mIndex.Add d.Kind & "/" & d.Id, d: Next d
    If Not mIndex.Exists("Template/" & templateId) Then Err.Raise 5, , "Unknown template"
    mTemplate = templateId
    If columnOffsets Is Nothing Then Err.Raise 5, , "Explicit column binding dictionary required (may be empty)"
    For Each key In columnOffsets.Keys
        If Not mIndex.Exists("Column/" & CStr(key)) Then Err.Raise 5, , "Unknown binding ColumnId"
        If VarType(columnOffsets(key)) <> vbLong And VarType(columnOffsets(key)) <> vbInteger Then Err.Raise 5, , "Column offset must be an integer"
        If columnOffsets(key) < 1 Or columnOffsets(key) > 16384 Then Err.Raise 5, , "Column offset outside Excel grid"
        mOffsets.Add CStr(key), CLng(columnOffsets(key))
    Next key
End Sub

' ByRef diagnostics describes fatal snapshot failure; no partial context escapes.
Public Function Resolve(ByVal workbook As Object, ByRef diagnostics As Collection) As CAtEvalContext
    Dim fields As Object, d As CAtDefinition, r As CAtEvalResult
    On Error GoTo Broken
    Set diagnostics = New Collection: Set fields = AtDictionary(): mRemaining = MAX_CELLS
    For Each d In mPackage.Definitions
        If d.Kind = "Field" Then
            If RefId(d, "OwnerTemplate") = mTemplate Then
                Set r = FieldValue(workbook, d.Id): fields.Add d.Id, r
            End If
        End If
    Next d
    Set Resolve = AtEvalContext(mPackage.Id, mPackage.Version, mTemplate, fields)
    Exit Function
Broken:
    If InStr(1, Err.Description, "budget", vbTextCompare) > 0 Then
        AtProblem diagnostics, "R_CONTEXT_LIMIT", mTemplate, "Context", "LIMIT", "Immutable-context resource cap exceeded; no partial context returned."
    Else
        AtProblem diagnostics, "R_CONTEXT", mTemplate, "Context", "ERROR", "No context returned: invalid host/configuration."
    End If
End Function

Public Function ReadField(ByVal workbook As Object, ByVal fieldId As String) As CAtEvalResult
    mRemaining = MAX_CELLS
    Set ReadField = FieldValue(workbook, fieldId)
End Function

Private Function FieldValue(ByVal workbook As Object, ByVal fieldId As String) As CAtEvalResult
    Dim d As CAtDefinition, locator As CAtDefinition, area As Object, failure As CAtEvalResult, r As CAtEvalResult
    Dim typeName As String, table As CAtTableSnapshot, columnId As String, tableDefinition As CAtDefinition, tableLocator As CAtDefinition
    typeName = "Text"
    On Error GoTo Broken
    If Not mIndex.Exists("Field/" & fieldId) Then
        Set r = AtExcelFailure(typeName, "MISSING", fieldId, "R_FIELD", "Field definition not found."): GoTo Done
    End If
    Set d = mIndex("Field/" & fieldId): typeName = AtValue(d, "DataType")
    If RefId(d, "OwnerTemplate") <> mTemplate Then
        Set r = AtExcelFailure(typeName, "MISSING", fieldId, "R_SCOPE", "Field belongs to another template."): GoTo Done
    End If
    Set locator = mIndex("Locator/" & RefId(d, "Locator"))
    If AtValue(locator, "Strategy") = "TableColumn" Then
        Set tableDefinition = mIndex("Table/" & RefId(locator, "Table"))
        Set tableLocator = mIndex("Locator/" & RefId(tableDefinition, "Locator"))
        If StrComp(AtValue(locator, "Sheet"), AtValue(tableLocator, "Sheet"), vbTextCompare) <> 0 Then
            Set r = AtExcelFailure(typeName, "INVALID", fieldId, "R_TABLE_SHEET", "TableColumn sheet must match the table bounds locator."): GoTo Done
        End If
        Set table = TableValue(workbook, RefId(locator, "Table"))
        If table.Status.State <> "OK" Then
            Set failure = table.Status
        ElseIf table.RowCount <> 1 Then
            Set failure = AtExcelFailure("Text", "CONFLICT", fieldId, "R_CARDINALITY", "Table column is not scalar; no implicit first row or aggregation.")
        Else
            columnId = RefId(locator, "Column")
            Set failure = table.ReadCell(1, columnId)
            If failure.ResultType <> typeName Then
                Set r = AtExcelFailure(typeName, "INVALID", fieldId, "R_TYPE", "Field and table column types differ.")
            Else
                Set r = AtEvalValue(typeName, failure.State, failure.Value, fieldId): r.MergeEffects failure
            End If
            Set failure = Nothing
        End If
    Else
        Set area = Locate(workbook, locator, failure)
        If failure Is Nothing Then Set r = CellValue(area, typeName, fieldId, False)
    End If
    If Not failure Is Nothing Then
        Set r = AtEvalValue(typeName, failure.State, "", fieldId): r.MergeEffects failure
    End If
    r.AddProvenance "Locator", locator.Id & "@" & locator.Version
Done:
    Set FieldValue = r: Exit Function
Broken:
    Set FieldValue = AtExcelFailure(typeName, "UNAVAILABLE", fieldId, "R_ACCESS", "Excel resource could not be read; no fallback value inferred.")
End Function

Public Function ReadTable(ByVal workbook As Object, ByVal tableId As String) As CAtTableSnapshot
    mRemaining = MAX_CELLS
    Set ReadTable = TableValue(workbook, tableId)
End Function

Private Function TableValue(ByVal workbook As Object, ByVal tableId As String) As CAtTableSnapshot
    Dim table As CAtDefinition, locator As CAtDefinition, d As CAtDefinition, area As Object
    Dim columns As New Collection, offsets As Object, failure As CAtEvalResult, snapshot As CAtTableSnapshot
    Dim row As Object, id As Variant, r As CAtEvalResult, i As Long, total As Long, offset As Long, mergeFlag As Variant, status As CAtEvalResult
    On Error GoTo Broken
    Set offsets = AtDictionary()
    If Not mIndex.Exists("Table/" & tableId) Then
        Set failure = AtExcelFailure("Text", "MISSING", tableId, "R_TABLE", "Table definition not found."): GoTo Rejected
    End If
    Set table = mIndex("Table/" & tableId)
    If RefId(table, "OwnerTemplate") <> mTemplate Then
        Set failure = AtExcelFailure("Text", "MISSING", tableId, "R_SCOPE", "Table belongs to another template."): GoTo Rejected
    End If
    Set locator = mIndex("Locator/" & RefId(table, "Locator"))
    If AtValue(locator, "Strategy") = "TableColumn" Then
        Set failure = AtExcelFailure("Text", "INVALID", tableId, "R_TABLE_LOCATOR", "Table bounds cannot recursively refer to TableColumn."): GoTo Rejected
    End If
    Set area = Locate(workbook, locator, failure)
    If Not failure Is Nothing Then GoTo Rejected
    If area.Rows.Count > 1000 Or area.Columns.Count > 256 Or area.Cells.CountLarge > MAX_CELLS Then
        Set failure = AtExcelFailure("Text", "LIMIT", tableId, "R_LIMIT", "Table bounds exceed 1000 rows, 256 columns or 100000 cells."): GoTo Rejected
    End If
    For Each d In mPackage.Definitions
        If d.Kind = "Column" Then
            If RefId(d, "OwnerTable") = tableId Then
                columns.Add d.Id
                If Not mOffsets.Exists(d.Id) Then
                    Set failure = AtExcelFailure("Text", "MISSING", d.Id, "R_COLUMN_BINDING", "Explicit relative column offset is required."): GoTo Rejected
                End If
                offset = mOffsets(d.Id)
                If offset > area.Columns.Count Then
                    Set failure = AtExcelFailure("Text", "MISSING", d.Id, "R_COLUMN_BOUNDS", "Column offset lies outside declared table bounds."): GoTo Rejected
                End If
                If offsets.Exists(CStr(offset)) Then
                    Set failure = AtExcelFailure("Text", "CONFLICT", d.Id, "R_COLUMN_CONFLICT", "Two ColumnIds map to the same physical column."): GoTo Rejected
                End If
                offsets.Add CStr(offset), True
            End If
        End If
    Next d
    If columns.Count = 0 Then
        Set failure = AtExcelFailure("Text", "MISSING", tableId, "R_COLUMNS", "Table has no declared columns."): GoTo Rejected
    End If
    Set snapshot = New CAtTableSnapshot
    Set status = AtEvalValue("Text", "OK", "", tableId)
    mergeFlag = area.MergeCells
    If IsNull(mergeFlag) Then
        Set status = AtExcelFailure("Text", "CONFLICT", tableId, "R_TABLE_MERGE", "Table bounds intersect merged cells, including unmapped columns.")
    ElseIf CBool(mergeFlag) Then
        Set status = AtExcelFailure("Text", "CONFLICT", tableId, "R_TABLE_MERGE", "Table bounds intersect merged cells, including unmapped columns.")
    End If
    snapshot.Initialize tableId, area.Row, columns, status
    For i = 1 To area.Rows.Count
        Set row = AtDictionary()
        For Each id In columns
            Set d = mIndex("Column/" & CStr(id))
            Set r = CellValue(area.Cells(i, mOffsets(id)), AtValue(d, "DataType"), CStr(id) & "/" & CStr(area.Row + i - 1), True)
            total = total + AtExcelEnvelopeSize(r)
            If total > 1048576 Or r.State = "LIMIT" Then
                Set failure = AtExcelFailure("Text", "LIMIT", tableId, "R_LIMIT", "Table snapshot budget exceeded; no partial rows returned."): GoTo Rejected
            End If
            row.Add CStr(id), r
        Next id
        snapshot.AppendRow row
    Next i
    Set TableValue = snapshot: Exit Function
Rejected:
    Set snapshot = New CAtTableSnapshot
    snapshot.Initialize tableId, 0, columns, failure
    Set TableValue = snapshot: Exit Function
Broken:
    Set failure = AtExcelFailure("Text", "UNAVAILABLE", tableId, "R_ACCESS", "Table resource could not be read.")
    Resume Rejected
End Function

Private Function Locate(ByVal workbook As Object, ByVal locator As CAtDefinition, ByRef failure As CAtEvalResult) As Object
    Dim sheet As Object, item As Object, area As Object, name As Object, candidate As Object
    Dim strategy As String, address As String, count As Long, qualified As String, raw As String
    Dim values As Variant, value As Variant, i As Long, j As Long, row As Long, column As Long
    On Error GoTo Broken
    If workbook Is Nothing Then GoTo Broken
    If TypeName(workbook) <> "Workbook" Then GoTo Broken
    If workbook.Worksheets.Count > mRemaining Then GoTo Limited
    mRemaining = mRemaining - workbook.Worksheets.Count
    For Each item In workbook.Worksheets
        If StrComp(item.Name, AtValue(locator, "Sheet"), vbTextCompare) = 0 Then Set sheet = item: Exit For
    Next item
    If sheet Is Nothing Then
        Set failure = AtExcelFailure("Text", "MISSING", locator.Id, "R_SHEET", "Declared worksheet not found."): Exit Function
    End If
    strategy = AtValue(locator, "Strategy")
    Select Case strategy
        Case "A1"
            Set area = sheet.Range(AtValue(locator, "Address"))
        Case "DefinedName"
            If workbook.Names.Count > 2000 Or workbook.Names.Count > mRemaining Then GoTo Limited
            mRemaining = mRemaining - workbook.Names.Count
            For Each name In workbook.Names
                qualified = name.Name
                If StrComp(qualified, AtValue(locator, "DefinedName"), vbTextCompare) = 0 Or StrComp(qualified, sheet.Name & "!" & AtValue(locator, "DefinedName"), vbTextCompare) = 0 Or StrComp(qualified, "'" & Replace(sheet.Name, "'", "''") & "'!" & AtValue(locator, "DefinedName"), vbTextCompare) = 0 Then
                    count = count + 1: Set candidate = name
                End If
            Next name
            If count = 0 Then
                Set failure = AtExcelFailure("Text", "MISSING", locator.Id, "R_NAME", "Defined name not found."): Exit Function
            ElseIf count > 1 Then
                Set failure = AtExcelFailure("Text", "CONFLICT", locator.Id, "R_NAME_CONFLICT", "Workbook and sheet scopes both define the name."): Exit Function
            End If
            raw = candidate.RefersTo
            If InStr(raw, "[") > 0 Or InStr(raw, "]") > 0 Then
                Set failure = AtExcelFailure("Text", "UNAVAILABLE", locator.Id, "R_EXTERNAL_NAME", "External name references are not opened or evaluated."): Exit Function
            ElseIf InStr(raw, "#REF!") > 0 Then
                Set failure = AtExcelFailure("Text", "MISSING", locator.Id, "R_NAME", "Defined name has a broken reference."): Exit Function
            End If
            address = AtExcelNameAddress(raw, sheet.Name)
            If address = "" Then
                Set failure = AtExcelFailure("Text", "INVALID", locator.Id, "R_NAME_REFERENCE", "Name must be a direct rectangle on the declared worksheet; expressions/unions are unsupported."): Exit Function
            End If
            Set area = sheet.Range(address)
        Case "LabelOffset"
            Set area = sheet.UsedRange
            If area.Cells.CountLarge > mRemaining Then GoTo Limited
            mRemaining = mRemaining - CLng(area.Cells.CountLarge)
            values = area.Value2
            For i = 1 To area.Rows.Count
                For j = 1 To area.Columns.Count
                    If area.Cells.CountLarge = 1 Then value = values Else value = values(i, j)
                    If VarType(value) = vbString Then
                        If StrComp(value, AtValue(locator, "Label"), vbBinaryCompare) = 0 Then
                            count = count + 1: Set candidate = area.Cells(i, j).MergeArea.Cells(1, 1)
                            If count > 1 Then
                                Set failure = AtExcelFailure("Text", "CONFLICT", locator.Id, "R_LABEL_CONFLICT", "Label has more than one exact match."): Exit Function
                            End If
                        End If
                    End If
                Next j
            Next i
            If count = 0 Then
                Set failure = AtExcelFailure("Text", "MISSING", locator.Id, "R_LABEL", "Exact label not found in bounded UsedRange."): Exit Function
            End If
            row = candidate.Row + CLng(AtValue(locator, "RowOffset")): column = candidate.Column + CLng(AtValue(locator, "ColumnOffset"))
            If row < 1 Or row > sheet.Rows.Count Or column < 1 Or column > sheet.Columns.Count Then
                Set failure = AtExcelFailure("Text", "MISSING", locator.Id, "R_OFFSET", "Label offset lies outside worksheet."): Exit Function
            End If
            Set area = sheet.Cells(row, column)
        Case Else
            Set failure = AtExcelFailure("Text", "INVALID", locator.Id, "R_LOCATOR", "Unsupported range locator."): Exit Function
    End Select
    Set Locate = area: Exit Function
Limited:
    Set failure = AtExcelFailure("Text", "LIMIT", locator.Id, "R_LIMIT", "Search/name budget exceeded."): Exit Function
Broken:
    Set failure = AtExcelFailure("Text", "UNAVAILABLE", locator.Id, "R_ACCESS", "Explicit workbook resource is unavailable.")
End Function

Private Function CellValue(ByVal area As Object, ByVal typeName As String, ByVal objectId As String, ByVal tableCell As Boolean) As CAtEvalResult
    Dim anchor As Object, merged As Object, r As CAtEvalResult
    If mRemaining < 1 Then
        Set CellValue = AtExcelFailure(typeName, "LIMIT", objectId, "R_LIMIT", "Read budget exceeded."): Exit Function
    End If
    mRemaining = mRemaining - 1
    Set anchor = area.Cells(1, 1): Set merged = anchor.MergeArea
    If tableCell And merged.Cells.CountLarge > 1 Then
        Set r = AtExcelFailure(typeName, "CONFLICT", objectId, "R_TABLE_MERGE", "Merged table cells cannot be silently copied across logical rows or columns.")
    ElseIf area.Cells.CountLarge > 1 And (merged.Cells.CountLarge = 1 Or area.Row < merged.Row Or area.Column < merged.Column Or area.Row + area.Rows.Count > merged.Row + merged.Rows.Count Or area.Column + area.Columns.Count > merged.Column + merged.Columns.Count) Then
        Set r = AtExcelFailure(typeName, "CONFLICT", objectId, "R_CARDINALITY", "Scalar locator resolves multiple cells outside one MergeArea.")
    Else
        Set anchor = merged.Cells(1, 1)
        Set r = AtExcelScalar(anchor.Value2, typeName, objectId)
        If anchor.HasFormula Then r.AddProvenance "FormulaResult", anchor.Parent.Name & "!" & anchor.Address
    End If
    r.AddProvenance "Workbook", area.Parent.Parent.Name
    r.AddProvenance "Requested", area.Parent.Name & "!" & area.Address
    r.AddProvenance "Cell", anchor.Parent.Name & "!" & anchor.Address
    Set CellValue = r
End Function

Private Function RefId(ByVal definition As CAtDefinition, ByVal role As String) As String
    Dim reference As Variant
    For Each reference In definition.References
        If reference("Role") = role Then RefId = CStr(reference("Id")): Exit Function
    Next reference
End Function
