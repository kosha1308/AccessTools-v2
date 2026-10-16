Attribute VB_Name = "modATRowModelTests"
Option Explicit
Public Sub AT_TestRowModel()
    Dim row As New ATRuntimeRow, rows As New Collection, value As New ATRuntimeValue, copy As ATRuntimeValue
    Dim result As ATEvalResult, caught As Boolean, wrong As New ATRuntimeRow, ignored As Variant
    row.Initialize "T", "12"
    row.AddField "N", AT_RuntimeOK(AT_RuntimeScalar("NUMBER", 2#)), "Sheet1!C12"
    rows.Add row
    value.InitializeRows "T", rows
    rows.Remove 1
    row.AddField "LATER", AT_RuntimeOK(AT_RuntimeScalar("BOOL", True))
    AT_Assert value.DataType = "COLLECTION<ROW<T>>" And value.Count = 1, "typed rows copy input collection"
    Set result = value.Row(1).ResolveField("LATER")
    AT_Assert result.State = atEvalError, "row snapshot isolates later input mutation"
    Set result = value.Row(1).ResolveField("N")
    AT_Assert result.Value.Scalar = 2#, "row scalar round trip"
    AT_Assert value.Row(1).CellRef("N") = "Sheet1!C12", "row cell metadata preserved"
    Set copy = value.Clone
    AT_Assert copy.Row(1).RowId = "12" And copy.RowTableId = "T", "row identity cloned"
    On Error Resume Next
    value.Row(1).AddField "BAD", AT_RuntimeOK(AT_RuntimeScalar("BOOL", True))
    caught = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert caught, "returned row frozen"
    On Error Resume Next
    ignored = value.Scalar
    caught = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert caught, "row collection cannot be scalar"
    wrong.Initialize "OTHER", "12"
    Set rows = New Collection
    rows.Add wrong
    Set value = New ATRuntimeValue
    On Error Resume Next
    value.InitializeRows "T", rows
    caught = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert caught, "wrong table row rejected"
    Set rows = New Collection
    rows.Add row
    rows.Add row
    Set value = New ATRuntimeValue
    On Error Resume Next
    value.InitializeRows "T", rows
    caught = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert caught, "duplicate row identity rejected"
End Sub
