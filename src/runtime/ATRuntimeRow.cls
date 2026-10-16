VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATRuntimeRow"
Option Explicit
Private mTable As String, mRow As String, mInitialized As Boolean, mFrozen As Boolean
Private mFields As Collection, mIds As Collection, mRefs As Collection
Private Sub Class_Initialize()
    Set mFields = New Collection: Set mIds = New Collection: Set mRefs = New Collection
End Sub
Public Sub Initialize(ByVal tableId As String, ByVal rowId As String)
    If mInitialized Then Err.Raise vbObjectError + 2500, , "RUNTIME_VALUE_FROZEN"
    If Len(Trim$(rowId)) = 0 Then Err.Raise vbObjectError + 2700, , "RUNTIME_ROW_ID_REQUIRED"
    mTable = AT_CanonicalId(tableId): mRow = rowId: mInitialized = True
End Sub
Public Sub AddField(ByVal id As String, ByVal result As ATEvalResult, Optional ByVal cellRef As String = "")
    If mFrozen Then Err.Raise vbObjectError + 2500, , "RUNTIME_VALUE_FROZEN"
    If Not mInitialized Then Err.Raise vbObjectError + 2700, , "RUNTIME_ROW_ID_REQUIRED"
    id = AT_CanonicalId(id)
    If HasField(id) Then Err.Raise vbObjectError + 2503, , "RUNTIME_DUPLICATE_FIELD"
    mFields.Add result.Clone, id: mIds.Add id: mRefs.Add cellRef, id
End Sub
Private Function HasField(ByVal id As String) As Boolean
    Dim result As ATEvalResult
    On Error Resume Next
    Set result = mFields(id)
    HasField = Not result Is Nothing
    On Error GoTo 0
End Function
Public Function ResolveField(ByVal id As String, Optional ByVal budget As ATEvalBudget = Nothing) As ATEvalResult
    If Not HasField(id) Then
        Set ResolveField = AT_RuntimeFailure(atEvalError, "RUNTIME_ROW_FIELD_MISSING", Nothing)
    Else
        If Not budget Is Nothing Then mFields(id).ValidateBudget budget
        Set ResolveField = mFields(id).Clone
    End If
End Function
Public Function CellRef(ByVal id As String) As String
    If HasField(id) Then CellRef = mRefs(id)
End Function
Public Function Snapshot() As ATRuntimeRow
    Dim result As New ATRuntimeRow, id As Variant
    If mFrozen Then
        Set Snapshot = Me
        Exit Function
    End If
    If Not mInitialized Then Err.Raise vbObjectError + 2700, , "RUNTIME_ROW_ID_REQUIRED"
    result.Initialize mTable, mRow
    For Each id In mIds
        result.Retain CStr(id), mFields(CStr(id)), CStr(mRefs(CStr(id)))
    Next id
    result.Freeze
    Set Snapshot = result
End Function
Friend Sub Retain(ByVal id As String, ByVal result As ATEvalResult, ByVal cellRef As String)
    If mFrozen Then Err.Raise vbObjectError + 2500, , "RUNTIME_VALUE_FROZEN"
    mFields.Add result, id: mIds.Add id: mRefs.Add cellRef, id
End Sub
Friend Sub Freeze()
    mFrozen = True
End Sub
Public Property Get TableId() As String
    TableId = mTable
End Property
Public Property Get RowId() As String
    RowId = mRow
End Property
