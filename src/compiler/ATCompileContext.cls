VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATCompileContext"
Option Explicit

Public Registry As ATRegistry
Private mFields As Collection
Private mSources As Collection
Private mNamed As Collection
Private mSourceColumns As Collection

Private Sub Class_Initialize()
    Set Registry = AT_CreateRegistry()
    Set mFields = New Collection
    Set mSources = New Collection
    Set mNamed = New Collection
    Set mSourceColumns = New Collection
End Sub

Public Sub AddField(ByVal idValue As String, ByVal typeValue As String, Optional ByVal tableValue As String = "")
    Dim field As New ATFieldDefinition, key As String
    field.Id = AT_CanonicalId(idValue)
    field.TypeName = UCase$(typeValue)
    If Len(tableValue) > 0 Then field.TableId = AT_CanonicalId(tableValue)
    key = field.TableId & ":" & field.Id
    mFields.Add field, key
End Sub

Public Function FindField(ByVal idValue As String, Optional ByVal tableValue As String = "") As ATFieldDefinition
    On Error Resume Next
    Set FindField = mFields(UCase$(tableValue) & ":" & UCase$(idValue))
    On Error GoTo 0
End Function

Public Sub AddSource(ByVal idValue As String, ByVal typeValue As String, Optional ByVal lookupKeyType As String = "")
    Dim source As New ATFieldDefinition
    source.Id = AT_CanonicalId(idValue)
    source.TypeName = UCase$(typeValue)
    source.LookupKeyType = UCase$(lookupKeyType)
    mSources.Add source, source.Id
End Sub

Public Sub AddSourceColumn(ByVal sourceId As String, ByVal columnId As String, ByVal typeName As String)
    Dim column As New ATFieldDefinition, source As ATFieldDefinition
    Set source = FindSource(sourceId)
    If source Is Nothing Then Err.Raise vbObjectError + 2302, "ATCompileContext", "RESOLVE_UNKNOWN_SOURCE: " & sourceId
    column.Id = AT_CanonicalId(columnId)
    column.TableId = source.Id
    column.TypeName = UCase$(typeName)
    mSourceColumns.Add column, source.Id & ":" & column.Id
End Sub

Public Function FindSourceColumn(ByVal sourceId As String, ByVal columnId As String) As ATFieldDefinition
    On Error Resume Next
    Set FindSourceColumn = mSourceColumns(AT_CanonicalId(sourceId) & ":" & AT_CanonicalId(columnId))
    On Error GoTo 0
End Function

Public Function FindSource(ByVal idValue As String) As ATFieldDefinition
    On Error Resume Next
    Set FindSource = mSources(AT_CanonicalId(idValue))
    On Error GoTo 0
End Function

Public Sub AddNamed(ByVal expression As ATNamedExpression)
    If expression Is Nothing Then Err.Raise vbObjectError + 2300, "ATCompileContext", "CTX_NULL_NAMED"
    mNamed.Add expression, AT_CanonicalId(expression.Id)
End Sub

Public Function FindNamed(ByVal idValue As String) As ATNamedExpression
    On Error Resume Next
    Set FindNamed = mNamed(AT_CanonicalId(idValue))
    On Error GoTo 0
End Function

Public Sub ValidateSchema()
    Dim field As ATFieldDefinition, named As ATNamedExpression, parameter As ATFieldDefinition
    For Each field In mFields
        Registry.ValidateConcreteType field.TypeName
    Next field
    For Each field In mSources
        Registry.ValidateConcreteType field.TypeName
        If Len(field.LookupKeyType) > 0 Then Registry.ValidateConcreteType field.LookupKeyType
    Next field
    For Each field In mSourceColumns
        Registry.ValidateConcreteType field.TypeName
    Next field
    For Each named In mNamed
        Registry.ValidateConcreteType named.ReturnType
        If named.Status <> "ACTIVE" And named.Status <> "DEPRECATED" Then Err.Raise vbObjectError + 2301, "ATCompileContext", "NAMED_INVALID_STATUS"
        If named.Scope <> "DOCUMENT" And named.Scope <> "ROW" Then Err.Raise vbObjectError + 2303, "ATCompileContext", "NAMED_INVALID_SCOPE"
        For Each parameter In named.Parameters
            Registry.ValidateConcreteType parameter.TypeName
        Next parameter
    Next named
End Sub

