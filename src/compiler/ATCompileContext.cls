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

Private Sub Class_Initialize()
    Set Registry = AT_CreateRegistry()
    Set mFields = New Collection
    Set mSources = New Collection
    Set mNamed = New Collection
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

Public Sub AddSource(ByVal idValue As String, ByVal typeValue As String)
    Dim source As New ATFieldDefinition
    source.Id = AT_CanonicalId(idValue)
    source.TypeName = UCase$(typeValue)
    mSources.Add source, source.Id
End Sub

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
    Next field
    For Each named In mNamed
        Registry.ValidateConcreteType named.ReturnType
        If named.Status <> "ACTIVE" And named.Status <> "DEPRECATED" Then Err.Raise vbObjectError + 2301, "ATCompileContext", "NAMED_INVALID_STATUS"
        For Each parameter In named.Parameters
            Registry.ValidateConcreteType parameter.TypeName
        Next parameter
    Next named
End Sub

