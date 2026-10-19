Attribute VB_Name = "modAtRuntime"
Option Explicit

Public Function AtEvalValue(ByVal typeName As String, ByVal state As String, ByVal value As String, Optional ByVal dependency As String = "") As CAtEvalResult
    If Not AtAstType(typeName) Or Not AtAstState(state) Then Err.Raise 5, , "Unknown envelope type/state"
    If state = "OK" Then
        If Not AtLiteralValid(typeName, value) Then Err.Raise 5, , "Invalid canonical scalar payload"
    ElseIf value <> "" Then
        Err.Raise 5, , "Non-OK envelope cannot carry scalar payload"
    End If
    Dim r As New CAtEvalResult
    r.SetData typeName, state, value, dependency
    If dependency <> "" Then r.AddProvenance "Input", dependency
    Set AtEvalValue = r
End Function

Public Function AtEvalContext(ByVal packageId As String, ByVal packageVersion As String, ByVal templateId As String, ByVal fields As Object) As CAtEvalContext
    Dim c As New CAtEvalContext
    c.Initialize packageId, packageVersion, templateId, fields
    Set AtEvalContext = c
End Function

' Diagnostic/requirement/provenance/trace records contain scalar strings only.
Public Function AtRecordCopy(ByVal source As Object) As Object
    Dim result As Object, key As Variant: Set result = AtDictionary()
    For Each key In source.Keys
        If VarType(source(key)) <> vbString Then Err.Raise 5, , "Record values must be strings"
        result.Add CStr(key), CStr(source(key))
    Next key
    Set AtRecordCopy = result
End Function

Public Function AtRecordsCopy(ByVal source As Collection) As Collection
    Dim result As New Collection, item As Variant
    For Each item In source: result.Add AtRecordCopy(item): Next item
    Set AtRecordsCopy = result
End Function
