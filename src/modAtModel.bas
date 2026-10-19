Attribute VB_Name = "modAtModel"
Option Explicit

Public Function AtDictionary() As Object
    Set AtDictionary = CreateObject("Scripting.Dictionary")
    AtDictionary.CompareMode = vbBinaryCompare
End Function

Public Function AtNewPackage(ByVal id As String, ByVal name As String) As CAtPackage
    Dim p As New CAtPackage
    p.Id = id: p.DisplayName = name
    Set AtNewPackage = p
End Function

Public Function AtNewDefinition(ByVal kind As String, ByVal id As String, ByVal name As String) As CAtDefinition
    Dim d As New CAtDefinition
    d.Kind = kind: d.Id = id: d.DisplayName = name
    Set AtNewDefinition = d
End Function

Public Sub AtAddReference(ByVal d As CAtDefinition, ByVal role As String, ByVal kind As String, ByVal id As String, ByVal version As String, Optional ByVal expectedType As String = "")
    Dim r As Object: Set r = AtDictionary()
    r.Add "Role", role: r.Add "Kind", kind: r.Add "Id", id
    r.Add "Version", version: r.Add "ExpectedType", expectedType
    d.References.Add r
End Sub

Public Sub AtProblem(ByVal diagnostics As Collection, ByVal code As String, ByVal objectId As String, ByVal member As String, ByVal reason As String, ByVal message As String, Optional ByVal severity As String = "Error")
    Dim d As Object: Set d = AtDictionary()
    d.Add "Code", code: d.Add "ObjectId", objectId: d.Add "Property", member
    d.Add "Reason", reason: d.Add "Message", message: d.Add "Severity", severity
    diagnostics.Add d
End Sub

Public Function AtHasErrors(ByVal diagnostics As Collection) As Boolean
    Dim d As Variant
    For Each d In diagnostics
        If d("Severity") = "Error" Then AtHasErrors = True: Exit Function
    Next d
End Function

Public Function AtValue(ByVal d As CAtDefinition, ByVal key As String) As String
    If d.Values.Exists(key) Then AtValue = CStr(d.Values(key))
End Function
