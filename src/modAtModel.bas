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
    d.Add "NodeId", ""
    diagnostics.Add d
End Sub

Public Function AtHasErrors(ByVal diagnostics As Collection) As Boolean
    Dim d As Variant
    For Each d In diagnostics
        If d("Severity") = "Error" Then AtHasErrors = True: Exit Function
    Next d
End Function

Public Function AtNewNode(ByVal id As String, ByVal kind As String, ByVal resultType As String) As CAtNode
    Dim n As New CAtNode
    n.Id = id: n.Kind = kind: n.ResultType = resultType
    Set AtNewNode = n
End Function

Public Function AtTypedParameter(ByVal typeName As String, ByVal value As String) As Object
    Dim p As Object: Set p = AtDictionary()
    p.Add "Type", typeName: p.Add "Value", value
    Set AtTypedParameter = p
End Function

Public Function AtSignature(ByVal resultType As String, ByVal argumentTypes As Collection) As Object
    Dim s As Object, states As New Collection, parameters As Object
    Set s = AtDictionary(): Set parameters = AtDictionary()
    states.Add "OK"
    s.Add "ReturnType", resultType: s.Add "Arguments", argumentTypes
    s.Add "Parameters", parameters: s.Add "States", states: s.Add "Propagation", "Strict"
    Set AtSignature = s
End Function

Public Function AtValue(ByVal d As CAtDefinition, ByVal key As String) As String
    If d.Values.Exists(key) Then AtValue = CStr(d.Values(key))
End Function
