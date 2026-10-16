Attribute VB_Name = "modATRuntimeValues"
Option Explicit

Public Function AT_RuntimeScalarType(ByVal typeName As String) As Boolean
    Select Case typeName
        Case "TEXT", "NUMBER", "BOOL", "DATE": AT_RuntimeScalarType = True
    End Select
End Function
Public Function AT_RuntimeScalar(ByVal typeName As String, ByVal payload As Variant) As ATRuntimeValue
    Dim result As New ATRuntimeValue
    If typeName = "TEXT" And VarType(payload) = vbString Then
        If Len(payload) = 0 Then
            result.InitializeScalar typeName, atValueEmpty, Empty
            Set AT_RuntimeScalar = result
            Exit Function
        End If
    End If
    result.InitializeScalar typeName, atValuePresent, payload
    Set AT_RuntimeScalar = result
End Function
Public Function AT_RuntimeState(ByVal typeName As String, ByVal state As ATValueState) As ATRuntimeValue
    Dim result As New ATRuntimeValue
    If state = atValuePresent Then Err.Raise vbObjectError + 2502, , "RUNTIME_VALUE_ACCESS"
    result.InitializeScalar typeName, state, Empty
    Set AT_RuntimeState = result
End Function
Public Function AT_RuntimeOK(ByVal value As ATRuntimeValue) As ATEvalResult
    Dim result As New ATEvalResult, diagnostics As New Collection, trace As New Collection
    result.Initialize atEvalOK, value, diagnostics, trace
    Set AT_RuntimeOK = result
End Function
Public Function AT_RuntimeFailure(ByVal state As ATEvalState, ByVal code As String, ByVal node As ATAstNode) As ATEvalResult
    Dim result As New ATEvalResult, diagnostics As New Collection
    diagnostics.Add AT_RuntimeDiagnostic(code, node)
    result.Initialize state, Nothing, diagnostics, Nothing
    Set AT_RuntimeFailure = result
End Function
Public Function AT_RuntimeDiagnostic(ByVal code As String, ByVal node As ATAstNode) As ATDiagnostic
    Dim result As New ATDiagnostic, start As Long, length As Long
    start = 1
    If Not node Is Nothing Then
        start = node.Start
        If start < 1 Then start = 1
        length = node.Length
    End If
    result.Initialize code, "RUNTIME", "ERROR", start, length, code
    If Not node Is Nothing Then result.SourceId = node.SourceId
    Set AT_RuntimeDiagnostic = result
End Function
Public Function AT_CopyDiagnostics(ByVal source As Collection) As Collection
    Dim result As New Collection, item As ATDiagnostic, copy As ATDiagnostic
    If Not source Is Nothing Then
        For Each item In source
            Set copy = New ATDiagnostic
            copy.Initialize item.Code, item.Category, item.Severity, item.Start, item.Length, item.Message
            copy.SourceId = item.SourceId
            result.Add copy
        Next item
    End If
    Set AT_CopyDiagnostics = result
End Function
