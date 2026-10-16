VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATIPSourceCache"
Option Explicit
Private mEntries As Object
Private Sub Class_Initialize()
    Set mEntries = CreateObject("Scripting.Dictionary")
End Sub
Public Function GetOrCompile(ByVal raw As ATEvalResult, ByVal sourceId As String, ByVal fields As Collection, ByVal fingerprint As String) As ATIPSourceResult
    Dim c As New ATIPSourceCompiler, result As ATIPSourceResult, key As String, entry As Variant, keys As Variant, f As ATEvalResult
    On Error GoTo Failed
    key = AT_IPSourceKey(sourceId, fields)
    If mEntries.Exists(key) Then
        entry = mEntries(key)
        If CStr(entry(0)) = fingerprint Then Set GetOrCompile = entry(1): Exit Function
        mEntries.Remove key
    End If
    Set result = c.Compile(raw, sourceId, fields, fingerprint)
    If result.State = atEvalOK Then
        If mEntries.Count >= 4 Then keys = mEntries.Keys: mEntries.Remove CStr(keys(0))
        mEntries.Add key, Array(fingerprint, result)
    End If
    Set GetOrCompile = result: Exit Function
Failed:
    Set result = New ATIPSourceResult: Set f = AT_IPSourceFailure(Err.Description)
    result.Initialize atEvalError, Nothing, f.Diagnostics: Set GetOrCompile = result
End Function
Public Property Get Count() As Long
    Count = mEntries.Count
End Property
