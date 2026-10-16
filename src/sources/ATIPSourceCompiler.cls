VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATIPSourceCompiler"
Option Explicit
Public Function Compile(ByVal raw As ATEvalResult, ByVal sourceId As String, ByVal fields As Collection, ByVal fingerprint As String) As ATIPSourceResult
    Dim r As New ATIPSourceResult, f As ATEvalResult, index As New ATIPSourceIndex
    On Error GoTo Failed
    index.Initialize raw, sourceId, fields, fingerprint
    r.Initialize atEvalOK, index, Nothing: Set Compile = r: Exit Function
Failed:
    Set f = AT_IPSourceFailure(Err.Description)
    r.Initialize atEvalError, Nothing, f.Diagnostics: Set Compile = r
End Function
