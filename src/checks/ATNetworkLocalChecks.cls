VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkLocalChecks"
Option Explicit
Public Function Run(ByVal source As ATNetworkReadResult, ByVal allowedProtocols As Collection) As ATCheckReport
    Dim documents As New ATNetworkDocumentChecks, tables As New ATNetworkTableChecks
    Dim first As ATCheckReport, second As ATCheckReport, result As New ATCheckReport
    Dim issues As New Collection, checks As New Collection, issue As ATCheckIssue, outcome As ATCheckOutcome
    Set first = documents.Run(source)
    If first.State <> atEvalOK Then
        result.Initialize atEvalError, "NETWORK_LOCAL", first.Issues, first.Checks: Set Run = result: Exit Function
    End If
    Set second = tables.Run(source, allowedProtocols)
    If second.State <> atEvalOK Then
        result.Initialize atEvalError, "NETWORK_LOCAL", second.Issues, second.Checks: Set Run = result: Exit Function
    End If
    For Each issue In first.Issues
        issues.Add issue
    Next issue
    For Each issue In second.Issues
        issues.Add issue
    Next issue
    For Each outcome In first.Checks
        checks.Add outcome
    Next outcome
    For Each outcome In second.Checks
        checks.Add outcome
    Next outcome
    result.Initialize atEvalOK, "NETWORK_LOCAL", issues, checks: Set Run = result
End Function
