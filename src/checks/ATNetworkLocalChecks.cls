VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkLocalChecks"
Option Explicit
Public Function Run(ByVal source As ATNetworkReadResult, ByVal allowedProtocols As Collection) As ATCheckReport
    Dim documents As New ATNetworkDocumentChecks, tables As New ATNetworkTableChecks
    Dim first As ATCheckReport, second As ATCheckReport, third As ATCheckReport, result As New ATCheckReport
    Dim matcher As New ATNetworkAddressMatch
    Dim issues As New Collection, checks As New Collection, issue As ATCheckIssue, outcome As ATCheckOutcome
    Dim failureCode As String
    On Error GoTo Failed
    Set first = documents.Run(source)
    If first.State <> atEvalOK Then
        result.Initialize atEvalError, "NETWORK_LOCAL", first.Issues, first.Checks: Set Run = result: Exit Function
    End If
    Set second = tables.Run(source, allowedProtocols)
    If second.State <> atEvalOK Then
        result.Initialize atEvalError, "NETWORK_LOCAL", second.Issues, second.Checks: Set Run = result: Exit Function
    End If
    Set third = matcher.Run(source)
    If third.State <> atEvalOK Then
        result.Initialize atEvalError, "NETWORK_LOCAL", third.Issues, third.Checks: Set Run = result: Exit Function
    End If
    For Each issue In first.Issues
        issues.Add issue
    Next issue
    For Each issue In second.Issues
        issues.Add issue
    Next issue
    For Each issue In third.Issues
        issues.Add issue
    Next issue
    For Each outcome In first.Checks
        checks.Add outcome
    Next outcome
    For Each outcome In second.Checks
        checks.Add outcome
    Next outcome
    For Each outcome In third.Checks
        checks.Add outcome
    Next outcome
    result.Initialize atEvalOK, "NETWORK_LOCAL", issues, checks: Set Run = result
    Exit Function
Failed:
    failureCode = AT_StableErrorCode(Err.Description, "NETWORK_LOCAL_INPUT_INVALID")
    If failureCode <> "CHECK_REPORT_LIMIT" Then failureCode = "NETWORK_LOCAL_INPUT_INVALID"
    Set result = New ATCheckReport: Set issues = New Collection: Set checks = New Collection
    Set issue = New ATCheckIssue: Set outcome = New ATCheckOutcome
    issue.Initialize "NET.INPUT", failureCode, "ERROR", "", "", "Не удалось собрать полный отчёт: проверьте объём заявки и исходные данные.": issues.Add issue
    outcome.Initialize "NET.INPUT", "BLOCKED", "Общий отчёт не завершён.": checks.Add outcome
    result.Initialize atEvalError, "NETWORK_LOCAL", issues, checks: Set Run = result
End Function
