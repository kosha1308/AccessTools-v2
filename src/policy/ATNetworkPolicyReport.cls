VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkPolicyReport"
Option Explicit
Private mReady As Boolean, mReport As ATCheckReport, mDecisions As Collection
Friend Sub Initialize(ByVal report As ATCheckReport, ByVal decisions As Collection)
    Dim item As ATNetworkPolicyDecision
    If mReady Then Err.Raise vbObjectError + 2960, , "POLICY_FROZEN"
    If report Is Nothing Or decisions Is Nothing Then Err.Raise vbObjectError + 2960, , "POLICY_INPUT_INVALID"
    If report.State = atEvalOK And report.Issues.Count <> decisions.Count Then Err.Raise vbObjectError + 2960, , "POLICY_INPUT_INVALID"
    Set mReport = report: Set mDecisions = New Collection
    For Each item In decisions: mDecisions.Add item: Next item
    mReady = True
End Sub
Private Sub RequireReady()
    If Not mReady Then Err.Raise vbObjectError + 2960, , "POLICY_ACCESS"
End Sub
Public Property Get AllReport() As ATCheckReport
    RequireReady: Set AllReport = mReport
End Property
Public Property Get Decisions() As Collection
    Dim copy As New Collection, item As ATNetworkPolicyDecision
    RequireReady
    For Each item In mDecisions: copy.Add item: Next item
    Set Decisions = copy
End Property
Public Property Get UserIssues() As Collection
    Set UserIssues = RoutedIssues("USER")
End Property
Public Property Get RequestIssues() As Collection
    Set RequestIssues = RoutedIssues("REQUEST")
End Property
Public Property Get SupersetIssues() As Collection
    Set SupersetIssues = RoutedIssues("SUPERSET")
End Property
Private Function RoutedIssues(ByVal channel As String) As Collection
    Dim copy As New Collection, item As ATNetworkPolicyDecision
    RequireReady
    If mReport.State <> atEvalOK Then
        If channel = "USER" Then Set copy = mReport.Issues
    Else
        For Each item In mDecisions
            Select Case channel
                Case "USER": If item.InUserReport Then copy.Add item.Issue
                Case "REQUEST": If item.InRequest Then copy.Add item.Issue
                Case "SUPERSET": If item.InSuperset Then copy.Add item.CoordinatorIssue
            End Select
        Next item
    End If
    Set RoutedIssues = copy
End Function
Public Property Get UserErrorCount() As Long
    Dim item As ATNetworkPolicyDecision
    RequireReady
    If mReport.State <> atEvalOK Then
        UserErrorCount = mReport.ErrorCount
    Else
        For Each item In mDecisions
            If item.InUserReport And item.UserError And item.Issue.Severity = "ERROR" Then UserErrorCount = UserErrorCount + 1
        Next item
    End If
End Property
