VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkPolicyDecision"
Option Explicit
Private mReady As Boolean, mIssue As ATCheckIssue, mCoordinator As ATCheckIssue, mKey As String
Private mUser As Boolean, mRequest As Boolean, mSuperset As Boolean, mUserError As Boolean
Friend Sub Initialize(ByVal issue As ATCheckIssue, ByVal coordinator As ATCheckIssue, ByVal key As String, ByVal inUser As Boolean, ByVal inRequest As Boolean, ByVal inSuperset As Boolean, ByVal userError As Boolean)
    If mReady Then Err.Raise vbObjectError + 2960, , "POLICY_FROZEN"
    If issue Is Nothing Or coordinator Is Nothing Then Err.Raise vbObjectError + 2960, , "POLICY_INPUT_INVALID"
    Set mIssue = issue: Set mCoordinator = coordinator: mKey = key
    mUser = inUser: mRequest = inRequest: mSuperset = inSuperset: mUserError = userError: mReady = True
End Sub
Private Sub RequireReady()
    If Not mReady Then Err.Raise vbObjectError + 2960, , "POLICY_ACCESS"
End Sub
Public Property Get Issue() As ATCheckIssue
    RequireReady: Set Issue = mIssue
End Property
Public Property Get CoordinatorIssue() As ATCheckIssue
    RequireReady: Set CoordinatorIssue = mCoordinator
End Property
Public Property Get Key() As String
    RequireReady: Key = mKey
End Property
Public Property Get InUserReport() As Boolean
    RequireReady: InUserReport = mUser
End Property
Public Property Get InRequest() As Boolean
    RequireReady: InRequest = mRequest
End Property
Public Property Get InSuperset() As Boolean
    RequireReady: InSuperset = mSuperset
End Property
Public Property Get UserError() As Boolean
    RequireReady: UserError = mUserError
End Property
