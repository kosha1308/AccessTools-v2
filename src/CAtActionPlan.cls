VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtActionPlan"
Option Explicit

Private mBook As Object, mRun As String, mStatus As String, mActions As Collection
Private mCode As String, mMessage As String, mSession As String

Friend Sub Initialize(ByVal book As Object, ByVal runId As String, ByVal session As String, ByVal actions As Collection, ByVal status As String, ByVal code As String, ByVal message As String)
    Set mBook = book: mRun = runId: mSession = session: Set mActions = AtRecordsCopy(actions)
    mStatus = status: mCode = code: mMessage = message
End Sub
Public Property Get RunId() As String
    RunId = mRun
End Property
Public Property Get Status() As String
    Status = mStatus
End Property
Public Property Get Code() As String
    Code = mCode
End Property
Public Property Get Message() As String
    Message = mMessage
End Property
Public Property Get Actions() As Collection
    Set Actions = AtRecordsCopy(mActions)
End Property
Friend Property Get Workbook() As Object
    Set Workbook = mBook
End Property
Friend Property Get SessionId() As String
    SessionId = mSession
End Property
