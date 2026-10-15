VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATDiagnostic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Code As String
Public Category As String
Public Severity As String
Public Start As Long
Public Length As Long
Public Message As String

Public Sub Initialize(ByVal codeValue As String, ByVal categoryValue As String, ByVal severityValue As String, ByVal startValue As Long, ByVal lengthValue As Long, ByVal messageValue As String)
    If startValue < 1 Or lengthValue < 0 Then Err.Raise vbObjectError + 2100, "ATDiagnostic", "DIAG_INVALID_SPAN"
    If Len(codeValue) = 0 Or Len(messageValue) = 0 Then Err.Raise vbObjectError + 2101, "ATDiagnostic", "DIAG_REQUIRED_METADATA"
    Select Case severityValue
        Case "ERROR", "WARNING", "INFO"
        Case Else: Err.Raise vbObjectError + 2102, "ATDiagnostic", "DIAG_INVALID_SEVERITY"
    End Select
    Code = codeValue
    Category = categoryValue
    Severity = severityValue
    Start = startValue
    Length = lengthValue
    Message = messageValue
End Sub

