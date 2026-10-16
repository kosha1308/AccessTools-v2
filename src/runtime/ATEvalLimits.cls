VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATEvalLimits"
Option Explicit
Public MaxSteps As Long, MaxDepth As Long, MaxListItems As Long, MaxTextLength As Long, MaxTraceEntries As Long
Private Sub Class_Initialize()
    MaxSteps = 200000: MaxDepth = 64: MaxListItems = 10000: MaxTextLength = 65536: MaxTraceEntries = 20000
End Sub
Public Sub Validate()
    If MaxSteps < 1 Or MaxDepth < 1 Or MaxDepth > 128 Or MaxListItems < 1 Or MaxTextLength < 1 Or MaxTraceEntries < 1 Then Err.Raise vbObjectError + 2520, , "LIMIT_INVALID_CONFIGURATION"
End Sub
