VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATLimits"
Option Explicit

Public MaxInputLength As Long
Public MaxTokens As Long
Public MaxAstNodes As Long
Public MaxNesting As Long
Public MaxCollectionNesting As Long
Public MaxNamedDepth As Long
Public MaxCompileSteps As Long

Private Sub Class_Initialize()
    MaxInputLength = 65536
    MaxTokens = 16384
    MaxAstNodes = 8192
    MaxNesting = 64
    MaxCollectionNesting = 8
    MaxNamedDepth = 64
    MaxCompileSteps = 200000
End Sub

Public Sub Validate()
    If MaxInputLength < 1 Or MaxTokens < 1 Or MaxAstNodes < 1 Or MaxNesting < 1 Or MaxCollectionNesting < 1 Or MaxNamedDepth < 1 Or MaxCompileSteps < 1 Then Err.Raise vbObjectError + 2400, "ATLimits", "LIMIT_INVALID_CONFIGURATION"
    If MaxNesting > 128 Or MaxNamedDepth > 128 Or MaxCollectionNesting > 32 Then Err.Raise vbObjectError + 2400, "ATLimits", "LIMIT_UNSAFE_RECURSION_CONFIGURATION"
End Sub
