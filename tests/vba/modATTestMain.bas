Attribute VB_Name = "modATTestMain"
Option Explicit

Private mPassed As Long
Private mFailed As Long
Private mFailures As String

Public Function AT_RunTests() As String
    mPassed = 0
    mFailed = 0
    mFailures = ""
    AT_Assert AT_SCHEMA_VERSION = "1.0.0", "schema version"
    AT_Assert AT_DSL_VERSION = "1.0.0", "DSL version"
    AT_TestRegistry
    AT_TestLexer
    AT_RunTests = "{""passed"":" & mPassed & ",""failed"":" & mFailed & _
        ",""failures"":""" & mFailures & """}"
End Function

Public Sub AT_Assert(ByVal condition As Boolean, ByVal label As String)
    If condition Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
        mFailures = mFailures & Replace(Replace(label, "\", "\\"), Chr$(34), "\" & Chr$(34)) & "; "
    End If
End Sub
