Attribute VB_Name = "modATTestMain"
Option Explicit

Private mPassed As Long
Private mFailed As Long
Private mFailures As String

Public Function AT_RunTests(Optional ByVal tracePath As String = "") As String
    mPassed = 0
    mFailed = 0
    mFailures = ""
    On Error GoTo TestFailure
    AT_Assert AT_SCHEMA_VERSION = "1.0.0", "schema version"
    AT_Assert AT_DSL_VERSION = "1.0.0", "DSL version"
    AT_TestTrace tracePath, "registry"
    AT_TestRegistry
    AT_TestTrace tracePath, "lexer"
    AT_TestLexer
    AT_TestTrace tracePath, "parser"
    AT_TestParser
    AT_TestTrace tracePath, "compiler"
    AT_TestCompiler
    AT_TestTrace tracePath, "limits"
    AT_TestLimits
    AT_TestTrace tracePath, "properties"
    AT_TestProperties
    AT_TestTrace tracePath, "appendix"
    AT_TestAppendix
    AT_TestTrace tracePath, "catalog"
    AT_TestCatalog
    AT_TestTrace tracePath, "source-schema"
    AT_TestSourceSchema
    AT_TestTrace tracePath, "diagnostics"
    AT_TestDiagnostics
    AT_TestTrace tracePath, "registry-examples"
    AT_TestRegistryExamples
    AT_TestTrace tracePath, "runtime-model"
    AT_TestRuntimeModel
    AT_TestTrace tracePath, "runtime-normalize"
    AT_TestNormalize
    AT_TestTrace tracePath, "runtime-evaluator"
    AT_TestEvaluator
    AT_TestEvalLimits
    AT_TestPrimitiveEval
    AT_TestRuntimeAcceptance
    AT_TestTrace tracePath, "domain-intervals"
    AT_TestIntervals
    AT_TestDomainParse
    AT_TestDomainEval
    AT_TestTrace tracePath, "row-binding"
    AT_TestRowBinding
    AT_TestRowModel
    AT_TestRowEval
    AT_TestTrace tracePath, "excel-locator"
    AT_TestExcelLocator
    AT_TestTrace tracePath, "excel-reader"
    AT_TestExcelReader
    AT_TestTrace tracePath, "excel-fields"
    AT_TestExcelFields
    AT_TestTrace tracePath, "document-fields"
    AT_TestDocumentFields
    AT_TestTrace tracePath, "superset-sets"
    AT_TestSupersetSets
    AT_TestTrace tracePath, "superset-read"
    AT_TestSupersetRead
    AT_TestTrace tracePath, "template-profile"
    AT_TestTemplateProfiles
    AT_TestTemplateRead
    AT_TestNetworkProfile
    AT_TestTrace tracePath, "complete"
BuildReport:
    AT_RunTests = "{""passed"":" & mPassed & ",""failed"":" & mFailed & _
        ",""failures"":""" & mFailures & """}"
    Exit Function
TestFailure:
    AT_Assert False, "Unhandled VBA error: " & CStr(Err.Number) & " " & Err.Description
    Resume BuildReport
End Function

Private Sub AT_TestTrace(ByVal path As String, ByVal stage As String)
    Dim handle As Integer
    If Len(path) = 0 Then Exit Sub
    handle = FreeFile
    Open path For Append As #handle
    Print #handle, stage
    Close #handle
End Sub

Public Sub AT_Assert(ByVal condition As Boolean, ByVal label As String)
    If condition Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
        label = Replace(Replace(label, "\", "\\"), Chr$(34), "\" & Chr$(34))
        label = Replace(Replace(Replace(label, vbCr, "\r"), vbLf, "\n"), vbTab, "\t")
        mFailures = mFailures & label & "; "
    End If
End Sub
