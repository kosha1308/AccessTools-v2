Attribute VB_Name = "modATPrimitiveEvalTests"
Option Explicit
Public Sub AT_TestPrimitiveEval()
    Dim expression As Variant, result As ATEvalResult, incoming As ATEvalResult
    For Each expression In Array("1 + 2 * 3 = 7", "-5.5 % 2 = -1.5", "5.5 % 2 = 1.5", "ROUND(2.5,0) = 2", "ROUND(3.5,0) = 4", "ABS(-2) = 2", "MIN(2,3) = 2", "MAX([1,3,2]) = 3", "SUM([1,2,3]) = 6", "AVG([1,2,3]) = 2", "2 BETWEEN 2 AND 3", "NOT (2 BETWEEN 3 AND 1)", "DATE_DIFF(DATE(""2024-02-28""),DATE(""2024-03-01"")) = 2", "TODAY() = DATE(""2026-10-16"")", "LEN(""abc"") = 3", "LOWER(""ABC"") = ""abc""", "UPPER(""abc"") = ""ABC""", "TRIM("" abc "") = ""abc""", """abc"" STARTS_WITH ""ab""", """abc"" ENDS_WITH ""bc""", """abc"" CONTAINS ""b""", "1 IN [1,2]", "[1,1] EXACT [1]", "[1] COVER [1,2]", "[1,2] CONTAINS [1]", "[1] ANY [1,2]", "[1] NONE [2]", "[1] OVERLAPS [1,2]", "[1,2] != [2,1]")
        Set result = AT_EvalFixture(CStr(expression))
        AT_Assert result.State = atEvalOK, "primitive state: " & CStr(expression) & AT_RuntimeTestDetails(result)
        If result.State = atEvalOK Then AT_Assert result.Value.Scalar = True, "primitive result: " & CStr(expression)
    Next expression
    For Each expression In Array("1 / 0 = 1", "1 % 0 = 1", "ROUND(1,29) = 1", "ROUND(1,0.5) = 1", """a"" MATCHES "".*""", """a"" LIKE ""*""")
        Set result = AT_EvalFixture(CStr(expression))
        AT_Assert result.State = atEvalError, "primitive error: " & CStr(expression)
    Next expression
    Set result = AT_EvalFixture("@X AND TRUE", AT_RuntimeFailure(atEvalManualReview, "NEEDS_REVIEW", Nothing))
    AT_Assert result.State = atEvalManualReview And result.Diagnostics(1).Code = "NEEDS_REVIEW", "manual review primary code retained"
End Sub

Public Function AT_RuntimeTestDetails(ByVal result As ATEvalResult) As String
    Dim diagnostic As ATDiagnostic
    For Each diagnostic In result.Diagnostics
        AT_RuntimeTestDetails = AT_RuntimeTestDetails & " [" & diagnostic.Code & ": " & diagnostic.Message & "]"
    Next diagnostic
End Function
