Attribute VB_Name = "modATEvalLimitTests"
Option Explicit
Public Sub AT_TestEvalLimits()
    Dim limits As New ATEvalLimits, result As ATEvalResult
    limits.MaxSteps = 1
    Set result = AT_EvalFixture("TRUE", , , limits)
    AT_Assert result.State = atEvalOK, "runtime exact step boundary"
    Set result = AT_EvalFixture("NOT TRUE", , , limits)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_LIMIT_STEPS", "runtime exceeded steps"
    limits.MaxSteps = 200000
    limits.MaxDepth = 1
    Set result = AT_EvalFixture("NOT TRUE", , , limits)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_LIMIT_DEPTH", "runtime depth limit"
    limits.MaxDepth = 64
    limits.MaxTraceEntries = 1
    Set result = AT_EvalFixture("@X AND TRUE", AT_RuntimeFailure(atEvalManualReview, "NEEDS_REVIEW", Nothing), , limits)
    AT_Assert result.State = atEvalManualReview, "trace overflow preserves primary state"
    AT_Assert result.Diagnostics(1).Code = "NEEDS_REVIEW", "trace overflow preserves primary code"
    AT_Assert result.Diagnostics.Count = 2, "trace overflow retained as secondary diagnostic"
    Set result = AT_EvalFixture("NOT TRUE", , , limits)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_LIMIT_TRACE", "runtime trace limit"
    limits.MaxTraceEntries = 20000
    limits.MaxTextLength = 1
    Set result = AT_EvalFixture("""ab"" = ""ab""", , , limits)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_LIMIT_TEXT", "runtime text limit"
    limits.MaxTextLength = 65536
    limits.MaxListItems = 1
    Set result = AT_EvalFixture("1 IN [1,2]", , , limits)
    AT_Assert result.Diagnostics(1).Code = "RUNTIME_LIMIT_LIST", "runtime list limit"
    limits.MaxDepth = 129
    Set result = AT_EvalFixture("TRUE", , , limits)
    AT_Assert result.Diagnostics(1).Code = "LIMIT_INVALID_CONFIGURATION", "runtime unsafe settings rejected"
End Sub
