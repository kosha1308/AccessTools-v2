Attribute VB_Name = "modATIntervalTests"
Option Explicit

Public Sub AT_TestIntervals()
    Dim sets(0 To 63) As ATIntervalSet, budget As New ATEvalBudget, limits As New ATEvalLimits
    Dim a As Long, b As Long, bit As Long, value As Long, caught As Boolean
    Dim full As New ATIntervalSet, gap As New ATIntervalSet
    budget.Initialize limits
    For a = 0 To 63
        Set sets(a) = New ATIntervalSet
        sets(a).Initialize 4294967295#
        bit = 1
        For value = 0 To 5
            If (a And bit) <> 0 Then sets(a).Add CDbl(value), CDbl(value), budget
            bit = bit * 2
        Next value
        sets(a).Freeze
    Next a
    For a = 0 To 63
        For b = 0 To 63
            AT_Assert sets(a).Overlaps(sets(b), budget) = ((a And b) <> 0), "interval overlap finite oracle"
            AT_Assert sets(a).IsSubsetOf(sets(b), budget) = ((a And b) = a), "interval coverage finite oracle"
        Next b
    Next a
    full.Initialize 4294967295#
    full.Add 2147483648#, 4294967295#, budget
    full.Add 0#, 2147483647#, budget
    full.Freeze
    AT_Assert full.Count = 1, "adjacent ranges merge across signed boundary"
    AT_Assert full.Item(1).Low = 0# And full.Item(1).High = 4294967295#, "IPv4 endpoints exact Double"
    gap.Initialize 4294967295#
    gap.Add 0#, 4#, budget
    gap.Add 6#, 10#, budget
    gap.Freeze
    AT_Assert gap.Count = 2, "interval gap retained"
    AT_Assert Not full.IsSubsetOf(gap, budget), "partial overlap never full coverage"
    On Error Resume Next
    full.Add 1#, 2#, budget
    caught = Err.Number <> 0
    Err.Clear
    On Error GoTo 0
    AT_Assert caught, "frozen interval payload rejects mutation"
End Sub
