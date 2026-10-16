Attribute VB_Name = "modATRuntimeModelTests"
Option Explicit

Public Sub AT_TestRuntimeModel()
    Dim value As ATRuntimeValue, copy As ATRuntimeValue, item As ATRuntimeValue
    Dim result As ATEvalResult, context As New ATEvaluationContext, snapshot As ATEvaluationContext
    Dim items As New Collection, state As Long, caught As Boolean
    Set value = AT_RuntimeScalar("NUMBER", 42#)
    AT_Assert value.DataType = "NUMBER" And value.Scalar = 42#, "runtime typed scalar"
    For state = atValueEmpty To atValueNull
        Set value = AT_RuntimeState("TEXT", state)
        AT_Assert value.ValueState = state, "runtime state distinct"
        caught = False
        On Error Resume Next
        Dim ignored As Variant
        ignored = value.Scalar
        caught = (Err.Number <> 0)
        Err.Clear
        On Error GoTo 0
        AT_Assert caught, "unusable scalar access guarded"
    Next state
    items.Add AT_RuntimeScalar("NUMBER", 1#)
    Set value = New ATRuntimeValue
    value.InitializeList "NUMBER", items
    items.Remove 1
    AT_Assert value.Count = 1, "runtime list copies input"
    Set copy = value.Clone
    AT_Assert copy.Item(1).Scalar = 1#, "runtime list clone"
    Set result = AT_RuntimeFailure(atEvalSkipped, "TEST_SKIP", Nothing)
    On Error Resume Next
    Set item = result.Value
    caught = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    AT_Assert caught, "non-OK result blocks value"
    context.Initialize "RUN1", "2026-10-16", "UTC", "INVARIANT", "BINARY"
    context.AddField "A", AT_RuntimeOK(value)
    Set snapshot = context.Snapshot
    context.AddField "B", AT_RuntimeOK(AT_RuntimeScalar("BOOL", True))
    Set result = snapshot.ResolveField("B")
    AT_Assert result.State = atEvalError, "context snapshot isolates additions"
    Set result = snapshot.ResolveField("A")
    AT_Assert result.Value.Count = 1, "context value cloned"
    On Error Resume Next
    context.AddField "A", result
    caught = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    AT_Assert caught, "duplicate fields rejected"
    On Error Resume Next
    Set value = AT_RuntimeScalar("BOOL", "TRUE")
    caught = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    AT_Assert caught, "runtime Boolean never coerces text"
    On Error Resume Next
    Set value = AT_RuntimeScalar("DATE", CDate(1.5))
    caught = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    AT_Assert caught, "date-only invariant enforced by public constructor"
End Sub
