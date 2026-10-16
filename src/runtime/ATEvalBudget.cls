VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATEvalBudget"
Option Explicit
Private mSteps As Long, mDepth As Long
Private mMaxSteps As Long, mMaxDepth As Long, mMaxList As Long, mMaxText As Long, mMaxTrace As Long
Private mTrace As Collection
Public Sub Initialize(ByVal limits As ATEvalLimits)
    limits.Validate
    mMaxSteps = limits.MaxSteps: mMaxDepth = limits.MaxDepth: mMaxList = limits.MaxListItems
    mMaxText = limits.MaxTextLength: mMaxTrace = limits.MaxTraceEntries
    mSteps = 0: mDepth = 0
    Set mTrace = New Collection
End Sub
Public Sub Step(Optional ByVal units As Long = 1)
    If units < 0 Or units > mMaxSteps - mSteps Then Err.Raise vbObjectError + 2521, , "RUNTIME_LIMIT_STEPS"
    mSteps = mSteps + units
End Sub
Public Sub EnterNode()
    If mDepth >= mMaxDepth Then Err.Raise vbObjectError + 2522, , "RUNTIME_LIMIT_DEPTH"
    mDepth = mDepth + 1
End Sub
Public Sub LeaveNode()
    mDepth = mDepth - 1
End Sub
Public Sub CheckList(ByVal count As Long)
    If count > mMaxList Or count < 0 Then Err.Raise vbObjectError + 2523, , "RUNTIME_LIMIT_LIST"
End Sub
Public Sub CheckText(ByVal length As Long)
    If length > mMaxText Or length < 0 Then Err.Raise vbObjectError + 2524, , "RUNTIME_LIMIT_TEXT"
End Sub
Public Sub AddTrace(ByVal entry As ATTraceEntry)
    If mTrace.Count >= mMaxTrace Then Err.Raise vbObjectError + 2525, , "RUNTIME_LIMIT_TRACE"
    mTrace.Add entry
End Sub
Public Property Get Trace() As Collection
    Dim copy As New Collection, entry As ATTraceEntry
    For Each entry In mTrace
        copy.Add entry
    Next entry
    Set Trace = copy
End Property
Public Property Get TraceCount() As Long
    TraceCount = mTrace.Count
End Property
