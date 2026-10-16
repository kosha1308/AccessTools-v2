VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATTraceEntry"
Option Explicit
Private mSource As String, mPrimitive As String, mStatus As String, mValueState As String, mReason As String
Private mStart As Long, mLength As Long, mOrdinal As Long, mFrozen As Boolean

Public Sub Initialize(ByVal sourceId As String, ByVal start As Long, ByVal length As Long, ByVal primitiveId As String, ByVal status As String, ByVal valueState As String, ByVal reason As String, ByVal ordinal As Long)
    If mFrozen Then Err.Raise vbObjectError + 2500, , "RUNTIME_VALUE_FROZEN"
    mSource = sourceId: mStart = start: mLength = length: mPrimitive = primitiveId
    mStatus = status: mValueState = valueState: mReason = reason: mOrdinal = ordinal
    mFrozen = True
End Sub
Public Property Get SourceId() As String
    SourceId = mSource
End Property
Public Property Get Start() As Long
    Start = mStart
End Property
Public Property Get Length() As Long
    Length = mLength
End Property
Public Property Get PrimitiveId() As String
    PrimitiveId = mPrimitive
End Property
Public Property Get Status() As String
    Status = mStatus
End Property
Public Property Get ValueState() As String
    ValueState = mValueState
End Property
Public Property Get Reason() As String
    Reason = mReason
End Property
Public Property Get Ordinal() As Long
    Ordinal = mOrdinal
End Property
