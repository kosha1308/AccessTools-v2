VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATSignature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mArguments As Collection
Private mReturnType As String
Private mFrozen As Boolean

Public Sub Initialize(ByVal arguments As String, ByVal resultType As String)
    Dim item As Variant
    If mFrozen Then Err.Raise vbObjectError + 2113, "ATSignature", "REG_FROZEN_SIGNATURE"
    Set mArguments = New Collection
    If Len(arguments) > 0 Then
        For Each item In Split(arguments, ",")
            If Len(Trim$(CStr(item))) = 0 Then Err.Raise vbObjectError + 2110, "ATSignature", "REG_EMPTY_ARGUMENT_TYPE"
            mArguments.Add UCase$(Trim$(CStr(item)))
        Next item
    End If
    If Len(resultType) = 0 Then Err.Raise vbObjectError + 2111, "ATSignature", "REG_EMPTY_RETURN_TYPE"
    mReturnType = UCase$(resultType)
End Sub

Public Sub Freeze()
    If mArguments Is Nothing Or Len(mReturnType) = 0 Then Err.Raise vbObjectError + 2112, "ATSignature", "REG_UNINITIALIZED_SIGNATURE"
    mFrozen = True
End Sub

Public Property Get ArgumentTypes() As Collection
    Dim copy As New Collection
    Dim item As Variant
    If mArguments Is Nothing Then Err.Raise vbObjectError + 2112, "ATSignature", "REG_UNINITIALIZED_SIGNATURE"
    For Each item In mArguments
        copy.Add item
    Next item
    Set ArgumentTypes = copy
End Property

Public Property Get ReturnType() As String
    ReturnType = mReturnType
End Property

