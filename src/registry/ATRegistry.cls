VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mEntries As Collection

Private Sub Class_Initialize()
    Set mEntries = New Collection
End Sub

Public Sub Register(ByVal primitive As ATPrimitive)
    Dim key As String
    Dim signature As ATSignature
    Dim operand As Variant
    If primitive Is Nothing Then Err.Raise vbObjectError + 2130, "ATRegistry", "REG_NULL_PRIMITIVE"
    key = AT_CanonicalId(primitive.Id)
    If Contains(key) Then Err.Raise vbObjectError + 2131, "ATRegistry", "REG_DUPLICATE_ID: " & key
    For Each signature In primitive.Signatures
        If primitive.Category = "DATATYPE" Then
            If signature.ArgumentTypes.Count <> 0 Or signature.ReturnType <> key Then Err.Raise vbObjectError + 2133, "ATRegistry", "REG_INVALID_DATATYPE_SIGNATURE"
        Else
            ValidateType signature.ReturnType
            For Each operand In signature.ArgumentTypes
                ValidateType CStr(operand)
            Next operand
        End If
    Next signature
    primitive.Freeze
    mEntries.Add primitive, key
End Sub

Private Sub ValidateType(ByVal typeName As String)
    If Not Contains(typeName) Then Err.Raise vbObjectError + 2134, "ATRegistry", "REG_UNKNOWN_TYPE: " & typeName
    If GetPrimitive(typeName).Category <> "DATATYPE" Then Err.Raise vbObjectError + 2134, "ATRegistry", "REG_NOT_A_TYPE: " & typeName
End Sub

Public Function Contains(ByVal idValue As String) As Boolean
    Dim entry As ATPrimitive
    Dim key As String
    key = AT_CanonicalId(idValue)
    On Error Resume Next
    Set entry = mEntries(key)
    Contains = Not entry Is Nothing
    On Error GoTo 0
End Function

Public Function GetPrimitive(ByVal idValue As String) As ATPrimitive
    Dim key As String
    key = AT_CanonicalId(idValue)
    If Not Contains(key) Then Err.Raise vbObjectError + 2132, "ATRegistry", "REG_UNKNOWN_ID: " & key
    Set GetPrimitive = mEntries(key)
End Function

Public Property Get Entries() As Collection
    Dim copy As New Collection
    Dim item As ATPrimitive
    For Each item In mEntries
        copy.Add item
    Next item
    Set Entries = copy
End Property

