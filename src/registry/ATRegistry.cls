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
Private mConversions As Collection

Private Sub Class_Initialize()
    Set mEntries = New Collection
    Set mConversions = New Collection
End Sub

Public Sub Register(ByVal primitive As ATPrimitive)
    Dim key As String
    Dim signature As ATSignature
    Dim operand As Variant
    If primitive Is Nothing Then Err.Raise vbObjectError + 2130, "ATRegistry", "REG_NULL_PRIMITIVE"
    key = AT_CanonicalId(primitive.Id)
    If Contains(key) Then Err.Raise vbObjectError + 2131, "ATRegistry", "REG_DUPLICATE_ID: " & key
    If HasSyntax(primitive.Syntax, primitive.Category) Then Err.Raise vbObjectError + 2136, "ATRegistry", "REG_AMBIGUOUS_SYNTAX: " & primitive.Syntax
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

Private Sub ValidateType(ByVal typeName As String, Optional ByVal depth As Long = 0)
    If depth > 16 Then Err.Raise vbObjectError + 2137, "ATRegistry", "TYPE_METADATA_DEPTH"
    If typeName = "T" Then Exit Sub
    If Left$(typeName, 5) = "LIST<" And Right$(typeName, 1) = ">" Then
        ValidateType Mid$(typeName, 6, Len(typeName) - 6), depth + 1
        Exit Sub
    End If
    If Left$(typeName, 4) = "SET<" And Right$(typeName, 1) = ">" Then
        ValidateType Mid$(typeName, 5, Len(typeName) - 5), depth + 1
        Exit Sub
    End If
    If typeName = "COLLECTION<ROW<T>>" Then Exit Sub
    If Not Contains(typeName) Then Err.Raise vbObjectError + 2134, "ATRegistry", "REG_UNKNOWN_TYPE: " & typeName
    If GetPrimitive(typeName).Category <> "DATATYPE" Then Err.Raise vbObjectError + 2134, "ATRegistry", "REG_NOT_A_TYPE: " & typeName
End Sub

Public Sub AllowConversion(ByVal fromType As String, ByVal toType As String)
    ValidateType fromType
    ValidateType toType
    mConversions.Add True, fromType & "->" & toType
End Sub

Public Function CanConvert(ByVal fromType As String, ByVal toType As String) As Boolean
    Dim found As Boolean
    If fromType = toType Then
        CanConvert = True
        Exit Function
    End If
    On Error Resume Next
    found = mConversions(fromType & "->" & toType)
    CanConvert = found
    On Error GoTo 0
End Function

Public Function FindSyntax(ByVal syntaxValue As String, ByVal categoryValue As String) As ATPrimitive
    Dim entry As ATPrimitive
    For Each entry In mEntries
        If entry.Syntax = syntaxValue And entry.Category = categoryValue Then
            Set FindSyntax = entry
            Exit Function
        End If
    Next entry
    Err.Raise vbObjectError + 2135, "ATRegistry", "REG_UNKNOWN_SYNTAX: " & syntaxValue
End Function

Public Function HasSyntax(ByVal syntaxValue As String, ByVal categoryValue As String) As Boolean
    Dim entry As ATPrimitive
    For Each entry In mEntries
        If entry.Syntax = syntaxValue And entry.Category = categoryValue Then
            HasSyntax = True
            Exit Function
        End If
    Next entry
End Function

Public Function MatchSyntax(ByVal tokens As Collection, ByVal index As Long, ByVal categoryValue As String) As ATPrimitive
    Dim entry As ATPrimitive, pieces As Variant, offset As Long
    Dim matched As Boolean, longest As Long
    For Each entry In mEntries
        If entry.Category = categoryValue Then
            pieces = Split(entry.Syntax, " ")
            matched = True
            If index + UBound(pieces) > tokens.Count Then matched = False
            If matched Then
                For offset = 0 To UBound(pieces)
                    If tokens(index + offset).Kind <> "IDENT" And tokens(index + offset).Kind <> "SYMBOL" Then matched = False
                    If tokens(index + offset).Value <> pieces(offset) Then matched = False
                Next offset
            End If
            If matched And UBound(pieces) + 1 > longest Then
                longest = UBound(pieces) + 1
                Set MatchSyntax = entry
            End If
        End If
    Next entry
End Function

Public Sub ValidateConcreteType(ByVal typeName As String, Optional ByVal depth As Long = 0)
    Dim inner As String
    If depth > 16 Then Err.Raise vbObjectError + 2137, "ATRegistry", "TYPE_METADATA_DEPTH"
    If Left$(typeName, 5) = "LIST<" And Right$(typeName, 1) = ">" Then
        ValidateConcreteType Mid$(typeName, 6, Len(typeName) - 6), depth + 1
    ElseIf Left$(typeName, 4) = "SET<" And Right$(typeName, 1) = ">" Then
        ValidateConcreteType Mid$(typeName, 5, Len(typeName) - 5), depth + 1
    ElseIf Left$(typeName, 11) = "COLLECTION<" And Right$(typeName, 1) = ">" Then
        ValidateConcreteType Mid$(typeName, 12, Len(typeName) - 12), depth + 1
    ElseIf Left$(typeName, 4) = "ROW<" And Right$(typeName, 1) = ">" Then
        inner = AT_CanonicalId(Mid$(typeName, 5, Len(typeName) - 5))
    Else
        If Not Contains(typeName) Then Err.Raise vbObjectError + 2134, "ATRegistry", "REG_UNKNOWN_TYPE: " & typeName
        If GetPrimitive(typeName).Category <> "DATATYPE" Then Err.Raise vbObjectError + 2134, "ATRegistry", "REG_NOT_A_TYPE: " & typeName
    End If
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

