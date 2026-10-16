VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATRuntimeValue"
Option Explicit

Private mType As String
Private mState As ATValueState
Private mScalar As Variant
Private mItems As Collection
Private mInitialized As Boolean
Private mDomain As ATIntervalSet
Private mDeclarations As Collection

Public Sub InitializeScalar(ByVal typeName As String, ByVal state As ATValueState, ByVal payload As Variant)
    If mInitialized Then Err.Raise vbObjectError + 2500, , "RUNTIME_VALUE_FROZEN"
    If state < atValuePresent Or state > atValueNull Then Err.Raise vbObjectError + 2501, , "RUNTIME_VALUE_STATE"
    mType = UCase$(typeName)
    If Len(mType) = 0 Then Err.Raise vbObjectError + 2501, , "RUNTIME_TYPE_MISMATCH"
    If state = atValuePresent Then
        Select Case mType
            Case "TEXT": If VarType(payload) <> vbString Then BadType
            Case "NUMBER": If VarType(payload) <> vbDouble Then BadType
            Case "BOOL": If VarType(payload) <> vbBoolean Then BadType
            Case "DATE"
                If VarType(payload) <> vbDate Then BadType
                If CDbl(payload) <> Fix(CDbl(payload)) Then BadType
            Case Else: BadType
        End Select
        mScalar = payload
    End If
    mState = state
    mInitialized = True
End Sub

Private Sub BadType()
    Err.Raise vbObjectError + 2501, "ATRuntimeValue", "RUNTIME_TYPE_MISMATCH"
End Sub

Public Sub InitializeList(ByVal elementType As String, ByVal items As Collection)
    Dim item As ATRuntimeValue
    If mInitialized Then Err.Raise vbObjectError + 2500, , "RUNTIME_VALUE_FROZEN"
    If Not AT_RuntimeScalarType(elementType) Then BadType
    Set mItems = New Collection
    For Each item In items
        If item.DataType <> elementType Then BadType
        mItems.Add item.Clone
    Next item
    mType = "LIST<" & elementType & ">"
    mState = atValuePresent
    mInitialized = True
End Sub

Public Property Get DataType() As String
    DataType = mType
End Property
Public Property Get ValueState() As ATValueState
    ValueState = mState
End Property
Public Property Get Scalar() As Variant
    If Not mInitialized Or mState <> atValuePresent Or Not mItems Is Nothing Then Err.Raise vbObjectError + 2502, , "RUNTIME_VALUE_ACCESS"
    If Not mDomain Is Nothing Then Err.Raise vbObjectError + 2502, , "RUNTIME_VALUE_ACCESS"
    Scalar = mScalar
End Property
Public Property Get Count() As Long
    If mItems Is Nothing Then Err.Raise vbObjectError + 2502, , "RUNTIME_VALUE_ACCESS"
    Count = mItems.Count
End Property
Public Function Item(ByVal index As Long) As ATRuntimeValue
    If mItems Is Nothing Then Err.Raise vbObjectError + 2502, , "RUNTIME_VALUE_ACCESS"
    Set Item = mItems(index).Clone
End Function
Public Function Clone() As ATRuntimeValue
    Dim result As New ATRuntimeValue
    If Not mInitialized Then Err.Raise vbObjectError + 2502, , "RUNTIME_VALUE_ACCESS"
    If Not mDomain Is Nothing Then
        result.InitializeDomain mType, mDomain, mDeclarations
    ElseIf mItems Is Nothing Then
        result.InitializeScalar mType, mState, mScalar
    Else
        result.InitializeList Mid$(mType, 6, Len(mType) - 6), mItems
    End If
    Set Clone = result
End Function

Public Sub ValidateBudget(ByVal budget As ATEvalBudget)
    Dim item As ATRuntimeValue
    If mState <> atValuePresent Then Exit Sub
    If mType = "TEXT" Then budget.CheckText Len(mScalar)
    If Not mDomain Is Nothing Then
        budget.CheckList mDomain.Count
        budget.CheckList mDeclarations.Count
        budget.Step mDomain.Count + mDeclarations.Count
    End If
    If Not mItems Is Nothing Then
        budget.CheckList mItems.Count
        For Each item In mItems
            budget.Step
            item.ValidateBudget budget
        Next item
    End If
End Sub

Public Sub InitializeDomain(ByVal typeName As String, ByVal intervals As ATIntervalSet, ByVal declarations As Collection)
    Dim key As Variant, expectedMaximum As Double
    If mInitialized Then Err.Raise vbObjectError + 2500, , "RUNTIME_VALUE_FROZEN"
    If Not AT_RuntimeDomainType(typeName) Then BadType
    If intervals Is Nothing Then BadType
    If Not intervals.Frozen Then BadType
    expectedMaximum = 4294967295#
    If typeName = "PORT" Or typeName = "PORT_SET" Then expectedMaximum = 65535#
    If intervals.Maximum <> expectedMaximum Then BadType
    Set mDomain = intervals
    Set mDeclarations = New Collection
    For Each key In declarations
        mDeclarations.Add CStr(key), CStr(key)
    Next key
    mType = typeName: mState = atValuePresent: mInitialized = True
End Sub
Public Property Get Domain() As ATIntervalSet
    If mDomain Is Nothing Then Err.Raise vbObjectError + 2502, , "RUNTIME_VALUE_ACCESS"
    Set Domain = mDomain
End Property
Public Property Get DeclarationKeys() As Collection
    Dim copy As New Collection, key As Variant
    If mDeclarations Is Nothing Then Err.Raise vbObjectError + 2502, , "RUNTIME_VALUE_ACCESS"
    For Each key In mDeclarations
        copy.Add CStr(key), CStr(key)
    Next key
    Set DeclarationKeys = copy
End Property
