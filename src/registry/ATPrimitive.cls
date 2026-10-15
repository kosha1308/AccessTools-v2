VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATPrimitive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mId As String
Private mCategory As String
Private mSignatures As Collection
Private mFrozen As Boolean
Private mDescription As String
Private mEvaluatorId As String
Private mVersion As String
Private mStatus As String
Private mDeterminism As String
Private mExamples As String
Private mTestContract As String
Private mSyntax As String

Public Sub Initialize(ByVal idValue As String, ByVal categoryValue As String, ByVal descriptionValue As String, ByVal evaluatorValue As String, ByVal versionValue As String, ByVal statusValue As String, ByVal determinismValue As String, ByVal examplesValue As String, ByVal testContractValue As String)
    If mFrozen Then Err.Raise vbObjectError + 2120, "ATPrimitive", "REG_FROZEN"
    mId = AT_CanonicalId(idValue)
    mSyntax = mId
    mCategory = UCase$(categoryValue)
    Select Case mCategory
        Case "DATATYPE", "OPERATOR", "FUNCTION", "QUANTIFIER"
        Case Else: Err.Raise vbObjectError + 2121, "ATPrimitive", "REG_INVALID_CATEGORY"
    End Select
    If Len(descriptionValue) = 0 Or Len(evaluatorValue) = 0 Or Len(versionValue) = 0 Or Len(examplesValue) = 0 Or Len(testContractValue) = 0 Then Err.Raise vbObjectError + 2122, "ATPrimitive", "REG_REQUIRED_METADATA"
    If statusValue <> "ACTIVE" And statusValue <> "DEPRECATED" Then Err.Raise vbObjectError + 2123, "ATPrimitive", "REG_INVALID_STATUS"
    If determinismValue <> "PURE" And determinismValue <> "CONTEXT" Then Err.Raise vbObjectError + 2124, "ATPrimitive", "REG_INVALID_DETERMINISM"
    mDescription = descriptionValue
    mEvaluatorId = evaluatorValue
    mVersion = versionValue
    mStatus = statusValue
    mDeterminism = determinismValue
    mExamples = examplesValue
    mTestContract = testContractValue
    Set mSignatures = New Collection
End Sub

Public Sub SetSyntax(ByVal syntaxValue As String)
    If mFrozen Then Err.Raise vbObjectError + 2120, "ATPrimitive", "REG_FROZEN"
    mSyntax = syntaxValue
End Sub

Public Property Get Syntax() As String
    Syntax = mSyntax
End Property

Public Property Get DisplayName() As String
    DisplayName = mId
End Property

Public Property Get CachePolicy() As String
    If mDeterminism = "PURE" Then CachePolicy = "SESSION" Else CachePolicy = "RUN_CONTEXT"
End Property

Public Property Get Associativity() As String
    If mSyntax = "NOT" Or mSyntax = "UNARY_PLUS" Or mSyntax = "UNARY_MINUS" Then
        Associativity = "RIGHT"
    ElseIf Precedence = 5 Then
        Associativity = "NONE"
    Else
        Associativity = "LEFT"
    End If
End Property

Public Property Get Precedence() As Long
    Select Case mSyntax
        Case "OR": Precedence = 8
        Case "AND": Precedence = 7
        Case "NOT": Precedence = 6
        Case "+", "-": Precedence = 4
        Case "*", "/", "%": Precedence = 3
        Case "UNARY_PLUS", "UNARY_MINUS": Precedence = 2
        Case Else
            If mCategory = "OPERATOR" Then Precedence = 5 Else Precedence = 1
    End Select
End Property

Public Sub AddSignature(ByVal argumentTypes As String, ByVal resultType As String)
    Dim signature As New ATSignature
    If mFrozen Then Err.Raise vbObjectError + 2120, "ATPrimitive", "REG_FROZEN"
    If mSignatures Is Nothing Then Err.Raise vbObjectError + 2125, "ATPrimitive", "REG_UNINITIALIZED"
    signature.Initialize argumentTypes, resultType
    mSignatures.Add signature
End Sub

Public Sub Freeze()
    Dim signature As ATSignature
    If mSignatures Is Nothing Then Err.Raise vbObjectError + 2125, "ATPrimitive", "REG_UNINITIALIZED"
    If mSignatures.Count = 0 Then Err.Raise vbObjectError + 2126, "ATPrimitive", "REG_MISSING_SIGNATURE"
    For Each signature In mSignatures
        signature.Freeze
    Next signature
    mFrozen = True
End Sub

Public Property Get Id() As String
    Id = mId
End Property
Public Property Get Category() As String
    Category = mCategory
End Property
Public Property Get Description() As String
    Description = mDescription
End Property
Public Property Get EvaluatorId() As String
    EvaluatorId = mEvaluatorId
End Property
Public Property Get Version() As String
    Version = mVersion
End Property
Public Property Get Status() As String
    Status = mStatus
End Property
Public Property Get Determinism() As String
    Determinism = mDeterminism
End Property
Public Property Get Examples() As String
    Examples = mExamples
End Property
Public Property Get TestContract() As String
    TestContract = mTestContract
End Property
Public Property Get Signatures() As Collection
    Dim copy As New Collection
    Dim item As ATSignature
    If mSignatures Is Nothing Then Err.Raise vbObjectError + 2125, "ATPrimitive", "REG_UNINITIALIZED"
    For Each item In mSignatures
        copy.Add item
    Next item
    Set Signatures = copy
End Property

