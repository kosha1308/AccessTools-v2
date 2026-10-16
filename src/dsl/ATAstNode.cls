VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATAstNode"
Option Explicit

Public Kind As String
Public Text As String
Public DataType As String
Public Start As Long
Public Length As Long
Public Children As Collection
Public Expansion As ATAstNode
Public Depth As Long
Public ContextualScalar As Boolean
Public SourceId As String
Public ParameterBindings As Collection
Public PrimitiveId As String
Public Signature As ATSignature
Public ConversionNormalizerId As String
Public FieldId As String
Public TableId As String
Public RowAlias As String
Public NamedScope As String

Private Sub Class_Initialize()
    Set Children = New Collection
    Set ParameterBindings = New Collection
End Sub

