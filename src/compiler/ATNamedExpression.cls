VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNamedExpression"
Option Explicit

Public Id As String
Public ReturnType As String
Public Expression As String
Public Version As String
Public Status As String
Public ReplacedBy As String
Public Parameters As Collection

Public Sub Initialize(ByVal idValue As String, ByVal typeValue As String, ByVal expressionValue As String)
    Id = AT_CanonicalId(idValue)
    ReturnType = UCase$(typeValue)
    Expression = expressionValue
    Version = "1.0.0"
    Status = "ACTIVE"
    Set Parameters = New Collection
End Sub

Public Sub AddParameter(ByVal idValue As String, ByVal typeValue As String, Optional ByVal defaultValue As String = "")
    Dim parameter As New ATFieldDefinition
    parameter.Id = AT_CanonicalId(idValue)
    parameter.TypeName = UCase$(typeValue)
    parameter.DefaultExpression = defaultValue
    Parameters.Add parameter, parameter.Id
End Sub

