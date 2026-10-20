VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtRuleFinding"
Option Explicit
Public RuleId As String, RuleName As String, FieldId As String, FieldName As String, Target As String
Public Message As String, Severity As String, Outcome As CAtRuleOutcome, Actual As CAtEvalResult
Public Property Get ActualText() As String
    Dim preview As New CAtAdminPreview
    Set preview.Actual = Actual: ActualText = preview.ActualText
End Property
