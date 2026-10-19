VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtEvalContext"
Option Explicit

Private mPackageId As String, mVersion As String, mTemplateId As String
Private mFields As Object, mInitialized As Boolean

Friend Sub Initialize(ByVal packageId As String, ByVal version As String, ByVal templateId As String, ByVal fields As Object)
    If mInitialized Then Err.Raise 5, , "Context is immutable"
    If Not AtValidId(packageId) Or Not AtValidVersion(version) Or Not AtValidId(templateId) Then Err.Raise 5, , "Invalid snapshot identity"
    If fields Is Nothing Then Err.Raise 5, , "Fields dictionary required"
    If fields.Count > 2000 Then Err.Raise 5, , "Context field budget exceeded"
    Dim key As Variant, r As CAtEvalResult, total As Long, checked As CAtEvalResult, record As Variant, member As Variant
    Set mFields = AtDictionary()
    For Each key In fields.Keys
        If Not AtValidId(CStr(key)) Then Err.Raise 5, , "Invalid context FieldId"
        Set r = fields(key)
        Set checked = AtEvalValue(r.ResultType, r.State, r.Value, r.Dependency)
        If r.Value <> "" Then total = total + Len(r.Value)
        If Len(r.Value) > 65536 Or total > 1048576 Then Err.Raise 5, , "Context text budget exceeded"
        If r.Requirements.Count <> 0 Or r.Trace.Count <> 0 Then Err.Raise 5, , "Resolver input cannot be a previous evaluation result"
        If r.State = "OK" And r.Diagnostics.Count > 0 Then Err.Raise 5, , "OK input cannot carry error diagnostics"
        If r.Diagnostics.Count > 128 Or r.Provenance.Count > 128 Then Err.Raise 5, , "Context metadata budget exceeded"
        For Each record In r.Diagnostics
            For Each member In record.Keys
                total = total + Len(CStr(record(member)))
                If total > 1048576 Then Err.Raise 5, , "Context metadata budget exceeded"
            Next member
        Next record
        For Each record In r.Provenance
            For Each member In record.Keys
                total = total + Len(CStr(record(member)))
                If total > 1048576 Then Err.Raise 5, , "Context metadata budget exceeded"
            Next member
        Next record
        checked.MergeEffects r
        mFields.Add CStr(key), checked
    Next key
    mPackageId = packageId: mVersion = version: mTemplateId = templateId: mInitialized = True
End Sub

Public Property Get PackageId() As String
    PackageId = mPackageId
End Property
Public Property Get PackageVersion() As String
    PackageVersion = mVersion
End Property
Public Property Get TemplateId() As String
    TemplateId = mTemplateId
End Property
Public Function ReadField(ByVal id As String) As CAtEvalResult
    If Not mInitialized Then Err.Raise 5, , "Context not initialized"
    If mFields.Exists(id) Then Set ReadField = mFields(id).Snapshot()
End Function
