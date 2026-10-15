VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATResolver"
Option Explicit

Public Function ResolveField(ByVal reference As String, ByVal context As ATCompileContext, ByVal aliases As Collection) As ATFieldDefinition
    Dim parts As Variant, aliasValue As ATFieldDefinition
    Dim found As ATFieldDefinition, candidate As ATFieldDefinition
    Dim matches As Long
    parts = Split(reference, ".")
    If UBound(parts) = 1 Then
        For Each aliasValue In aliases
            If aliasValue.Id = parts(0) Then
                Set found = context.FindField(CStr(parts(1)), aliasValue.TableId)
                If found Is Nothing Then Err.Raise vbObjectError + 2310, "ATResolver", "RESOLVE_UNKNOWN_ROW_FIELD: " & reference
                Set ResolveField = found
                Exit Function
            End If
        Next aliasValue
        Err.Raise vbObjectError + 2311, "ATResolver", "RESOLVE_UNKNOWN_ALIAS: " & reference
    End If
    Set found = context.FindField(reference)
    If Not found Is Nothing Then matches = 1
    For Each aliasValue In aliases
        Set candidate = context.FindField(reference, aliasValue.TableId)
        If Not candidate Is Nothing Then
            matches = matches + 1
            Set found = candidate
        End If
    Next aliasValue
    If matches = 0 Then Err.Raise vbObjectError + 2312, "ATResolver", "RESOLVE_UNKNOWN_FIELD: " & reference
    If matches > 1 Then Err.Raise vbObjectError + 2313, "ATResolver", "RESOLVE_AMBIGUOUS_FIELD: " & reference
    Set ResolveField = found
End Function

