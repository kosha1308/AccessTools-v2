Attribute VB_Name = "modAtExcel"
Option Explicit

Public Function AtNewExcelResolver(ByVal package As CAtPackage, ByVal templateId As String, ByVal columnOffsets As Object, Optional ByVal catalog As Object = Nothing) As CAtExcelResolver
    Dim resolver As New CAtExcelResolver
    resolver.Initialize package, templateId, columnOffsets, catalog
    Set AtNewExcelResolver = resolver
End Function

' Excel boundary codecs only; AST coercion rules remain unchanged.
Public Function AtExcelScalar(ByVal value As Variant, ByVal typeName As String, ByVal objectId As String) As CAtEvalResult
    Dim text As String, supported As Boolean, kind As VbVarType
    kind = VarType(value)
    If IsError(value) Then
        Set AtExcelScalar = AtExcelFailure(typeName, "INVALID", objectId, "R_EXCEL_ERROR", "Excel error value; formula was not recalculated."): Exit Function
    End If
    If IsEmpty(value) Then
        Set AtExcelScalar = AtEvalValue(typeName, "EMPTY", "", objectId): Exit Function
    End If
    If kind = vbString Then
        If Len(value) = 0 Then Set AtExcelScalar = AtEvalValue(typeName, "EMPTY", "", objectId): Exit Function
        text = value
        supported = True
    ElseIf typeName = "Bool" And kind = vbBoolean Then
        text = IIf(value, "true", "false"): supported = True
    ElseIf typeName = "Int64" And (kind = vbDouble Or kind = vbInteger Or kind = vbLong) Then
        If Abs(CDbl(value)) <= 9007199254740991# Then
            If Fix(CDbl(value)) = CDbl(value) Then text = Format$(value, "0"): supported = True
        End If
    End If
    If InStr(1, "|Email|Hostname|IpSet|PortSet|", "|" & typeName & "|", vbBinaryCompare) > 0 Then
        Set AtExcelScalar = AtExcelFailure(typeName, "INVALID", objectId, "R_CODEC_UNSUPPORTED", "No codec is registered for this declared type.")
    ElseIf Not supported Then
        Set AtExcelScalar = AtExcelFailure(typeName, "INVALID", objectId, "R_TYPE", "Cell storage type is not accepted by the declared boundary codec.")
    ElseIf Len(text) > 65536 Then
        Set AtExcelScalar = AtExcelFailure(typeName, "LIMIT", objectId, "R_LIMIT", "Scalar text budget exceeded.")
    ElseIf Not AtLiteralValid(typeName, text) Then
        Set AtExcelScalar = AtExcelFailure(typeName, "INVALID", objectId, "R_TYPE", "Value is not in the declared type's canonical lexical form.")
    Else
        Set AtExcelScalar = AtEvalValue(typeName, "OK", text, objectId)
    End If
End Function

Public Function AtExcelFailure(ByVal typeName As String, ByVal state As String, ByVal objectId As String, ByVal code As String, ByVal message As String) As CAtEvalResult
    Dim r As CAtEvalResult
    Set r = AtEvalValue(typeName, state, "", objectId)
    r.AddDiagnostic code, objectId, "", message
    Set AtExcelFailure = r
End Function

Public Function AtExcelEnvelopeSize(ByVal result As CAtEvalResult) As Long
    Dim record As Variant, member As Variant, size As Long
    size = Len(result.Value)
    For Each record In result.Diagnostics
        For Each member In record.Keys: size = size + Len(CStr(record(member))): Next member
    Next record
    For Each record In result.Provenance
        For Each member In record.Keys: size = size + Len(CStr(record(member))): Next member
    Next record
    AtExcelEnvelopeSize = size
End Function

' Only direct same-workbook A1 names are accepted. Never Evaluate/Range(name).
Public Function AtExcelNameAddress(ByVal refersTo As String, ByVal sheetName As String) As String
    Dim separator As Long, sheet As String, address As String, re As Object
    If Left$(refersTo, 1) <> "=" Or InStr(refersTo, "[") > 0 Or InStr(refersTo, "]") > 0 Then Exit Function
    separator = InStrRev(refersTo, "!")
    If separator < 3 Then Exit Function
    sheet = Mid$(refersTo, 2, separator - 2)
    If Left$(sheet, 1) = "'" And Right$(sheet, 1) = "'" Then sheet = Replace(Mid$(sheet, 2, Len(sheet) - 2), "''", "'")
    If StrComp(sheet, sheetName, vbTextCompare) <> 0 Then Exit Function
    address = Mid$(refersTo, separator + 1)
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^\$[A-Z]{1,3}\$[1-9][0-9]{0,6}(:\$[A-Z]{1,3}\$[1-9][0-9]{0,6})?$"
    If re.Test(address) Then AtExcelNameAddress = address
End Function
