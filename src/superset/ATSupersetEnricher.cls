VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATSupersetEnricher"
Option Explicit

Public Function Read(ByVal ws As Worksheet, ByVal layout As ATExcelTableLayout, ByVal lastDataRow As Long, ByVal summaryIds As Collection) As ATEvalResult
    Dim buckets As Object, ordered As New Collection, raw As ATEvalResult, reader As New ATExcelTableReader
    Dim source As ATRuntimeValue
    Dim item As Variant, id As String, bucket As ATSupersetBucket, row As ATRuntimeRow, ids As Collection
    Dim link As String, key As Variant, values(0 To 4) As Variant, fields As Variant, i As Long
    Dim relevant As Boolean, work As Long, rows As New Collection, result As New ATRuntimeValue, code As String
    On Error GoTo Failed
    ValidateSchema layout
    If summaryIds Is Nothing Then Err.Raise vbObjectError + 2820, , "SUPERSET_SUMMARY_IDS"
    If summaryIds.Count > 10000 Then Err.Raise vbObjectError + 2820, , "SUPERSET_SUMMARY_IDS"
    Set buckets = CreateObject("Scripting.Dictionary")
    For Each item In summaryIds
        id = CanonicalNumber(item, "SUPERSET_SUMMARY_IDS")
        If buckets.Exists(id) Then Err.Raise vbObjectError + 2820, , "SUPERSET_SUMMARY_IDS"
        Set bucket = New ATSupersetBucket: bucket.Initialize id
        buckets.Add id, bucket: ordered.Add id
    Next item
    Set raw = reader.Read(ws, layout, lastDataRow)
    If raw.State <> atEvalOK Then
        Set Read = raw
        Exit Function
    End If
    fields = Array("HOSTNAME", "SERVICE", "IS_NAME", "SKUF_GROUP", "SUPPORT_TEAM")
    Set source = raw.Value
    For i = 1 To source.Count
        StepWork work
        Set row = source.Row(i)
        If Not ws.Rows(CLng(row.RowId)).Hidden Then
            link = CellText(row, "SUMMARY_LINK")
            Set ids = ParseLinks(link, work): relevant = False
            For Each key In ids
                If buckets.Exists(CStr(key)) Then relevant = True
            Next key
            If relevant Then
                For Each item In Array(0, 1, 2, 3, 4)
                    values(CLng(item)) = CellText(row, CStr(fields(CLng(item))))
                Next item
                For Each key In ids
                    If buckets.Exists(CStr(key)) Then
                        StepWork work
                        Set bucket = buckets(CStr(key)): bucket.AddMatch values
                    End If
                Next key
            End If
        End If
    Next i
    For Each key In ordered
        Set bucket = buckets(CStr(key)): rows.Add bucket.Snapshot
    Next key
    result.InitializeRows "SUPERSET_SUMMARY", rows
    Set Read = AT_RuntimeOK(result)
    Exit Function
Failed:
    code = Err.Description
    If Left$(code, 9) <> "SUPERSET_" And Left$(code, 6) <> "EXCEL_" Then code = "SUPERSET_READ_FAILED"
    Set Read = AT_RuntimeFailure(atEvalError, code, Nothing)
End Function

Private Sub ValidateSchema(ByVal layout As ATExcelTableLayout)
    Dim names As Object, i As Long, id As Variant
    If layout Is Nothing Then Err.Raise vbObjectError + 2820, , "SUPERSET_SCHEMA"
    Set names = CreateObject("Scripting.Dictionary")
    For i = 1 To layout.Count
        names.Add layout.FieldId(i), layout.Usage(i)
    Next i
    For Each id In Array("SUMMARY_LINK", "HOSTNAME", "SERVICE", "IS_NAME", "SKUF_GROUP", "SUPPORT_TEAM")
        If Not names.Exists(CStr(id)) Then Err.Raise vbObjectError + 2820, , "SUPERSET_SCHEMA"
        If id = "SUMMARY_LINK" Then
            If names(CStr(id)) <> "DIRECT" Then Err.Raise vbObjectError + 2820, , "SUPERSET_SCHEMA"
        Else
            If names(CStr(id)) <> "NONE" Then Err.Raise vbObjectError + 2820, , "SUPERSET_SCHEMA"
        End If
    Next id
End Sub

Private Function CellText(ByVal row As ATRuntimeRow, ByVal id As String) As String
    Dim field As ATEvalResult, value As ATRuntimeValue
    Set field = row.ResolveField(id)
    If field.State <> atEvalOK Then
        If field.Diagnostics.Count > 0 Then Err.Raise vbObjectError + 2820, , field.Diagnostics(1).Code
        Err.Raise vbObjectError + 2820, , "SUPERSET_CELL_ERROR"
    End If
    Set value = field.Value
    If value.DataType <> "TEXT" Then Err.Raise vbObjectError + 2820, , "SUPERSET_VALUE_STATE"
    Select Case value.ValueState
        Case atValueEmpty: CellText = ""
        Case atValuePresent: CellText = CStr(value.Scalar)
        Case Else: Err.Raise vbObjectError + 2820, , "SUPERSET_VALUE_STATE"
    End Select
End Function

Private Function ParseLinks(ByVal text As String, ByRef work As Long) As Collection
    Dim keys As Object, result As New Collection, part As Variant, id As String
    Set keys = CreateObject("Scripting.Dictionary")
    text = Replace(Replace(Replace(text, vbCrLf, ","), vbCr, ","), vbLf, ",")
    text = Replace(text, ";", ",")
    For Each part In Split(text, ",")
        StepWork work
        If Not AT_ExcelBlank(part) Then
            id = CanonicalNumber(part, "SUPERSET_LINK_INVALID")
            If Not keys.Exists(id) Then keys.Add id, True: result.Add id
        End If
    Next part
    Set ParseLinks = result
End Function

Private Function CanonicalNumber(ByVal raw As Variant, ByVal code As String) As String
    Dim text As String, i As Long, char As String, first As Long
    Select Case VarType(raw)
        Case vbString: text = Trim$(Replace(CStr(raw), ChrW(160), " "))
        Case vbByte, vbInteger, vbLong: text = CStr(raw)
        Case Else: Err.Raise vbObjectError + 2820, , code
    End Select
    If Len(text) = 0 Or Len(text) > 32767 Then Err.Raise vbObjectError + 2820, , code
    For i = 1 To Len(text)
        char = Mid$(text, i, 1)
        If char < "0" Or char > "9" Then Err.Raise vbObjectError + 2820, , code
    Next i
    first = 1
    Do While first < Len(text)
        If Mid$(text, first, 1) <> "0" Then Exit Do
        first = first + 1
    Loop
    text = Mid$(text, first)
    If Len(text) > 10 Or text = "0" Then Err.Raise vbObjectError + 2820, , code
    If Len(text) = 10 Then
        If StrComp(text, "2147483647", vbBinaryCompare) > 0 Then Err.Raise vbObjectError + 2820, , code
    End If
    CanonicalNumber = text
End Function

Private Sub StepWork(ByRef work As Long)
    work = work + 1
    If work > 200000 Then Err.Raise vbObjectError + 2820, , "SUPERSET_WORK_LIMIT"
End Sub
