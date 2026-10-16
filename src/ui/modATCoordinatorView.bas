Attribute VB_Name = "modATCoordinatorView"
Option Explicit
Public Function AT_ViewFields() As Variant
    AT_ViewFields = Array("HOSTNAME", "SOURCE_TELEGRAPH", "SOURCE_CXD", "SOURCE_AWX", "IS_NAME", "SERVICE", "SKUF_GROUP", "SUPPORT_TEAM")
End Function
Public Function AT_ViewHeaders() As Variant
    AT_ViewHeaders = Array("№ из сводки", "FQDN", "SOURCE:ТЕЛЕГРАФ", "SOURCE: ЦХД", "SOURCE: AWX", "ПОЛНОЕ ИМЯ ИС", "СЕРВИС", "СКУФ ГРУППА ТП", "КОМАНДА ПОДДЕРЖКИ")
End Function
Public Function AT_ViewLiteral(ByVal text As String) As String
    If Len(text) > 32767 Then Err.Raise vbObjectError + 3120, , "VIEW_TEXT_LIMIT"
    If Left$(text, 1) = "'" Then text = "'" & text
    AT_ViewLiteral = text
End Function
Public Function AT_ViewText(ByVal field As ATEvalResult) As String
    Dim valid As Boolean, budget As New ATEvalBudget, limits As New ATEvalLimits
    budget.Initialize limits
    AT_ViewText = AT_RegistryText(field, valid, budget)
    If Not valid Then Err.Raise vbObjectError + 3120, , "VIEW_CELL_ERROR"
End Function
Public Function AT_ViewRows(ByVal entries As Collection, ByRef rowCount As Long) As Variant
    Dim rows As Object, links As Object, entry As ATNetworkRegistryEntry, matches As ATRuntimeValue, row As ATRuntimeRow
    Dim fields As Variant, data() As Variant, keys() As Long, key As String, i As Long, j As Long, linkCount As Long
    Dim field As ATEvalResult, text As String, totalText As Long, value As Variant
    Set rows = CreateObject("Scripting.Dictionary"): Set links = CreateObject("Scripting.Dictionary")
    fields = AT_ViewFields()
    For Each entry In entries
        Set matches = entry.Matches
        For i = 1 To matches.Count
            linkCount = linkCount + 1
            If linkCount > 100000 Then Err.Raise vbObjectError + 3120, , "VIEW_LINK_LIMIT"
            Set row = matches.Row(i): key = row.RowId
            If Len(key) = 0 Or Len(key) > 7 Then Err.Raise vbObjectError + 3120, , "VIEW_ROW_ID"
            If Not key Like String$(Len(key), "#") Then Err.Raise vbObjectError + 3120, , "VIEW_ROW_ID"
            If CLng(key) < 1 Or CLng(key) > 1048576 Or CStr(CLng(key)) <> key Then Err.Raise vbObjectError + 3120, , "VIEW_ROW_ID"
            If Not rows.Exists(key) Then
                If rows.Count >= 10000 Then Err.Raise vbObjectError + 3120, , "VIEW_ROW_LIMIT"
                rows.Add key, row: links.Add key, entry.Id
            Else
                text = links(key) & "," & entry.Id
                If Len(text) > 32767 Then Err.Raise vbObjectError + 3120, , "VIEW_TEXT_LIMIT"
                links(key) = text
            End If
        Next i
    Next entry
    rowCount = rows.Count
    If rowCount = 0 Then Exit Function
    ReDim keys(1 To rowCount): i = 0
    For Each value In rows.Keys: i = i + 1: keys(i) = CLng(value): Next value
    SortKeys keys, 1, rowCount
    ReDim data(1 To rowCount, 1 To 9)
    For i = 1 To rowCount
        key = CStr(keys(i)): Set row = rows(key): data(i, 1) = links(key)
        totalText = totalText + Len(CStr(data(i, 1)))
        For j = 0 To 7
            Set field = row.ResolveField(CStr(fields(j)))
            If field.State = atEvalError And field.Diagnostics.Count = 1 Then
                If field.Diagnostics(1).Code = "EXCEL_CELL_ERROR" Then
                    data(i, j + 2) = CVErr(xlErrValue)
                Else
                    Err.Raise vbObjectError + 3120, , "VIEW_INPUT_INVALID"
                End If
            Else
                text = AT_ViewText(field): totalText = totalText + Len(text)
                data(i, j + 2) = AT_ViewLiteral(text)
            End If
            If totalText > 2000000 Then Err.Raise vbObjectError + 3120, , "VIEW_TEXT_LIMIT"
        Next j
    Next i
    AT_ViewRows = data
End Function
Private Sub SortKeys(ByRef keys() As Long, ByVal low As Long, ByVal high As Long)
    Dim left As Long, right As Long, pivot As Long, saved As Long
    left = low: right = high: pivot = keys((low + high) \ 2)
    Do While left <= right
        Do While keys(left) < pivot: left = left + 1: Loop
        Do While keys(right) > pivot: right = right - 1: Loop
        If left <= right Then
            saved = keys(left): keys(left) = keys(right): keys(right) = saved
            left = left + 1: right = right - 1
        End If
    Loop
    If low < right Then SortKeys keys, low, right
    If left < high Then SortKeys keys, left, high
End Sub
Public Function AT_ViewSpec() As ATExcelTableSpec
    Dim spec As New ATExcelTableSpec
    spec.Initialize "SUPERSET_VIEW", 5, 5, 1, 9
    spec.AddColumn "SUMMARY_LINK", "№ из сводки", "DIRECT"
    spec.AddColumn "HOSTNAME", "FQDN", "NONE"
    spec.AddColumn "SERVICE", "СЕРВИС", "NONE"
    spec.AddColumn "IS_NAME", "ПОЛНОЕ ИМЯ ИС", "NONE"
    spec.AddColumn "SKUF_GROUP", "СКУФ ГРУППА ТП", "NONE"
    spec.AddColumn "SUPPORT_TEAM", "КОМАНДА ПОДДЕРЖКИ", "NONE"
    Set AT_ViewSpec = spec
End Function
