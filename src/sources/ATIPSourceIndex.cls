VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATIPSourceIndex"
Option Explicit
Private mReady As Boolean, mRows As ATRuntimeValue, mId As String, mFingerprint As String, mIssues As ATEvalResult, mIssueCount As Long
Private mLow() As Double, mHigh() As Double, mPrefix() As Double, mOrdinal() As Long, mCount As Long, mCapacity As Long
Friend Sub Initialize(ByVal raw As ATEvalResult, ByVal sourceId As String, ByVal fields As Collection, ByVal fingerprint As String)
    Dim limits As New ATEvalLimits, budget As New ATEvalBudget, row As ATRuntimeRow, field As Variant, cell As ATEvalResult
    Dim value As ATRuntimeValue, parsed As ATEvalResult, parsedValue As ATRuntimeValue, region As ATIntervalSet, issues As New Collection, issueValue As New ATRuntimeValue
    Dim text As String, id As String, key As String, code As String, i As Long, j As Long, maximum As Double
    If mReady Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_FROZEN"
    key = AT_IPSourceKey(sourceId, fields): mId = AT_CanonicalId(sourceId)
    If Len(fingerprint) = 0 Or Len(fingerprint) > 4096 Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_FINGERPRINT_INVALID"
    mFingerprint = fingerprint
    limits.MaxSteps = 20000000: limits.MaxListItems = 100000: budget.Initialize limits
    If raw Is Nothing Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_INPUT_INVALID"
    If raw.State <> atEvalOK Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_INPUT_INVALID"
    Set mRows = raw.Value
    If mRows.DataType <> "COLLECTION<ROW<" & mId & "_RAW>>" Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_SCHEMA_INVALID"
    budget.CheckList mRows.Count: budget.Step mRows.Count * 2
    For i = 1 To mRows.Count
        Set row = mRows.Row(i): Set region = New ATIntervalSet: region.Initialize 4294967295#
        For Each field In fields
            budget.Step: id = AT_CanonicalId(CStr(field)): Set cell = row.ResolveField(id)
            text = "": code = ""
            If cell.State <> atEvalOK Then
                If cell.Diagnostics.Count = 1 Then
                    If cell.Diagnostics(1).Code = "EXCEL_CELL_ERROR" Then code = "EXCEL_CELL_ERROR"
                End If
                If Len(code) = 0 Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_SCHEMA_INVALID"
            Else
                Set value = cell.Value
                If value.DataType <> "TEXT" Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_SCHEMA_INVALID"
                If value.ValueState = atValuePresent Then
                    text = CStr(value.Scalar): budget.CheckText Len(text): budget.Step Len(text) + 1
                ElseIf value.ValueState <> atValueEmpty Then
                    Err.Raise vbObjectError + 3000, , "IP_SOURCE_SCHEMA_INVALID"
                End If
                If Not AT_ExcelBlank(text) Then
                    Set parsed = AT_ParseDomain("IP_EXPRESSION", Replace(Replace(text, "х", "x"), "Х", "x"), Nothing, budget)
                    If parsed.State <> atEvalOK Then Err.Raise vbObjectError + 3000, , parsed.Diagnostics(1).Code
                    Set parsedValue = parsed.Value
                    If parsedValue.ValueState = atValuePresent Then
                        For j = 1 To parsedValue.Domain.Count
                            region.Add parsedValue.Domain.Item(j).Low, parsedValue.Domain.Item(j).High, budget
                        Next j
                    Else
                        code = "IP_SOURCE_CELL_INVALID"
                        If parsed.Diagnostics.Count > 0 Then code = parsed.Diagnostics(1).Code
                    End If
                End If
            End If
            If Len(code) > 0 Then AddDataIssue issues, row, id, text, code, budget
        Next field
        For j = 1 To region.Count: Append region.Item(j).Low, region.Item(j).High, i, budget: Next j
    Next i
    SortIntervals budget
    If mCount > 0 Then
        ReDim mPrefix(1 To mCount)
        For i = 1 To mCount
            budget.Step
            If mHigh(i) > maximum Then maximum = mHigh(i)
            mPrefix(i) = maximum
        Next i
    End If
    mIssueCount = issues.Count: issueValue.InitializeRows "IP_SOURCE_ISSUE", issues: Set mIssues = AT_RuntimeOK(issueValue)
    mReady = True
End Sub
Private Sub Append(ByVal low As Double, ByVal high As Double, ByVal ordinal As Long, ByVal budget As ATEvalBudget)
    budget.Step
    If mCount >= 200000 Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_INTERVAL_LIMIT"
    If mCount = mCapacity Then
        If mCapacity = 0 Then mCapacity = 128 Else mCapacity = mCapacity * 2
        If mCapacity > 200000 Then mCapacity = 200000
        ReDim Preserve mLow(1 To mCapacity): ReDim Preserve mHigh(1 To mCapacity): ReDim Preserve mOrdinal(1 To mCapacity)
    End If
    mCount = mCount + 1: mLow(mCount) = low: mHigh(mCount) = high: mOrdinal(mCount) = ordinal
End Sub
Private Sub SortIntervals(ByVal budget As ATEvalBudget)
    Dim lows() As Double, highs() As Double, ordinals() As Long, width As Long, first As Long, middle As Long, last As Long
    Dim left As Long, right As Long, target As Long, chosen As Long
    If mCount < 2 Then Exit Sub
    ReDim lows(1 To mCount): ReDim highs(1 To mCount): ReDim ordinals(1 To mCount): width = 1
    Do While width < mCount
        For first = 1 To mCount Step width * 2
            middle = first + width - 1: If middle > mCount Then middle = mCount
            last = first + width * 2 - 1: If last > mCount Then last = mCount
            left = first: right = middle + 1
            For target = first To last
                budget.Step
                If left > middle Then
                    chosen = right: right = right + 1
                ElseIf right > last Then
                    chosen = left: left = left + 1
                ElseIf mLow(left) <= mLow(right) Then
                    chosen = left: left = left + 1
                Else
                    chosen = right: right = right + 1
                End If
                lows(target) = mLow(chosen): highs(target) = mHigh(chosen): ordinals(target) = mOrdinal(chosen)
            Next target
        Next first
        For target = 1 To mCount
            budget.Step: mLow(target) = lows(target): mHigh(target) = highs(target): mOrdinal(target) = ordinals(target)
        Next target
        width = width * 2
    Loop
End Sub
Private Sub AddDataIssue(ByVal issues As Collection, ByVal source As ATRuntimeRow, ByVal field As String, ByVal text As String, ByVal code As String, ByVal budget As ATEvalBudget)
    Dim row As New ATRuntimeRow
    budget.CheckList issues.Count + 1: budget.Step
    row.Initialize "IP_SOURCE_ISSUE", CStr(issues.Count + 1)
    row.AddField "SOURCE_ROW", AT_RuntimeOK(AT_RuntimeScalar("TEXT", source.RowId))
    row.AddField "SOURCE_FIELD", AT_RuntimeOK(AT_RuntimeScalar("TEXT", field))
    row.AddField "VALUE", AT_RuntimeOK(AT_RuntimeScalar("TEXT", text)), source.CellRef(field)
    row.AddField "CODE", AT_RuntimeOK(AT_RuntimeScalar("TEXT", code))
    issues.Add row
End Sub
Public Function Query(ByVal expression As String, ByVal budget As ATEvalBudget) As ATIPMatchResult
    Dim r As New ATIPMatchResult, f As ATEvalResult, parsed As ATEvalResult, parsedValue As ATRuntimeValue, rows As New Collection, value As New ATRuntimeValue
    Dim hits As Object, positions() As Long, keys As Variant, part As ATIntervalRange, i As Long, position As Long
    Dim low As Long, high As Long, middle As Long
    On Error GoTo Failed
    RequireReady
    If budget Is Nothing Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_BUDGET_REQUIRED"
    Set parsed = AT_ParseDomain("IP_EXPRESSION", Replace(Replace(expression, "х", "x"), "Х", "x"), Nothing, budget)
    If parsed.State <> atEvalOK Then Err.Raise vbObjectError + 3000, , parsed.Diagnostics(1).Code
    Set parsedValue = parsed.Value
    If parsedValue.ValueState <> atValuePresent Then
        Set f = AT_RuntimeFailure(atEvalManualReview, "IP_SOURCE_QUERY_INVALID", Nothing)
        r.Initialize atEvalManualReview, Nothing, False, f.Diagnostics: Set Query = r: Exit Function
    End If
    Set hits = CreateObject("Scripting.Dictionary")
    For i = 1 To parsedValue.Domain.Count
        Set part = parsedValue.Domain.Item(i): low = 1: high = mCount: position = 0
        Do While low <= high
            budget.Step: middle = low + (high - low) \ 2
            If mLow(middle) <= part.High Then position = middle: low = middle + 1 Else high = middle - 1
        Loop
        Do While position > 0
            budget.Step
            If mPrefix(position) < part.Low Then Exit Do
            If mHigh(position) >= part.Low Then
                If Not hits.Exists(CStr(mOrdinal(position))) Then
                    budget.CheckList hits.Count + 1: hits.Add CStr(mOrdinal(position)), mOrdinal(position)
                End If
            End If
            position = position - 1
        Loop
    Next i
    If hits.Count > 0 Then
        keys = hits.Items: ReDim positions(1 To hits.Count)
        For i = 1 To hits.Count: budget.Step: positions(i) = CLng(keys(i - 1)): Next i
        AT_IPSortPositions positions, hits.Count, budget
        For i = 1 To hits.Count: budget.Step 4: rows.Add mRows.Row(positions(i)): Next i
    End If
    value.InitializeRows mRows.RowTableId, rows
    r.Initialize atEvalOK, value, mIssueCount = 0, Nothing: Set Query = r: Exit Function
Failed:
    Set f = AT_IPSourceFailure(Err.Description)
    r.Initialize atEvalError, Nothing, False, f.Diagnostics: Set Query = r
End Function
Private Sub RequireReady()
    If Not mReady Then Err.Raise vbObjectError + 3000, , "IP_SOURCE_ACCESS"
End Sub
Public Property Get SourceRowCount() As Long
    RequireReady: SourceRowCount = mRows.Count
End Property
Public Property Get IssueCount() As Long
    RequireReady: IssueCount = mIssueCount
End Property
Public Property Get SourceId() As String
    RequireReady: SourceId = mId
End Property
Public Property Get Fingerprint() As String
    RequireReady: Fingerprint = mFingerprint
End Property
Public Property Get DataIssues() As ATEvalResult
    RequireReady: Set DataIssues = mIssues.Clone
End Property
Public Property Get IntervalCount() As Long
    RequireReady: IntervalCount = mCount
End Property
