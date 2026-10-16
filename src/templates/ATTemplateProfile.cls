VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATTemplateProfile"
Option Explicit
Private mId As String, mVersion As String, mSheet As String, mProtection As Boolean
Private mInitialized As Boolean, mSealed As Boolean, mCells As Double
Private mAnchors As Collection, mRelations As Collection, mIds As Object

Public Sub Initialize(ByVal templateId As String, ByVal profileVersion As String, ByVal sheetName As String, Optional ByVal requireSheetProtection As Boolean = False)
    If mInitialized Then Err.Raise vbObjectError + 2830, , "TEMPLATE_PROFILE_FROZEN"
    If Len(Trim$(profileVersion)) = 0 Or Len(profileVersion) > 128 Or Len(sheetName) = 0 Or Len(sheetName) > 31 Then Err.Raise vbObjectError + 2830, , "TEMPLATE_PROFILE_INVALID"
    mId = AT_CanonicalId(templateId): mVersion = profileVersion: mSheet = sheetName: mProtection = requireSheetProtection
    Set mAnchors = New Collection: Set mRelations = New Collection: Set mIds = CreateObject("Scripting.Dictionary")
    mInitialized = True
End Sub
Private Sub Mutable()
    If Not mInitialized Or mSealed Then Err.Raise vbObjectError + 2830, , "TEMPLATE_PROFILE_FROZEN"
End Sub
Public Sub AddAnchor(ByVal id As String, ByVal label As String, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal mergeRows As Long, ByVal mergeCols As Long, Optional ByVal requireLocked As Boolean = True)
    Dim spec As New ATExcelTableSpec, cells As Double
    Mutable
    id = AT_CanonicalId(id)
    If mIds.Exists(id) Or mAnchors.Count >= 256 Then Err.Raise vbObjectError + 2830, , "TEMPLATE_ANCHOR_CONFIG"
    spec.Initialize "TEMPLATE_ANCHORS", firstRow, lastRow, firstCol, lastCol
    spec.AddColumn "ANCHOR", label, "NONE"
    If mergeRows < 1 Or mergeCols < 1 Or mergeRows > lastRow - firstRow + 1 Or mergeCols > lastCol - firstCol + 1 Then Err.Raise vbObjectError + 2830, , "TEMPLATE_ANCHOR_CONFIG"
    cells = CDbl(lastRow - firstRow + 1) * (lastCol - firstCol + 1)
    If mCells + cells > 1000000# Then Err.Raise vbObjectError + 2830, , "TEMPLATE_SCAN_LIMIT"
    mAnchors.Add Array(id, spec, mergeRows, mergeCols, requireLocked, Empty)
    mIds.Add id, True: mCells = mCells + cells
End Sub
Public Sub AddTableAnchor(ByVal id As String, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal headers As Variant, Optional ByVal requireLocked As Boolean = False)
    Dim spec As New ATExcelTableSpec, maps As New Collection, entry As Variant
    Dim col As Long, width As Long, leftCol As Long, rightCol As Long, cells As Double
    Mutable
    id = AT_CanonicalId(id)
    If mIds.Exists(id) Or mAnchors.Count >= 256 Then Err.Raise vbObjectError + 2830, , "TEMPLATE_ANCHOR_CONFIG"
    If Not IsArray(headers) Then Err.Raise vbObjectError + 2830, , "TEMPLATE_ANCHOR_CONFIG"
    If UBound(headers) - LBound(headers) < 1 Then Err.Raise vbObjectError + 2830, , "TEMPLATE_ANCHOR_CONFIG"
    spec.Initialize "TEMPLATE_ANCHORS", firstRow, lastRow, firstCol, lastCol
    leftCol = lastCol: rightCol = firstCol
    For Each entry In headers
        If Not IsArray(entry) Then Err.Raise vbObjectError + 2830, , "TEMPLATE_ANCHOR_CONFIG"
        If LBound(entry) <> 0 Or UBound(entry) <> 2 Then Err.Raise vbObjectError + 2830, , "TEMPLATE_ANCHOR_CONFIG"
        If VarType(entry(0)) <> vbString Or (VarType(entry(1)) <> vbInteger And VarType(entry(1)) <> vbLong) Or (VarType(entry(2)) <> vbInteger And VarType(entry(2)) <> vbLong) Then Err.Raise vbObjectError + 2830, , "TEMPLATE_ANCHOR_CONFIG"
        col = CLng(entry(1)): width = CLng(entry(2))
        If col < firstCol Or col > lastCol Or width < 1 Or width > lastCol - col + 1 Then Err.Raise vbObjectError + 2830, , "TEMPLATE_ANCHOR_CONFIG"
        spec.AddColumn "COL_" & CStr(maps.Count + 1), CStr(entry(0)), "NONE"
        maps.Add Array(col, width)
        If col < leftCol Then leftCol = col
        If col + width - 1 > rightCol Then rightCol = col + width - 1
    Next entry
    cells = CDbl(lastRow - firstRow + 1) * (lastCol - firstCol + 1)
    If mCells + cells > 1000000# Then Err.Raise vbObjectError + 2830, , "TEMPLATE_SCAN_LIMIT"
    mAnchors.Add Array(id, spec, 1, rightCol - leftCol + 1, requireLocked, maps)
    mIds.Add id, True: mCells = mCells + cells
End Sub
Public Sub AddRelation(ByVal fromId As String, ByVal toId As String, ByVal minRows As Long, ByVal maxRows As Long, ByVal minCols As Long, ByVal maxCols As Long)
    Mutable
    fromId = AT_CanonicalId(fromId): toId = AT_CanonicalId(toId)
    If fromId = toId Or minRows > maxRows Or minCols > maxCols Or minRows < -1048575 Or maxRows > 1048575 Or minCols < -16383 Or maxCols > 16383 Or mRelations.Count >= 1024 Then Err.Raise vbObjectError + 2830, , "TEMPLATE_RELATION_CONFIG"
    mRelations.Add Array(fromId, toId, minRows, maxRows, minCols, maxCols)
End Sub
Public Sub Seal()
    Dim relation As Variant
    Mutable
    If mAnchors.Count < 2 Or mRelations.Count = 0 Then Err.Raise vbObjectError + 2830, , "TEMPLATE_PROFILE_INVALID"
    For Each relation In mRelations
        If Not mIds.Exists(relation(0)) Or Not mIds.Exists(relation(1)) Then Err.Raise vbObjectError + 2830, , "TEMPLATE_RELATION_CONFIG"
    Next relation
    mSealed = True
End Sub
Public Function Verify(ByVal ws As Worksheet) As ATTemplateGateResult
    Dim item As Variant, spec As ATExcelTableSpec, layout As ATExcelTableLayout, area As Range
    Dim positions As Object, used As Object, shape As String, locked As Variant
    Dim point As Variant, source As Variant, destination As Variant, deltaRow As Long, deltaCol As Long
    Dim code As String, detail As String, current As String, result As New ATTemplateGateResult
    Dim i As Long, maps As Collection, expected As Variant
    On Error GoTo Failed
    If Not mSealed Then
        code = "TEMPLATE_PROFILE_INVALID": detail = "PROFILE_NOT_SEALED": GoTo BuildResult
    End If
    If ws Is Nothing Then
        code = "TEMPLATE_CHECK_ERROR": detail = "WORKSHEET_REQUIRED": GoTo BuildResult
    End If
    If StrComp(ws.Name, mSheet, vbBinaryCompare) <> 0 Then detail = "SHEET_NAME": GoTo Mismatch
    If mProtection Then
        If Not ws.ProtectContents Then detail = "SHEET_PROTECTION": GoTo Mismatch
    End If
    Set positions = CreateObject("Scripting.Dictionary"): Set used = CreateObject("Scripting.Dictionary")
    For Each item In mAnchors
        current = CStr(item(0)): Set spec = item(1): Set layout = spec.Locate(ws)
        If layout.HeaderLastRow - layout.HeaderFirstRow + 1 <> item(2) Or layout.LastColumn - layout.FirstColumn + 1 <> item(3) Then detail = current & ":MERGE_SHAPE": GoTo Mismatch
        Set maps = Nothing
        If IsObject(item(5)) Then Set maps = item(5)
        For i = 1 To layout.Count
            Set area = ws.Cells(layout.HeaderFirstRow, layout.ColumnNumber(i)).MergeArea
            If Not maps Is Nothing Then
                expected = maps(i)
                If area.Column <> expected(0) Or area.Columns.Count <> expected(1) Or area.Rows.Count <> 1 Then detail = current & ":COLUMN_SHAPE": GoTo Mismatch
            End If
            If area.Cells(1, 1).HasFormula Then detail = current & ":FORMULA_HEADER": GoTo Mismatch
            If CBool(item(4)) Then
                locked = area.Locked
                If IsNull(locked) Then detail = current & ":HEADER_LOCKING": GoTo Mismatch
                If Not CBool(locked) Then detail = current & ":HEADER_LOCKING": GoTo Mismatch
            End If
        Next i
        shape = CStr(layout.HeaderFirstRow) & ":" & CStr(layout.FirstColumn)
        If used.Exists(shape) Then detail = current & ":DUPLICATE_FACT": GoTo Mismatch
        used.Add shape, True
        positions.Add current, Array(layout.HeaderFirstRow, layout.FirstColumn)
    Next item
    For Each item In mRelations
        source = positions(item(0)): destination = positions(item(1))
        deltaRow = destination(0) - source(0): deltaCol = destination(1) - source(1)
        If deltaRow < item(2) Or deltaRow > item(3) Or deltaCol < item(4) Or deltaCol > item(5) Then detail = CStr(item(0)) & ":" & CStr(item(1)) & ":POSITION": GoTo Mismatch
    Next item
    code = "TEMPLATE_MATCH": detail = mId & "@" & mVersion: GoTo BuildResult
Mismatch:
    code = "TEMPLATE_MISMATCH"
BuildResult:
    On Error GoTo 0
    result.Initialize code = "TEMPLATE_MATCH", code, detail
    Set Verify = result
    Exit Function
Failed:
    If Left$(Err.Description, 6) = "EXCEL_" Then
        detail = current & ":" & Err.Description: code = "TEMPLATE_MISMATCH"
    Else
        detail = current & ":READ_FAILED": code = "TEMPLATE_CHECK_ERROR"
    End If
    Resume BuildResult
End Function
Public Property Get TemplateId() As String
    TemplateId = mId
End Property
Public Property Get ProfileVersion() As String
    ProfileVersion = mVersion
End Property
