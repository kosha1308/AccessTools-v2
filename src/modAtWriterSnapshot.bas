Attribute VB_Name = "modAtWriterSnapshot"
Option Explicit

Public Function AtWHex(ByVal text As String) As String
    Dim i As Long, result As String
    For i = 1 To Len(text): result = result & Right$("0000" & Hex$(AscW(Mid$(text, i, 1)) And &HFFFF&), 4): Next i
    AtWHex = result
End Function
Public Function AtWUnhex(ByVal text As String) As String
    Dim i As Long, result As String
    If Len(text) Mod 4 <> 0 Then Err.Raise 5, , "Invalid journal scalar"
    For i = 1 To Len(text) Step 4: result = result & ChrW$(CLng("&H" & Mid$(text, i, 4))): Next i
    AtWUnhex = result
End Function
Public Function AtWPack(ByVal record As Object) As String
    Dim key As Variant, result As String
    For Each key In record.Keys: result = result & CStr(key) & "=" & AtWHex(CStr(record(key))) & vbLf: Next key
    AtWPack = result
End Function
Public Function AtWUnpack(ByVal text As String) As Object
    Dim result As Object, line As Variant, separator As Long
    Set result = AtDictionary()
    For Each line In Split(text, vbLf)
        If line <> "" Then
            separator = InStr(line, "=")
            If separator < 2 Then Err.Raise 5, , "Invalid journal record"
            result.Add Left$(line, separator - 1), AtWUnhex(Mid$(line, separator + 1))
        End If
    Next line
    Set AtWUnpack = result
End Function

Public Function AtWScalar(ByVal value As Variant) As String
    If IsNull(value) Then AtWScalar = "NULL": Exit Function
    If IsEmpty(value) Then AtWScalar = "EMPTY": Exit Function
    AtWScalar = CStr(VarType(value)) & ":" & CStr(value)
End Function

Private Function RgbFill(ByVal interior As Object, Optional ByVal property As String = "ThemeColor") As Boolean
    On Error GoTo DirectRgb
    Dim theme As Long: theme = CallByName(interior, property, VbGet)
    RgbFill = theme = 0 Or theme = -4142
    Exit Function
DirectRgb:
    ' Excel raises 1004 for a direct RGB fill's ThemeColor; other errors refuse.
    RgbFill = Err.Number = 1004
End Function

Private Function ThreadedMode(ByVal cell As Object) As String
    Dim threaded As Object
    On Error GoTo Unavailable
    Set threaded = cell.CommentThreaded
    ThreadedMode = IIf(threaded Is Nothing, "ABSENT", "PRESENT")
    Exit Function
Unavailable:
    If Err.Number <> 1004 And Err.Number <> 438 Then Err.Raise 5, , "Threaded capability query failed"
    ThreadedMode = "UNAVAILABLE"
End Function

Private Function FontTheme(ByVal font As Object, ByVal property As String) As String
    On Error GoTo Direct
    FontTheme = AtWScalar(CallByName(font, property, VbGet))
    Exit Function
Direct:
    ' Excel's supported getter returns E_INVALIDARG for a direct, non-theme font.
    If Err.Number = 5 And property = "ThemeColor" Then FontTheme = "NO_THEME": Exit Function
    If Err.Number <> 1004 And Err.Number <> -2147024809 Then Err.Raise 5, , "Font theme capability unavailable"
    FontTheme = "NO_THEME"
End Function

Private Function ColorLink(ByVal color As Object) As String
    Dim scheme As String
    If color.Type = 2 Then
        If color.SchemeColor <> 80 And color.SchemeColor <> 81 Then Err.Raise 5, , "Unproven note color scheme"
        scheme = AtWScalar(color.SchemeColor)
    ElseIf color.Type <> 1 Then
        Err.Raise 5, , "Non-RGB/system note shape color unsupported"
    End If
    ColorLink = AtWScalar(color.Type) & "/" & AtWScalar(color.RGB) & "/" & scheme & "/" & AtWScalar(color.TintAndShade)
End Function

' Capture every property this limited writer mutates, plus data/formula guards.
' Existing user notes are refused by the planner, never serialized/restored.
Public Function AtWCapture(ByVal area As Object, ByRef code As String, Optional ByRef reason As String = "") As Object
    Dim result As Object, cell As Object, first As String, profile As String, i As Long
    Dim threaded As Object, note As Object, shape As Object, font As Object, key As Variant, v As Variant
    On Error GoTo Unsupported
    code = "CAPABILITY_UNSUPPORTED"
    reason = "Protection or region size"
    If area.Parent.ProtectContents Or area.Parent.ProtectDrawingObjects Or area.Parent.Parent.ProtectStructure Then Exit Function
    If area.Cells.CountLarge > 64 Then Exit Function
    Set result = AtDictionary()
    result.Add "Sheet", area.Parent.Name: result.Add "CodeName", area.Parent.CodeName: result.Add "Address", area.Address
    result.Add "FileFormat", CStr(area.Parent.Parent.FileFormat)
    result.Add "ThreadedMode", ThreadedMode(area.Cells(1, 1))
    For Each cell In area.Cells
        reason = "Conditional or non-direct fill"
        If cell.FormatConditions.Count > 0 Then Exit Function
        If cell.Interior.Pattern <> -4142 And cell.Interior.Pattern <> 1 Then Exit Function
        If cell.Interior.TintAndShade <> 0 Or cell.Interior.PatternTintAndShade <> 0 Then Exit Function
        If Not RgbFill(cell.Interior) Or Not RgbFill(cell.Interior, "PatternThemeColor") Then Exit Function
        reason = "Threaded comments capability"
        If ThreadedMode(cell) = "PRESENT" Then Exit Function
        profile = CStr(cell.Interior.Pattern) & "/" & CStr(cell.Interior.Color) & "/" & CStr(cell.Interior.PatternColor)
        reason = "Nonuniform merged fill"
        If i = 0 Then first = profile Else If profile <> first Then Exit Function
        i = i + 1
        reason = "Cell data/formula guards"
        result.Add "Data" & i, AtWScalar(cell.Value2)
        result.Add "Formula" & i, AtWScalar(cell.Formula)
        result.Add "Merge" & i, cell.MergeArea.Address
    Next cell
    result.Add "Pattern", CStr(area.Cells(1, 1).Interior.Pattern)
    result.Add "Color", CStr(area.Cells(1, 1).Interior.Color)
    result.Add "PatternColor", CStr(area.Cells(1, 1).Interior.PatternColor)
    reason = "Legacy note capability"
    Set note = area.Cells(1, 1).Comment
    result.Add "Note", "0"
    If Not note Is Nothing Then
        result("Note") = "1": result.Add "NText", note.Text
        If Len(note.Text) > 512 Then Exit Function
        result.Add "NAuthor", note.Author
        Set shape = note.Shape
        If shape.Fill.Type <> 1 Or shape.Shadow.Visible <> 0 Or shape.ThreeD.Visible <> 0 Then Exit Function
        For Each key In Array("Left", "Top", "Width", "Height", "Visible", "Rotation", "Locked", "Placement", "AlternativeText", "OnAction", "Name", "AutoShapeType", "LockAspectRatio")
            result.Add "NShape_" & key, AtWScalar(CallByName(shape, CStr(key), VbGet))
        Next key
        For Each key In Array("AutoSize", "AutoMargins", "MarginLeft", "MarginRight", "MarginTop", "MarginBottom", "HorizontalAlignment", "VerticalAlignment", "Orientation", "ReadingOrder")
            result.Add "NFrame_" & key, AtWScalar(CallByName(shape.TextFrame, CStr(key), VbGet))
        Next key
        Set font = shape.TextFrame.Characters.Font
        For Each key In Array("Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Subscript", "Superscript", "Color", "Shadow", "OutlineFont")
            v = CallByName(font, CStr(key), VbGet)
            If IsNull(v) Then Exit Function
            result.Add "NFont_" & key, AtWScalar(v)
        Next key
        For Each key In Array("ColorIndex", "ThemeColor", "TintAndShade", "ThemeFont")
            reason = "Legacy font linkage " & key
            result.Add "NFontLink_" & key, FontTheme(font, CStr(key))
        Next key
        result.Add "NFill", AtWScalar(shape.Fill.Visible) & "/" & AtWScalar(shape.Fill.Type) & "/" & AtWScalar(shape.Fill.ForeColor.RGB) & "/" & AtWScalar(shape.Fill.BackColor.RGB) & "/" & AtWScalar(shape.Fill.Transparency)
        reason = "Legacy fill color linkage"
        result.Add "NFillLink", ColorLink(shape.Fill.ForeColor) & "/" & ColorLink(shape.Fill.BackColor)
        result.Add "NLine", AtWScalar(shape.Line.Visible) & "/" & AtWScalar(shape.Line.ForeColor.RGB) & "/" & AtWScalar(shape.Line.Weight) & "/" & AtWScalar(shape.Line.DashStyle) & "/" & AtWScalar(shape.Line.Style) & "/" & AtWScalar(shape.Line.Transparency)
        reason = "Legacy line color linkage"
        result.Add "NLineLink", ColorLink(shape.Line.ForeColor) & "/" & ColorLink(shape.Line.BackColor)
        For Each key In Array("Pattern", "BeginArrowheadLength", "BeginArrowheadStyle", "BeginArrowheadWidth", "EndArrowheadLength", "EndArrowheadStyle", "EndArrowheadWidth")
            result.Add "NLine_" & key, AtWScalar(CallByName(shape.Line, CStr(key), VbGet))
        Next key
        For Each key In Array("Visible", "Type", "Transparency", "Obscured", "OffsetX", "OffsetY")
            result.Add "NShadow_" & key, AtWScalar(CallByName(shape.Shadow, CStr(key), VbGet))
        Next key
        result.Add "NZOrder", AtWScalar(shape.ZOrderPosition)
        result.Add "NFlip", AtWScalar(shape.HorizontalFlip) & "/" & AtWScalar(shape.VerticalFlip)
    End If
    If Len(AtWPack(result)) > 65536 Then Exit Function
    code = "": reason = "": Set AtWCapture = result
    Exit Function
Unsupported:
    reason = reason & ": " & CStr(Err.Number) & " " & Err.Description
    ' Unknown property/capability is a refusal, never successful partial capture.
End Function

Public Function AtWWithoutNote(ByVal source As Object) As Object
    Dim result As Object, key As Variant: Set result = AtDictionary()
    For Each key In source.Keys
        If Left$(CStr(key), 1) <> "N" Then result.Add CStr(key), CStr(source(key))
    Next key
    result.Add "Note", "0"
    Set AtWWithoutNote = result
End Function

Public Function AtWEqual(ByVal first As Object, ByVal second As Object) As Boolean
    Dim key As Variant
    If first Is Nothing Or second Is Nothing Then Exit Function
    If first.Count <> second.Count Then Exit Function
    For Each key In first.Keys
        If Not second.Exists(key) Then Exit Function
        If first(key) <> second(key) Then Exit Function
    Next key
    AtWEqual = True
End Function

' Historical inputs are not owned by Writer. Only these exact capture keys
' may differ at reconciliation; missing/extra evidence still fails closed.
Private Function InputGuard(ByVal key As String) As Boolean
    Dim i As Long
    For i = 1 To 64
        If key = "Data" & i Or key = "Formula" & i Then InputGuard = True: Exit Function
    Next i
End Function

Public Function AtWOwnedEqual(ByVal current As Object, ByVal expected As Object) As Boolean
    Dim key As Variant
    If current Is Nothing Or expected Is Nothing Then Exit Function
    If current.Count <> expected.Count Then Exit Function
    For Each key In current.Keys
        If Not expected.Exists(key) Then Exit Function
        If Not InputGuard(CStr(key)) Then
            If current(key) <> expected(key) Then Exit Function
        End If
    Next key
    AtWOwnedEqual = True
End Function

' Before is a restoration/ownership snapshot, never a source of old inputs.
' Full fresh Observed/StepBefore/Expected snapshots retain the input guards.
Public Function AtWAnnotationState(ByVal snapshot As Object) As Object
    Dim result As Object, key As Variant: Set result = AtDictionary()
    For Each key In snapshot.Keys
        If Not InputGuard(CStr(key)) Then result.Add CStr(key), CStr(snapshot(key))
    Next key
    Set AtWAnnotationState = result
End Function

Public Function AtWFillState(ByVal current As Object, ByVal desired As Object) As Object
    Dim result As Object: Set result = AtRecordCopy(current)
    result("Pattern") = desired("Pattern"): result("Color") = desired("Color"): result("PatternColor") = desired("PatternColor")
    Set AtWFillState = result
End Function

Public Sub AtWSetFill(ByVal area As Object, ByVal desired As Object)
    With area.Interior
        .Color = CLng(desired("Color")): .PatternColor = CLng(desired("PatternColor"))
        .TintAndShade = 0: .PatternTintAndShade = 0: .Pattern = CLng(desired("Pattern"))
    End With
End Sub
