VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATSupersetTextSet"
Option Explicit
Private mKeys As Object, mValues As Collection, mLength As Long
Private Sub Class_Initialize()
    Set mKeys = CreateObject("Scripting.Dictionary")
    mKeys.CompareMode = vbBinaryCompare
    Set mValues = New Collection
End Sub
Public Sub AddLines(ByVal text As String)
    Dim line As Variant, value As String, key As String, length As Long
    If Len(text) > 32767 Then Err.Raise vbObjectError + 2820, , "SUPERSET_TEXT_LIMIT"
    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    For Each line In Split(text, vbLf)
        value = TrimValue(CStr(line))
        If Len(value) > 0 Then
            key = LCase$(value)
            If Not mKeys.Exists(key) Then
                If mValues.Count >= 10000 Then Err.Raise vbObjectError + 2820, , "SUPERSET_SET_LIMIT"
                length = mLength + Len(value)
                If mValues.Count > 0 Then length = length + 1
                If length > 32767 Then Err.Raise vbObjectError + 2820, , "SUPERSET_TEXT_LIMIT"
                mKeys.Add key, True: mValues.Add value: mLength = length
            End If
        End If
    Next line
End Sub
Private Function TrimValue(ByVal text As String) As String
    Dim first As Long, last As Long
    first = 1: last = Len(text)
    Do While first <= last
        If Not OuterSpace(Mid$(text, first, 1)) Then Exit Do
        first = first + 1
    Loop
    Do While last >= first
        If Not OuterSpace(Mid$(text, last, 1)) Then Exit Do
        last = last - 1
    Loop
    If last >= first Then TrimValue = Mid$(text, first, last - first + 1)
End Function
Private Function OuterSpace(ByVal char As String) As Boolean
    Select Case AscW(char) And &HFFFF&
        Case 9, 32, 160, 8203, 8204, 8205, 65279: OuterSpace = True
    End Select
End Function
Public Function Text(Optional ByVal uniqueOnly As Boolean = False) As String
    Dim value As Variant
    If uniqueOnly And mValues.Count <> 1 Then Exit Function
    For Each value In mValues
        If Len(Text) > 0 Then Text = Text & vbLf
        Text = Text & CStr(value)
    Next value
End Function
Public Property Get Count() As Long
    Count = mValues.Count
End Property
