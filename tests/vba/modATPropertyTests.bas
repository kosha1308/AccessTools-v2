Attribute VB_Name = "modATPropertyTests"
Option Explicit

Public Sub AT_TestProperties()
    Dim parser As New ATParser, serializer As New ATSerializer
    Dim result As ATParseResult, other As ATParseResult
    Dim text As String, canonical As String, alphabet As String
    Dim index As Long, position As Long, seed As Long
    alphabet = "@$.(),:[]+-*/%=<> ABCxyz012" & Chr$(34) & "\"
    seed = 1337
    For index = 1 To 300
        text = ""
        For position = 1 To 20
            seed = (seed * 251 + 17) Mod 32749
            text = text & Mid$(alphabet, (seed Mod Len(alphabet)) + 1, 1)
        Next position
        Set result = parser.Parse(text)
        AT_Assert Not result Is Nothing, "fuzz returns result"
        If result.Diagnostics.Count = 0 Then
            canonical = serializer.Serialize(result.Root)
            Set other = parser.Parse(canonical)
            AT_Assert other.Diagnostics.Count = 0, "fuzz valid round trip"
            AT_Assert canonical = serializer.Serialize(other.Root), "fuzz canonical stable"
        Else
            AT_Assert result.Root Is Nothing, "fuzz malformed input has no executable root"
        End If
    Next index
    For index = 1 To 30
        text = "(@A + " & CStr(index) & ") > 10 AND NOT (@B = 0 OR FALSE)"
        Set result = parser.Parse(text)
        canonical = serializer.Serialize(result.Root)
        Set other = parser.Parse(canonical)
        AT_Assert other.Diagnostics.Count = 0, "generated valid expression round trip"
        AT_Assert canonical = serializer.Serialize(other.Root), "generated canonical stable"
    Next index
End Sub
