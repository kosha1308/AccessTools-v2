VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATSupersetBucket"
Option Explicit
Private mId As String, mCount As Long, mSets As Collection
Public Sub Initialize(ByVal summaryId As String)
    Dim i As Long, value As ATSupersetTextSet
    If Not mSets Is Nothing Or Len(summaryId) = 0 Then Err.Raise vbObjectError + 2820, , "SUPERSET_BUCKET_CONFIG"
    mId = summaryId: Set mSets = New Collection
    For i = 1 To 5
        Set value = New ATSupersetTextSet: mSets.Add value
    Next i
End Sub
Public Sub AddMatch(ByVal values As Variant)
    Dim i As Long
    If mSets Is Nothing Then Err.Raise vbObjectError + 2820, , "SUPERSET_BUCKET_CONFIG"
    If Not IsArray(values) Then Err.Raise vbObjectError + 2820, , "SUPERSET_BUCKET_VALUES"
    If UBound(values) - LBound(values) <> 4 Then Err.Raise vbObjectError + 2820, , "SUPERSET_BUCKET_VALUES"
    For i = LBound(values) To UBound(values)
        If VarType(values(i)) <> vbString Then Err.Raise vbObjectError + 2820, , "SUPERSET_BUCKET_VALUES"
    Next i
    For i = 1 To 5
        mSets(i).AddLines CStr(values(LBound(values) + i - 1))
    Next i
    mCount = mCount + 1
End Sub
Public Function Snapshot() As ATRuntimeRow
    Dim row As New ATRuntimeRow, ids As Variant, i As Long
    If mSets Is Nothing Then Err.Raise vbObjectError + 2820, , "SUPERSET_BUCKET_CONFIG"
    ids = Array("HOSTNAME", "SERVICE", "IS_NAME", "SKUF_GROUP", "SUPPORT_TEAM")
    row.Initialize "SUPERSET_SUMMARY", mId
    row.AddField "MATCH_COUNT", AT_RuntimeOK(AT_RuntimeScalar("NUMBER", CDbl(mCount)))
    For i = 1 To 5
        row.AddField CStr(ids(i - 1)), AT_RuntimeOK(AT_RuntimeScalar("TEXT", mSets(i).Text(i = 1)))
    Next i
    Set Snapshot = row.Snapshot
End Function
