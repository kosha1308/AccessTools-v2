Attribute VB_Name = "modATSupersetSetTests"
Option Explicit
Public Sub AT_TestSupersetSets()
    Dim values As New ATSupersetTextSet, bucket As New ATSupersetBucket, row As ATRuntimeRow
    Dim snapshot As ATRuntimeRow, caught As Boolean, i As Long, full As New ATSupersetTextSet
    values.AddLines "  Team A " & vbCrLf & "team a" & vbCr & ChrW(160) & "Team B" & ChrW(160)
    values.AddLines vbTab & vbLf & "" & ChrW(&H200B)
    AT_Assert values.Count = 2 And values.Text = "Team A" & vbLf & "Team B", "Superset values unique case-insensitive ordered trimmed lines"
    AT_Assert values.Text(True) = "", "multiple hostname values ambiguous"
    Set values = New ATSupersetTextSet
    values.AddLines "host.example": values.AddLines "HOST.EXAMPLE"
    AT_Assert values.Text(True) = "host.example", "one unique hostname preserved"
    Set values = New ATSupersetTextSet
    values.AddLines "Группа А" & vbLf & "группа а"
    AT_Assert values.Count = 1 And values.Text = "Группа А", "Cyrillic values deduplicated without changing first spelling"
    On Error Resume Next
    values.AddLines String$(32768, "x")
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught, "Superset joined text overflow rejected"
    For i = 1 To 10000
        full.AddLines ChrW(&H4E00 + (i Mod 1000)) & ChrW(&H4E00 + (i \ 1000))
    Next i
    AT_Assert full.Count = 10000, "exact unique text set count boundary"
    On Error Resume Next
    full.AddLines "extra"
    caught = Err.Number <> 0: Err.Clear
    On Error GoTo 0
    AT_Assert caught, "Superset unique text set bounded"
    bucket.Initialize "1"
    Set row = bucket.Snapshot
    AT_Assert row.RowId = "1" And row.ResolveField("MATCH_COUNT").Value.Scalar = 0#, "zero-match summary retained"
    AT_Assert row.ResolveField("SKUF_GROUP").Value.ValueState = atValueEmpty, "no-match Superset enrichment empty"
    bucket.AddMatch Array("one.example", "svc1", "system1", "group1", "team1")
    Set snapshot = bucket.Snapshot
    bucket.AddMatch Array("two.example", "svc2", "system2", "group2", "team2")
    Set row = bucket.Snapshot
    AT_Assert snapshot.ResolveField("MATCH_COUNT").Value.Scalar = 1#, "bucket snapshot immutable after more matches"
    AT_Assert row.ResolveField("MATCH_COUNT").Value.Scalar = 2#, "distinct source rows counted"
    AT_Assert row.ResolveField("HOSTNAME").Value.ValueState = atValueEmpty, "ambiguous summary hostname left empty"
    AT_Assert row.ResolveField("SKUF_GROUP").Value.Scalar = "group1" & vbLf & "group2", "revised SKUF field aggregates all values"
    AT_Assert row.ResolveField("SUPPORT_TEAM").Value.Scalar = "team1" & vbLf & "team2", "revised support field aggregates all values"
End Sub
