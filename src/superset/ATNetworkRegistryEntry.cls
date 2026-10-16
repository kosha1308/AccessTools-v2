VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkRegistryEntry"
Option Explicit
Private mInitialized As Boolean, mReady As Boolean, mId As String, mAddress As String, mKey As String, mDomain As ATIntervalSet
Private mRefs As Collection, mRefKeys As Object, mHosts As Object, mReferences As ATRuntimeValue, mMatches As ATRuntimeValue
Private mStatus As String, mComplete As Boolean, mRefined As Boolean, mRequestHostIncomplete As Boolean
Friend Sub Initialize(ByVal address As String, ByVal key As String, ByVal domain As ATIntervalSet)
    If mInitialized Then Err.Raise vbObjectError + 3010, , "REGISTRY_FROZEN"
    If Not domain.Frozen Then Err.Raise vbObjectError + 3010, , "REGISTRY_INPUT_INVALID"
    mAddress = address: mKey = key: Set mDomain = domain
    Set mRefs = New Collection: Set mRefKeys = CreateObject("Scripting.Dictionary"): Set mHosts = CreateObject("Scripting.Dictionary")
    mInitialized = True
End Sub
Friend Sub AddReference(ByVal field As String, ByVal ref As String, ByVal budget As ATEvalBudget)
    Dim row As New ATRuntimeRow, key As String
    If Not mInitialized Or mReady Then Err.Raise vbObjectError + 3010, , "REGISTRY_FROZEN"
    If Len(ref) = 0 Then Err.Raise vbObjectError + 3010, , "REGISTRY_INPUT_INVALID"
    budget.CheckText Len(ref): budget.Step Len(ref) + 1: key = field & "|" & ref
    If mRefKeys.Exists(key) Then Exit Sub
    budget.CheckList mRefs.Count + 1: mRefKeys.Add key, True
    row.Initialize "REGISTRY_REFERENCE", CStr(mRefs.Count + 1)
    row.AddField "FIELD_ID", AT_RuntimeOK(AT_RuntimeScalar("TEXT", field))
    row.AddField "CELL_REF", AT_RuntimeOK(AT_RuntimeScalar("TEXT", ref)), ref
    mRefs.Add row
End Sub
Friend Sub AddHostname(ByVal text As String, ByVal budget As ATEvalBudget, Optional ByVal valid As Boolean = True)
    Dim key As String
    If Not mInitialized Or mReady Then Err.Raise vbObjectError + 3010, , "REGISTRY_FROZEN"
    If Not valid Then mRequestHostIncomplete = True: Exit Sub
    budget.CheckText Len(text): budget.Step Len(text) + 1: key = AT_RegistryHost(text)
    If Len(key) > 0 And Not mHosts.Exists(key) Then
        budget.CheckList mHosts.Count + 1: mHosts.Add key, text
    End If
End Sub
Friend Property Get QueryText() As String
    If Not mInitialized Then Err.Raise vbObjectError + 3010, , "REGISTRY_ACCESS"
    QueryText = mAddress
End Property
Friend Sub Finish(ByVal id As String, ByVal raw As ATRuntimeValue, ByVal complete As Boolean, ByRef hostIncomplete As Boolean, ByVal budget As ATEvalBudget)
    Dim selected As New Collection, row As ATRuntimeRow, i As Long, valid As Boolean, host As String, narrow As Boolean
    If Not mInitialized Or mReady Then Err.Raise vbObjectError + 3010, , "REGISTRY_FROZEN"
    If mRefs.Count = 0 Or raw.DataType <> "COLLECTION<ROW<SUPERSET_RAW>>" Then Err.Raise vbObjectError + 3010, , "REGISTRY_INPUT_INVALID"
    budget.CheckList raw.Count: budget.Step raw.Count + mRefs.Count * 4
    hostIncomplete = False
    If mHosts.Count > 0 And mDomain.Count = 1 Then
        If mDomain.Item(1).Low = mDomain.Item(1).High Then
            For i = 1 To raw.Count
                Set row = raw.Row(i): host = AT_RegistryText(row.ResolveField("HOSTNAME", budget), valid, budget)
                If Not valid Then hostIncomplete = True
                If valid Then
                    If mHosts.Exists(AT_RegistryHost(host)) Then selected.Add row
                End If
            Next i
            narrow = selected.Count > 0 And Not hostIncomplete And Not mRequestHostIncomplete
        End If
    End If
    If narrow Then
        Set mMatches = New ATRuntimeValue: mMatches.InitializeRows "SUPERSET_RAW", selected
    Else
        Set mMatches = raw.Clone
    End If
    Set mReferences = New ATRuntimeValue: mReferences.InitializeRows "REGISTRY_REFERENCE", mRefs
    mComplete = complete And Not hostIncomplete: mRefined = narrow: mId = id
    mStatus = "FOUND"
    If mMatches.Count = 0 Then
        mStatus = "UNCONFIRMED"
        If mComplete Then mStatus = "NOT_FOUND"
    End If
    mReady = True
End Sub
Private Sub RequireReady()
    If Not mReady Then Err.Raise vbObjectError + 3010, , "REGISTRY_ACCESS"
End Sub
Public Property Get Id() As String
    RequireReady: Id = mId
End Property
Public Property Get Address() As String
    RequireReady: Address = mAddress
End Property
Public Property Get Key() As String
    RequireReady: Key = mKey
End Property
Public Property Get References() As ATRuntimeValue
    RequireReady: Set References = mReferences.Clone
End Property
Public Property Get Matches() As ATRuntimeValue
    RequireReady: Set Matches = mMatches.Clone
End Property
Public Property Get Status() As String
    RequireReady: Status = mStatus
End Property
Public Property Get HostnameRefined() As Boolean
    RequireReady: HostnameRefined = mRefined
End Property
Public Property Get SourceComplete() As Boolean
    RequireReady: SourceComplete = mComplete
End Property
Public Property Get Domain() As ATIntervalSet
    RequireReady: Set Domain = mDomain
End Property
Public Property Get Hostnames() As Collection
    Dim r As New Collection, value As Variant
    RequireReady
    For Each value In mHosts.Items: r.Add CStr(value): Next value
    Set Hostnames = r
End Property
