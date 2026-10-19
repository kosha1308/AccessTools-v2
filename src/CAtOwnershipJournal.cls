VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtOwnershipJournal"
Option Explicit

Private mBook As Object, mFolder As String, mSession As String, mRecovery As Boolean
Private mEntries As Object, mFiles As Object, mSequence As Long, mBytes As Long
Private mRecovering As Boolean, mNormalRevisions As Long

Friend Sub Initialize(ByVal book As Object, ByVal directory As String, ByVal normalRevisions As Long)
    Dim fso As Object, file As Object, path As String, stream As Object, text As String
    If Not mEntries Is Nothing Then Err.Raise 5, , "Journal identity cannot be rebound"
    Set mEntries = AtDictionary(): Set mFiles = AtDictionary(): Set mBook = book
    Set fso = CreateObject("Scripting.FileSystemObject")
    If normalRevisions < 1 Or normalRevisions > 256 Then Err.Raise 5, , "Invalid journal revision budget"
    mNormalRevisions = normalRevisions
    If book Is Nothing Then Err.Raise 5, , "Explicit workbook required"
    If TypeName(book) <> "Workbook" Then Err.Raise 5, , "Explicit workbook required"
    If Not fso.FolderExists(directory) Then Err.Raise 5, , "Existing dedicated journal directory required"
    mFolder = fso.GetAbsolutePathName(directory)
    If Left$(mFolder, 2) = "\\" Or Mid$(mFolder, 2, 2) <> ":\" Then Err.Raise 5, , "Local journal directory required"
    If fso.GetDrive(fso.GetDriveName(mFolder)).DriveType <> 2 Then Err.Raise 5, , "Fixed local drive required; mapped/network storage unsupported"
    mSession = UCase$(Replace(fso.GetTempName(), ".", "_"))
    For Each file In fso.GetFolder(mFolder).Files
        If LCase$(fso.GetExtensionName(file.Name)) = "atj" Then mRecovery = True
    Next file
    If mRecovery Then Exit Sub
    ' Fixed-name lease prevents two live journals claiming the same empty folder.
    path = mFolder & "\session.atj": text = "AccessTools ownership session " & mSession
    Set stream = fso.CreateTextFile(path, False, True): stream.Write text: stream.Close
    Set stream = fso.OpenTextFile(path, 1, False, -1)
    If stream.ReadAll <> text Then stream.Close: Err.Raise 5, , "Journal lease verification failed"
    stream.Close: mFiles.Add path, text
End Sub
Public Property Get SessionId() As String
    SessionId = mSession
End Property
Public Property Get RecoveryRequired() As Boolean
    RecoveryRequired = mRecovery
End Property
Public Property Get EntryCount() As Long
    EntryCount = mEntries.Count
End Property
Public Property Get RevisionCount() As Long
    RevisionCount = mSequence
End Property
Friend Sub RequireRecovery()
    mRecovery = True
End Sub
Friend Sub Recovered()
    mRecovery = False
End Sub
Friend Property Let Recovering(ByVal value As Boolean)
    mRecovering = value
End Property
Friend Property Get Workbook() As Object
    Set Workbook = mBook
End Property
Friend Function BoundTo(ByVal book As Object) As Boolean
    BoundTo = book Is mBook
End Function
Friend Function Entry(ByVal target As String) As Object
    If mEntries.Exists(target) Then Set Entry = AtRecordCopy(mEntries(target))
End Function
Friend Property Get Entries() As Collection
    Dim result As New Collection, key As Variant
    For Each key In mEntries.Keys: result.Add AtRecordCopy(mEntries(key)): Next key
    Set Entries = result
End Property
Friend Function Verify() As Boolean
    Dim fso As Object, key As Variant, stream As Object, file As Object
    On Error GoTo Broken
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.GetFolder(mFolder).Files.Count > 1025 Then GoTo Broken
    For Each file In fso.GetFolder(mFolder).Files
        If LCase$(fso.GetExtensionName(file.Name)) = "atj" Then
            If Not mFiles.Exists(file.Path) Then GoTo Broken
        End If
    Next file
    For Each key In mFiles.Keys
        If fso.GetFile(CStr(key)).Size <> 2 + 2 * Len(CStr(mFiles(key))) Then GoTo Broken
        Set stream = fso.OpenTextFile(CStr(key), 1, False, -1)
        If stream.ReadAll <> mFiles(key) Then stream.Close: GoTo Broken
        stream.Close
    Next key
    Verify = True: Exit Function
Broken:
    mRecovery = True
End Function

' Append-only, close and read back BEFORE the caller may mutate Excel.
Friend Sub Persist(ByVal entry As Object)
    Dim fso As Object, stream As Object, path As String, xml As String
    On Error GoTo Broken
    If Not Verify() Then Err.Raise 5, , "Journal changed or unavailable"
    entry("BookName") = CStr(mBook.Name): entry("BookPath") = CStr(mBook.Path)
    entry("BookFormat") = CStr(mBook.FileFormat): entry("BookSession") = mSession
    If mEntries.Count >= 32 And Not mEntries.Exists(entry("Target")) Then Err.Raise 5, , "Journal target budget exceeded"
    If mRecovering Then
        If mSequence >= 1024 Then Err.Raise 5, , "Recovery revision reserve exhausted"
    Else
        If mSequence >= mNormalRevisions Then Err.Raise 5, , "Journal revision budget exceeded"
    End If
    xml = "<?xml version=""1.0"" encoding=""UTF-16""?><ownership format=""1"" session=""" & mSession & """ sequence=""" & CStr(mSequence + 1) & """>" & AtWPack(entry) & "</ownership>"
    If Len(xml) > 2097152 Then Err.Raise 5, , "Journal record budget exceeded"
    If mRecovering Then
        If mBytes + Len(xml) > 67108864 Then Err.Raise 5, , "Recovery size reserve exhausted"
    Else
        If mBytes + Len(xml) > 8388608 Then Err.Raise 5, , "Journal size budget exceeded"
    End If
    path = mFolder & "\ATJ_" & mSession & "_" & Format$(mSequence + 1, "0000") & ".atj"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(path, False, True): stream.Write xml: stream.Close
    Set stream = fso.OpenTextFile(path, 1, False, -1)
    If stream.ReadAll <> xml Then stream.Close: Err.Raise 5, , "Journal verification failed"
    stream.Close
    mSequence = mSequence + 1: mBytes = mBytes + Len(xml)
    mFiles.Add path, xml
    If mEntries.Exists(entry("Target")) Then mEntries.Remove entry("Target")
    mEntries.Add entry("Target"), AtRecordCopy(entry)
    Exit Sub
Broken:
    mRecovery = True
    Err.Raise 5, , "Journal persistence failed; Excel success is not acknowledged"
End Sub
