VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtJournalStore"
Option Explicit

Private root As String, created As Object

Friend Sub Initialize(ByVal directory As String)
    If Not created Is Nothing Then Err.Raise 5, , "Store already initialized"
    Dim fso As Object, parent As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If directory = "" Then
        parent = Environ$("LOCALAPPDATA")
        AssertLocal fso, parent
        parent = fso.BuildPath(parent, "AccessTools")
        If Not fso.FolderExists(parent) Then fso.CreateFolder parent
        AssertLocal fso, parent
        directory = fso.BuildPath(parent, "Journals")
        If Not fso.FolderExists(directory) Then fso.CreateFolder directory
    End If
    AssertLocal fso, directory
    root = fso.GetAbsolutePathName(directory): Set created = AtDictionary()
End Sub

Private Sub AssertLocal(ByVal fso As Object, ByVal directory As String)
    Dim path As String
    If Mid$(directory, 2, 2) <> ":\" Then Err.Raise 5, , "Absolute local directory required"
    path = fso.GetAbsolutePathName(directory)
    If Mid$(path, 2, 2) <> ":\" Or Not fso.FolderExists(path) Then Err.Raise 5, , "Existing local directory required"
    If fso.GetDrive(fso.GetDriveName(path)).DriveType <> 2 Then Err.Raise 5, , "Fixed local directory required"
    Do While Len(path) > 3
        If (fso.GetFolder(path).Attributes And 1024) <> 0 Then Err.Raise 5, , "Reparse directories unsupported"
        path = fso.GetParentFolderName(path)
    Loop
End Sub

Public Function NewDirectory() As String
    Dim fso As Object, path As String
    Set fso = CreateObject("Scripting.FileSystemObject"): AssertLocal fso, root
    path = fso.BuildPath(root, "AT_" & fso.GetTempName)
    ' CreateFolder refuses a collision; ownership is never adopted from disk.
    fso.CreateFolder path: created.Add path, True: NewDirectory = path
End Function

Public Function PruneCompleted(ByVal journal As CAtOwnershipJournal, ByVal directory As String) As Boolean
    Dim fso As Object, manifest As Object, folder As Object, file As Object, key As Variant
    On Error GoTo KeepEvidence
    If Not created.Exists(directory) Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    AssertLocal fso, directory
    If StrComp(fso.GetParentFolderName(directory), root, vbTextCompare) <> 0 Then Exit Function
    Set manifest = journal.CompletedManifest()
    If manifest Is Nothing Then Exit Function
    Set folder = fso.GetFolder(directory)
    If folder.SubFolders.Count <> 0 Or folder.Files.Count <> manifest.Count Then Exit Function
    For Each key In manifest.Keys
        If StrComp(fso.GetParentFolderName(CStr(key)), directory, vbTextCompare) <> 0 Then Exit Function
        If Not ExactFile(fso, CStr(key), CStr(manifest(key))) Then Exit Function
    Next key
    For Each key In manifest.Keys
        If Not ExactFile(fso, CStr(key), CStr(manifest(key))) Then Exit Function
        fso.DeleteFile CStr(key), False
    Next key
    ' RmDir is nonrecursive: any new/unknown file prevents folder deletion.
    RmDir directory
    created.Remove directory: PruneCompleted = True
KeepEvidence:
    ' No age-based purge, cross-session adoption, or recovery by guessed state.
End Function

Private Function ExactFile(ByVal fso As Object, ByVal path As String, ByVal text As String) As Boolean
    Dim stream As Object, actual As String
    If Not fso.FileExists(path) Then Exit Function
    If (fso.GetFile(path).Attributes And 1024) <> 0 Then Exit Function
    If fso.GetFile(path).Size <> 2 + 2 * Len(text) Then Exit Function
    Set stream = fso.OpenTextFile(path, 1, False, -1)
    actual = stream.ReadAll: stream.Close: ExactFile = actual = text
End Function
