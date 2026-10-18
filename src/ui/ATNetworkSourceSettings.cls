VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkSourceSettings"
Option Explicit
Private mRule As String, mSuperset As String
#If VBA7 Then
Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal source As LongPtr, ByVal target As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function MoveFileExW Lib "kernel32" (ByVal source As Long, ByVal target As Long, ByVal flags As Long) As Long
#End If
Public Function Load(Optional ByVal folder As String = "") As Boolean
    Dim fs As Object, stream As Object, parts As Variant, path As String, text As String
    On Error GoTo Failed
    mRule = "": mSuperset = ""
    Set fs = CreateObject("Scripting.FileSystemObject"): path = SettingsFolder(folder) & "\sources.ini"
    If Not fs.FileExists(path) Then Exit Function
    If fs.GetFile(path).Size > 65536 Then Exit Function
    Set stream = fs.OpenTextFile(path, 1, False, -1): text = stream.ReadAll: stream.Close: Set stream = Nothing
    parts = Split(Replace(text, vbCrLf, vbLf), vbLf)
    If UBound(parts) <> 2 Then Exit Function
    If parts(0) <> "AccessToolsv2Sources/1" Then Exit Function
    mRule = CleanPath(CStr(parts(1)), False): mSuperset = CleanPath(CStr(parts(2)), False)
    If StrComp(mRule, mSuperset, vbTextCompare) = 0 Then GoTo Failed
    Load = True: Exit Function
Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    mRule = "": mSuperset = ""
    Err.Clear
End Function
Public Sub Save(ByVal rulePath As String, ByVal supersetPath As String, Optional ByVal folder As String = "")
    Dim fs As Object, stream As Object, target As String, temporary As String, base As String, code As String
    On Error GoTo Failed
    rulePath = CleanPath(rulePath, True): supersetPath = CleanPath(supersetPath, True)
    If StrComp(rulePath, supersetPath, vbTextCompare) = 0 Then Err.Raise vbObjectError + 3200, , "SOURCE_PATH_COLLISION"
    Set fs = CreateObject("Scripting.FileSystemObject"): base = SettingsFolder(folder)
    If Not fs.FolderExists(base) Then fs.CreateFolder base
    target = fs.BuildPath(base, "sources.ini"): temporary = fs.BuildPath(base, fs.GetTempName)
    Set stream = fs.CreateTextFile(temporary, False, True)
    stream.Write "AccessToolsv2Sources/1" & vbCrLf & rulePath & vbCrLf & supersetPath
    stream.Close: Set stream = Nothing
    If MoveFileExW(StrPtr(temporary), StrPtr(target), 9) = 0 Then Err.Raise vbObjectError + 3200, , "SETTINGS_SAVE_FAILED"
    temporary = "": mRule = rulePath: mSuperset = supersetPath: Exit Sub
Failed:
    code = AT_StableErrorCode(Err.Description, "SETTINGS_SAVE_FAILED")
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If Len(temporary) > 0 Then
        If fs.FileExists(temporary) Then fs.DeleteFile temporary, True
    End If
    On Error GoTo 0
    Err.Raise vbObjectError + 3200, "ATNetworkSourceSettings", code
End Sub
Private Function SettingsFolder(ByVal folder As String) As String
    Dim fs As Object, base As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Len(folder) = 0 Then
        base = Environ$("LOCALAPPDATA")
        If Len(base) = 0 Or Not fs.FolderExists(base) Then Err.Raise vbObjectError + 3200, , "SETTINGS_FOLDER_UNAVAILABLE"
        folder = fs.BuildPath(base, "AccessToolsv2")
    End If
    SettingsFolder = fs.GetAbsolutePathName(folder)
End Function
Private Function CleanPath(ByVal path As String, ByVal requireExists As Boolean) As String
    Dim fs As Object
    If Len(Trim$(path)) = 0 Or InStr(path, vbCr) > 0 Or InStr(path, vbLf) > 0 Or InStr(path, "://") > 0 Then Err.Raise vbObjectError + 3200, , "SOURCE_PATH_INVALID"
    Set fs = CreateObject("Scripting.FileSystemObject")
    path = fs.GetAbsolutePathName(path)
    If LCase$(fs.GetExtensionName(path)) <> "xlsx" Then Err.Raise vbObjectError + 3200, , "SOURCE_FORMAT_UNSUPPORTED"
    If requireExists And Not fs.FileExists(path) Then Err.Raise vbObjectError + 3200, , "SOURCE_FILE_MISSING"
    CleanPath = path
End Function
Public Property Get RulebookPath() As String
    RulebookPath = mRule
End Property
Public Property Get SupersetPath() As String
    SupersetPath = mSuperset
End Property
Public Property Get Available() As Boolean
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Available = Len(mRule) > 0 And Len(mSuperset) > 0 And fs.FileExists(mRule) And fs.FileExists(mSuperset)
End Property
