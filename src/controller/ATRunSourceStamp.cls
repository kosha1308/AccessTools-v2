VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATRunSourceStamp"
Option Explicit
Private mReady As Boolean, mPath As String, mFingerprint As String
Public Sub Initialize(ByVal path As String)
    Dim fs As Object, extension As String
    If mReady Then Err.Raise vbObjectError + 3180, , "SOURCE_STAMP_FROZEN"
    If Len(Trim$(path)) = 0 Or InStr(path, "://") > 0 Then Err.Raise vbObjectError + 3180, , "SOURCE_PATH_INVALID"
    Set fs = CreateObject("Scripting.FileSystemObject"): mPath = fs.GetAbsolutePathName(path)
    extension = LCase$(fs.GetExtensionName(mPath))
    If extension <> "xlsx" Then Err.Raise vbObjectError + 3180, , "SOURCE_FORMAT_UNSUPPORTED"
    mFingerprint = CurrentFingerprint(): mReady = True
End Sub
Private Function CurrentFingerprint() As String
    Dim fs As Object, file As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(mPath) Then Err.Raise vbObjectError + 3180, , "SOURCE_FILE_MISSING"
    Set file = fs.GetFile(mPath)
    CurrentFingerprint = LCase$(mPath) & "|" & Trim$(Str$(CDbl(file.Size))) & "|" & Trim$(Str$(CDbl(file.DateLastModified)))
End Function
Public Sub VerifyCurrent()
    If Not mReady Then Err.Raise vbObjectError + 3180, , "SOURCE_STAMP_REQUIRED"
    If CurrentFingerprint() <> mFingerprint Then Err.Raise vbObjectError + 3180, , "SOURCE_CHANGED_DURING_RUN"
End Sub
Public Property Get Path() As String
    If Not mReady Then Err.Raise vbObjectError + 3180, , "SOURCE_STAMP_REQUIRED"
    Path = mPath
End Property
Public Property Get Fingerprint() As String
    If Not mReady Then Err.Raise vbObjectError + 3180, , "SOURCE_STAMP_REQUIRED"
    Fingerprint = mFingerprint
End Property
