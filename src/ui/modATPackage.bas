Attribute VB_Name = "modATPackage"
Option Explicit
Public Const AT_PACKAGE_VERSION As String = "DEV"
Public Const AT_PACKAGE_RELEASE As String = "DEV"
Public Const AT_PACKAGE_PROFILE As String = "DEV"
Public Const AT_PACKAGE_COMMIT As String = "DEV"
Public Function AT_PackageCheck(ByVal folder As String, ByVal version As String, ByVal release As String, ByVal profile As String, ByVal commit As String) As String
    Dim fs As Object, expected As String
    On Error GoTo Failed
    If profile <> "COORDINATOR" Then AT_PackageCheck = "PACKAGE_PROFILE_UNSUPPORTED": Exit Function
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(folder & "\VERSION.txt") Or Not fs.FileExists(folder & "\RELEASE_INFO.txt") Or Not fs.FileExists(folder & "\config\runtime.ini") Then AT_PackageCheck = "PACKAGE_INCOMPLETE": Exit Function
    expected = "Version=" & version & vbLf & "ReleaseId=" & release & vbLf & "RuntimeProfile=" & profile & vbLf & "Commit=" & commit
    If ReadMarker(fs, folder & "\VERSION.txt") <> version Or ReadMarker(fs, folder & "\RELEASE_INFO.txt") <> expected Then AT_PackageCheck = "PACKAGE_MISMATCH": Exit Function
    expected = "SchemaVersion=" & AT_SCHEMA_VERSION & vbLf & "RuntimeProfile=" & profile
    If ReadMarker(fs, folder & "\config\runtime.ini") <> expected Then AT_PackageCheck = "PACKAGE_MISMATCH"
    Exit Function
Failed:
    AT_PackageCheck = "PACKAGE_READ_FAILED"
End Function
Private Function ReadMarker(ByVal fs As Object, ByVal path As String) As String
    Dim stream As Object, text As String
    Set stream = fs.OpenTextFile(path, 1, False)
    text = stream.ReadAll: stream.Close
    text = Replace(text, vbCrLf, vbLf)
    Do While Right$(text, 1) = vbLf
        text = Left$(text, Len(text) - 1)
    Loop
    ReadMarker = text
End Function
Public Function AT_PackageInstalledCode() As String
    AT_PackageInstalledCode = AT_PackageCheck(ThisWorkbook.Path, AT_PACKAGE_VERSION, AT_PACKAGE_RELEASE, AT_PACKAGE_PROFILE, AT_PACKAGE_COMMIT)
End Function
Public Function AT_PackageIdentity() As String
    AT_PackageIdentity = AT_PACKAGE_VERSION & "|" & AT_PACKAGE_RELEASE & "|" & AT_PACKAGE_PROFILE & "|" & AT_PACKAGE_COMMIT
End Function
Public Function AT_PackageAllowCommand() As Boolean
    Dim code As String
    If AT_PackageIdentity() = "DEV|DEV|DEV|DEV" Then AT_PackageAllowCommand = True: Exit Function
    code = AT_PackageInstalledCode()
    If Len(code) = 0 Then AT_PackageAllowCommand = True: Exit Function
    MsgBox "Комплект AccessTools неполный или файлы относятся к разным версиям. Распакуйте весь комплект в одну папку и подключите надстройку из неё." & vbCrLf & "Код: " & code, vbExclamation, "AccessTools"
End Function
