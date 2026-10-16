Attribute VB_Name = "modATPackageTests"
Option Explicit
Public Sub AT_TestPackage()
    Dim fs As Object, folder As String, stream As Object, message As String
    On Error GoTo Failed
    Set fs = CreateObject("Scripting.FileSystemObject")
    folder = fs.BuildPath(fs.GetSpecialFolder(2), "AccessTools-package-" & fs.GetTempName): fs.CreateFolder folder
    AT_Assert AT_PackageCheck(folder, "v", "r", "COORDINATOR", "c") = "PACKAGE_INCOMPLETE", "missing package metadata rejects execution"
    fs.CreateFolder folder & "\config"
    Set stream = fs.CreateTextFile(folder & "\VERSION.txt"): stream.Write "v": stream.Close
    Set stream = fs.CreateTextFile(folder & "\RELEASE_INFO.txt"): stream.Write "Version=v" & vbLf & "ReleaseId=r" & vbLf & "RuntimeProfile=COORDINATOR" & vbLf & "Commit=c": stream.Close
    Set stream = fs.CreateTextFile(folder & "\config\runtime.ini"): stream.Write "SchemaVersion=1.0.0" & vbLf & "RuntimeProfile=COORDINATOR": stream.Close
    AT_Assert Len(AT_PackageCheck(folder, "v", "r", "COORDINATOR", "c")) = 0, "matching coordinator package metadata accepted"
    AT_Assert AT_PackageCheck(folder, "v2", "r", "COORDINATOR", "c") = "PACKAGE_MISMATCH", "mixed installed version rejected"
    AT_Assert AT_PackageCheck(folder, "v", "r", "FIRST_LINE", "c") = "PACKAGE_PROFILE_UNSUPPORTED", "coordinator pilot cannot masquerade as first line"
    Set stream = fs.CreateTextFile(folder & "\config\runtime.ini", True): stream.Write "SchemaVersion=2.0.0" & vbLf & "RuntimeProfile=COORDINATOR": stream.Close
    AT_Assert AT_PackageCheck(folder, "v", "r", "COORDINATOR", "c") = "PACKAGE_MISMATCH", "mixed schema configuration rejected"
Cleanup:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If Len(folder) > 0 Then
        If StrComp(fs.GetParentFolderName(folder), CStr(fs.GetSpecialFolder(2)), vbTextCompare) = 0 And Left$(fs.GetFileName(folder), 20) = "AccessTools-package-" Then fs.DeleteFolder folder, True
    End If
    On Error GoTo 0
    Exit Sub
Failed:
    message = Err.Description: AT_Assert False, "package unexpected error: " & message: Resume Cleanup
End Sub
