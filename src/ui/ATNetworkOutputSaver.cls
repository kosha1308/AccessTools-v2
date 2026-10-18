VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ATNetworkOutputSaver"
Option Explicit
Public Function Save(ByVal reportBook As Workbook, ByVal coordinatorBook As Workbook, ByVal parentFolder As String, Optional ByVal includeReport As Boolean = True) As String
    Dim fs As Object, folder As String, prefix As String, reportPath As String, coordinatorPath As String
    Dim attempt As Long, owned As Boolean, previousEvents As Boolean, armed As Boolean, code As String
    On Error GoTo Failed
    If includeReport Then RequireOutput reportBook
    RequireOutput coordinatorBook
    If includeReport Then
        If reportBook Is coordinatorBook Then Err.Raise vbObjectError + 3190, , "OUTPUT_BOOK_COLLISION"
    End If
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Len(Trim$(parentFolder)) = 0 Or InStr(parentFolder, "://") > 0 Then Err.Raise vbObjectError + 3190, , "OUTPUT_FOLDER_REQUIRED"
    parentFolder = fs.GetAbsolutePathName(parentFolder)
    If Not fs.FolderExists(parentFolder) Then Err.Raise vbObjectError + 3190, , "OUTPUT_FOLDER_MISSING"
    prefix = "AccessTools_" & Format$(Now, "yyyymmdd_hhnnss")
    For attempt = 1 To 999
        folder = fs.BuildPath(parentFolder, prefix & "_" & Format$(attempt, "000"))
        If Not fs.FolderExists(folder) And Not fs.FileExists(folder) Then Exit For
    Next attempt
    If attempt > 999 Then Err.Raise vbObjectError + 3190, , "OUTPUT_FOLDER_COLLISION"
    ' Ownership starts only after CreateFolder succeeds, including race collisions.
    fs.CreateFolder folder: owned = True
    reportPath = fs.BuildPath(folder, "Отчёт-проверок.xlsx")
    coordinatorPath = fs.BuildPath(folder, "Сводка-координатора.xlsx")
    If Not includeReport Then coordinatorPath = fs.BuildPath(folder, "Superset-с-результатами.xlsx")
    previousEvents = Application.EnableEvents: armed = True: Application.EnableEvents = False
    If includeReport Then reportBook.SaveCopyAs reportPath
    coordinatorBook.SaveCopyAs coordinatorPath
    If includeReport Then
        If Not fs.FileExists(reportPath) Then Err.Raise vbObjectError + 3190, , "OUTPUT_COPY_MISSING"
    End If
    If Not fs.FileExists(coordinatorPath) Then Err.Raise vbObjectError + 3190, , "OUTPUT_COPY_MISSING"
    Application.EnableEvents = previousEvents: armed = False
    Save = folder: Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "OUTPUT_SAVE_FAILED")
    On Error Resume Next
    If armed Then Application.EnableEvents = previousEvents
    If owned Then
        If fs.FileExists(reportPath) Then fs.DeleteFile reportPath, True
        If fs.FileExists(coordinatorPath) Then fs.DeleteFile coordinatorPath, True
        ' Never recurse: a foreign file added meanwhile prevents folder removal.
        RmDir folder
    End If
    On Error GoTo 0
    Err.Raise vbObjectError + 3190, "ATNetworkOutputSaver", code
End Function
Private Sub RequireOutput(ByVal book As Workbook)
    Dim candidate As Workbook, found As Boolean
    If book Is Nothing Then Err.Raise vbObjectError + 3190, , "OUTPUT_BOOK_REQUIRED"
    For Each candidate In Application.Workbooks
        If candidate Is book Then found = True: Exit For
    Next candidate
    If Not found Then Err.Raise vbObjectError + 3190, , "OUTPUT_BOOK_CLOSED"
    If book.FileFormat <> 51 And book.FileFormat <> 61 Then Err.Raise vbObjectError + 3190, , "OUTPUT_FORMAT_UNSUPPORTED"
End Sub
