Attribute VB_Name = "modATRunSources"
Option Explicit
Public Sub AT_RunAssertNew(ByVal book As Workbook, ByVal previous As Collection)
    Dim item As Workbook
    If book Is Nothing Or previous Is Nothing Then Err.Raise vbObjectError + 3180, , "SOURCE_OWNERSHIP_INVALID"
    For Each item In previous
        If book Is item Then Err.Raise vbObjectError + 3180, , "SOURCE_ALREADY_OPEN"
    Next item
End Sub
Public Sub AT_RunRejectOpen(ByVal path As String)
    Dim book As Workbook, fs As Object, file As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FileExists(path) Then Err.Raise vbObjectError + 3180, , "SOURCE_FILE_MISSING"
    Set file = fs.GetFile(path)
    For Each book In Application.Workbooks
        If StrComp(book.Name, CStr(file.Name), vbTextCompare) = 0 Or StrComp(book.Name, CStr(file.ShortName), vbTextCompare) = 0 Then Err.Raise vbObjectError + 3180, , "SOURCE_ALREADY_OPEN"
        If Len(book.Path) > 0 Then
            If StrComp(book.FullName, path, vbTextCompare) = 0 Then Err.Raise vbObjectError + 3180, , "SOURCE_ALREADY_OPEN"
        End If
    Next book
End Sub
Public Function AT_RunOpenOwned(ByVal path As String) As Workbook
    Dim previous As New Collection, item As Workbook, book As Workbook, owned As Boolean, code As String
    On Error GoTo Failed
    AT_RunRejectOpen path
    For Each item In Application.Workbooks: previous.Add item: Next item
    Set book = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, Password:="", WriteResPassword:="", IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    AT_RunAssertNew book, previous
    If book.FileFormat <> 51 And book.FileFormat <> 61 Then Err.Raise vbObjectError + 3180, , "SOURCE_FORMAT_UNSUPPORTED"
    If Not book.ReadOnly Then Err.Raise vbObjectError + 3180, , "SOURCE_READONLY_REQUIRED"
    Set AT_RunOpenOwned = book: Exit Function
Failed:
    code = AT_StableErrorCode(Err.Description, "SOURCE_OPEN_FAILED")
    On Error Resume Next
    If Not book Is Nothing Then
        owned = True
        For Each item In previous
            If book Is item Then owned = False
        Next item
        If owned Then book.Close False
    End If
    On Error GoTo 0
    Err.Raise vbObjectError + 3180, "AT_RunOpenOwned", code
End Function
Public Function AT_RunReadSource(ByVal book As Workbook, ByVal sheetName As String, ByVal sourceId As String) As ATEvalResult
    Dim ws As Worksheet, reader As New ATNetworkSourceReader, result As ATEvalResult
    On Error Resume Next
    Set ws = book.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Err.Raise vbObjectError + 3180, , "SOURCE_SHEET_MISSING"
    Set result = reader.Read(ws, sourceId)
    If result.State <> atEvalOK Then Err.Raise vbObjectError + 3180, , result.Diagnostics(1).Code
    Set AT_RunReadSource = result
End Function
Public Function AT_RunFailure(ByVal code As String, ByVal phase As String) As ATCheckReport
    Dim report As New ATCheckReport, issues As New Collection, checks As New Collection
    Dim issue As New ATCheckIssue, outcome As New ATCheckOutcome, message As String
    message = "Запуск не завершён. Проверьте заявку и файлы источников. Этап: " & phase & ". Код: " & code & "."
    If code = "SOURCE_ALREADY_OPEN" Then message = "Файл Rulebook/Superset или книга с таким именем уже открыты. Сохраните нужные изменения, закройте книгу и повторите запуск. Пользовательская книга не изменялась."
    issue.Initialize "NET.RUN_CONTROL", code, "ERROR", "", "", message: issues.Add issue
    outcome.Initialize "NET.RUN_CONTROL", "BLOCKED", message: checks.Add outcome
    report.Initialize atEvalError, "NETWORK_RUN_FAILURE", issues, checks: Set AT_RunFailure = report
End Function
