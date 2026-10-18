Attribute VB_Name = "modATFamiliarWorkflowTests"
Option Explicit
Public Sub AT_TestFamiliarWorkflow()
    Dim writer As New ATCoordinatorWorkbookWriter, book As Workbook, raw As Workbook, view As Workbook
    Dim registry As New ATNetworkRegistryResult, entries As New Collection, issue As String
    On Error GoTo Failed
    ' The copy layer accepts an already calculated coordinator book, never reruns checks.
    Set raw = Application.Workbooks.Add(xlWBATWorksheet)
    raw.Worksheets(1).Name = "Исходная выгрузка": raw.Worksheets(1).Range("J2").Value2 = "Extra source column"
    raw.Worksheets(1).Range("J2").Interior.Color = RGB(12, 34, 56)
    Set book = Application.Workbooks.Add(xlWBATWorksheet): book.Worksheets(1).Name = "Сводка координатора"
    book.Worksheets(1).Range("E6").Value2 = "Calculated service"
    book.Worksheets.Add(After:=book.Worksheets(1)).Name = "Superset"
    Set view = CallByName(writer, "CreateWorkingCopy", VbMethod, raw, book)
    AT_Assert Not view Is Nothing, "working Superset copy is created from existing output without recomputation"
    If Not view Is Nothing Then
        AT_Assert view.Worksheets.Count = 3 And view.Worksheets(1).Range("J2").Value2 = "Extra source column", "working copy preserves all source columns"
        AT_Assert view.Worksheets(1).Range("J2").Interior.Color = RGB(12, 34, 56), "working copy retains source formatting"
        AT_Assert view.Worksheets("Сводка координатора").Range("E6").Value2 = "Calculated service", "coordinator output lives in working Superset"
        AT_Assert raw.Worksheets.Count = 1 And book.Worksheets.Count = 2, "copy layer does not change either input book"
        view.Close False: Set view = Nothing
    End If
    raw.Worksheets(1).Name = "Superset"
    Set view = CallByName(writer, "CreateWorkingCopy", VbMethod, raw, book)
    AT_Assert view.Worksheets("Superset").Index = 3 And view.Worksheets(1).Range("J2").Value2 = "Extra source column", "reserved result names rename only copied source sheets"
    AT_Assert raw.Worksheets(1).Name = "Superset", "source sheet name survives result-name collision"
Cleanup:
    On Error Resume Next
    If Not view Is Nothing Then view.Close False
    If Not raw Is Nothing Then raw.Close False
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    Exit Sub
Failed:
    issue = Err.Description: AT_Assert False, "working copy unavailable: " & issue: Resume Cleanup
End Sub
Public Sub AT_TestFamiliarForms()
    Dim window As Object, picker As Object, message As String
    On Error GoTo Failed
    Set window = VBA.UserForms.Add("frmATResults")
    AT_Assert window.Controls("btnCopy").Caption = "Копировать" And window.Controls("btnGo").Caption = "Перейти", "legacy result window provides copy and cell navigation"
    AT_Assert window.Controls("btnRepeat").Caption = "Повторить проверку" And window.Controls("btnClear").Caption = "Очистить отметки", "legacy result window provides repeat and owned-mark cleanup"
    Set picker = VBA.UserForms.Add("frmATTemplates")
    AT_Assert picker.Controls("txtSearch").Value = "" And picker.Controls("lstTemplates").ListCount = 1, "template picker starts with supported templates and a search field"
    picker.Controls("txtSearch").Value = "несуществующий шаблон"
    AT_Assert picker.Controls("lstTemplates").ListCount = 0 And Not picker.Controls("btnCheck").Enabled, "template search prevents actions on an empty selection"
    picker.Controls("txtSearch").Value = "связан"
    AT_Assert picker.Controls("lstTemplates").ListCount = 1 And picker.Controls("btnCheck").Enabled, "template search matches Russian display names"
Cleanup:
    On Error Resume Next
    If Not window Is Nothing Then Unload window
    If Not picker Is Nothing Then Unload picker
    On Error GoTo 0
    Exit Sub
Failed:
    message = Err.Description: AT_Assert False, "familiar UI unavailable: " & message: Resume Cleanup
End Sub
