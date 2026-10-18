Attribute VB_Name = "modATTemplateMenu"
Option Explicit
Private mPicker As Object
Public Function AT_AvailableTemplates(Optional ByVal query As String = "") As Collection
    Dim items As New Collection, title As String
    ' Only adapters with implemented checks are offered. Add future registered adapters here.
    title = "Сетевая связанность"
    query = Trim$(query)
    If Len(query) = 0 Or InStr(1, title & " NETWORK_ACCESS", query, vbTextCompare) > 0 Then items.Add Array("NETWORK_ACCESS", title)
    Set AT_AvailableTemplates = items
End Function
Public Sub ШаблоныИПравила()
    If Not AT_PackageAllowCommand() Then Exit Sub
    On Error GoTo Failed
    If Not mPicker Is Nothing Then Unload mPicker
    Set mPicker = VBA.UserForms.Add("frmATTemplates")
    mPicker.Show 0
    Exit Sub
Failed:
    MsgBox "Не удалось открыть список шаблонов. " & Err.Description, vbExclamation, "AccessTools 2.0"
End Sub
Public Sub AT_RibbonTemplates(ByVal control As Object)
    ШаблоныИПравила
End Sub
Public Sub AT_RibbonDiagnostic(ByVal control As Object)
    If AT_PackageAllowCommand() Then AT_ResultReport
End Sub
Public Sub AT_RibbonClear(ByVal control As Object)
    If AT_PackageAllowCommand() Then AT_ResultClear
End Sub
Public Sub AT_RibbonRules(ByVal control As Object)
    AT_OpenNetworkRules
End Sub
