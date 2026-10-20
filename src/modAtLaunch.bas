Attribute VB_Name = "modAtLaunch"
Option Explicit

Public Function AtNewJournalStore(Optional ByVal directory As String = "") As CAtJournalStore
    Dim store As New CAtJournalStore
    store.Initialize directory: Set AtNewJournalStore = store
End Function

' Shared entry for the demo and a future installed XLAM Ribbon callback.
Public Sub AtRibbonAdmin(ByVal control As Object)
    AtAdminOpen
End Sub

Public Function AtBindDemoLauncher() As Boolean
    On Error GoTo NotDemo
    Dim button As Object, expected As String
    Set button = ThisWorkbook.Worksheets("AccessTools").Buttons("AtOpenConstructor")
    expected = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AtAdminOpen"
    If button.OnAction <> expected Then button.OnAction = expected
    AtBindDemoLauncher = button.OnAction = expected
NotDemo:
End Function
