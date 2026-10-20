Private Sub Workbook_Open()
    AtBindDemoLauncher
End Sub
Private Sub Workbook_Activate()
    AtBindDemoLauncher
End Sub
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If Success Then AtBindDemoLauncher
End Sub
