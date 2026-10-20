VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CAtAdminWindowEvents"
Option Explicit
Private WithEvents app As Excel.Application
Friend Sub Attach(ByVal instance As Excel.Application)
    Set app = instance
End Sub
Private Sub app_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    AtUiKeepVisible
End Sub
Private Sub app_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    AtUiKeepVisible
End Sub
