Attribute VB_Name = "AtAdminResults"
Option Explicit
Private Sub UserForm_Initialize()
    AtUiStyle Me
    lblResult.Font.Size = 20
End Sub
Private Sub cmdDetails_Click()
    AtUiToggleDiagnostics Me
End Sub
Private Sub cmdCopy_Click()
    AtUiCopyRemark
End Sub
Private Sub cmdPageUp_Click()
    AtUiScroll Me, -1
End Sub
Private Sub cmdPageDown_Click()
    AtUiScroll Me, 1
End Sub
Private Sub cmdClose_Click()
    Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: Me.Hide
End Sub
