Attribute VB_Name = "AtAdminResults"
Option Explicit
Private Sub UserForm_Initialize()
    Dim control As Object
    For Each control In Me.Controls
        control.Font.Name = "Segoe UI": control.Font.Size = 10
    Next control
    lblResult.Font.Size = 20
End Sub
Private Sub cmdCopy_Click()
    AtUiCopyRemark
End Sub
Private Sub cmdClose_Click()
    Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: Me.Hide
End Sub
