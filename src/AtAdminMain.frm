Attribute VB_Name = "AtAdminMain"
Option Explicit
Private Sub UserForm_Initialize()
    AtUiStyle Me
    lblHeading.Font.Size = 17: lblRuleTitle.Font.Size = 15: lblPropsTitle.Font.Size = 15
    titleWhen.Font.Bold = True: titleWhat.Font.Bold = True: titleExceptions.Font.Bold = True: titleAction.Font.Bold = True
    lblStatus.Font.Size = 10: lblSaved.Font.Size = 10
End Sub
Private Sub txtTemplate_Change()
    AtUiChanged Me
End Sub
Private Sub txtField_Change()
    AtUiChanged Me
End Sub
Private Sub txtAddress_Change()
    AtUiChanged Me
End Sub
Private Sub txtExpected_Change()
    AtUiChanged Me
End Sub
Private Sub txtMessage_Change()
    AtUiChanged Me
End Sub
Private Sub cboSeverity_Change()
    AtUiChanged Me
End Sub
Private Sub cboCondition_Change()
    AtUiChanged Me
End Sub
Private Sub cboColor_Change()
    AtUiChanged Me
End Sub
Private Sub cboMode_Change()
    AtUiChanged Me
End Sub
Private Sub cboSheet_Change()
    AtUiChanged Me
End Sub
Private Sub chkNote_Click()
    AtUiChanged Me
End Sub
Private Sub cboBook_Change()
    AtUiBookChanged Me
End Sub
Private Sub cmdRefresh_Click()
    AtUiRefreshBooks Me
End Sub
Private Sub cmdTest_Click()
    AtUiCommand Me, "TEST"
End Sub
Private Sub cmdLoad_Click()
    AtUiCommand Me, "LOAD"
End Sub
Private Sub cmdPick_Click()
    AtUiCommand Me, "PICK"
End Sub
Private Sub cmdSave_Click()
    AtUiCommand Me, "SAVE"
End Sub
Private Sub cmdPreview_Click()
    AtUiCommand Me, "PREVIEW"
End Sub
Private Sub cmdPickName_Click()
    AtUiCommand Me, "PICK_NAME"
End Sub
Private Sub cmdPageUp_Click()
    AtUiScroll Me, -1
End Sub
Private Sub cmdPageDown_Click()
    AtUiScroll Me, 1
End Sub
Private Sub cmdClear_Click()
    AtUiCommand Me, "CLEAR"
End Sub
Private Sub cmdClose_Click()
    AtUiHideMain
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: AtUiHideMain
End Sub
Private Sub UserForm_Terminate()
    AtUiStopWatching
End Sub
