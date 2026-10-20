Attribute VB_Name = "modAtUiLayout"
Option Explicit

Public Function AtUiControlCount(ByVal parent As Object) As Long
    Dim control As Object
    For Each control In parent.Controls
        If control.Parent Is parent Then
            AtUiControlCount = AtUiControlCount + 1
            If TypeName(control) = "Frame" Then AtUiControlCount = AtUiControlCount + AtUiControlCount(control)
        End If
    Next control
End Function

Public Sub AtUiStyle(ByVal parent As Object)
    Dim control As Object
    For Each control In parent.Controls
        control.Font.Name = "Segoe UI": control.Font.Size = 10
        If TypeName(control) = "Frame" Then AtUiStyle control
    Next control
End Sub

' Keep readable fonts and a fixed action area; scroll the content, never zoom it.
Public Sub AtUiLayout(ByVal form As Object, ByVal availableWidth As Double, ByVal availableHeight As Double)
    Dim main As Boolean, names As Variant, name As Variant, control As Object
    Dim x As Double, rows As Long, row As Long, width As Double, height As Double, contentTop As Double, footerTop As Double
    main = form.Name = "AtAdminMain"
    If main Then
        names = Array("cmdLoad", "cmdSave", "cmdTest", "cmdApply", "cmdClear", "cmdClose")
        width = 884: height = 638: contentTop = 46
    Else
        names = Array("cmdCopy", "cmdDetails", "cmdClose")
        width = 710: height = 470: contentTop = 56
        If form.txtDiagnostic.Visible Then height = 600
    End If
    If availableWidth > 0 Then width = WorksheetFunction.Min(width, availableWidth)
    If availableHeight > 0 Then height = WorksheetFunction.Min(height, availableHeight)
    form.Zoom = 100: form.Width = width: form.Height = height
    x = 12: rows = 1
    For Each name In names
        Set control = form.Controls(CStr(name))
        If x + control.Width > form.InsideWidth - 12 And x > 12 Then rows = rows + 1: x = 12
        x = x + control.Width + 8
    Next name
    footerTop = form.InsideHeight - 42 - rows * 38
    x = 12: row = 0
    For Each name In names
        Set control = form.Controls(CStr(name))
        If x + control.Width > form.InsideWidth - 12 And x > 12 Then row = row + 1: x = 12
        control.Move x, footerTop + row * 38, control.Width, 30
        x = x + control.Width + 8
    Next name
    With form.pnlBody
        .Move 12, contentTop, form.InsideWidth - 24, WorksheetFunction.Max(20, footerTop - contentTop - 10)
        .ScrollBars = 3: .KeepScrollBarsVisible = 3
        .ScrollWidth = IIf(main, 872, 700)
        If main Then
            .ScrollHeight = 526
        Else
            .ScrollHeight = IIf(form.txtDiagnostic.Visible, 458, 372)
        End If
    End With
    If main Then
        form.lblHeading.Width = form.InsideWidth - 24
        form.lblStatus.Move 12, form.InsideHeight - 37, form.InsideWidth - 24, 32
    Else
        form.lblResult.Width = form.InsideWidth - 24
        form.lblCopy.Move 12, form.InsideHeight - 37, form.InsideWidth - 24, 32
    End If
End Sub

Public Sub AtUiToggleDiagnostics(ByVal form As Object)
    form.txtDiagnostic.Visible = Not form.txtDiagnostic.Visible
    form.lblDiagnostic.Visible = form.txtDiagnostic.Visible
    form.cmdDetails.Caption = IIf(form.txtDiagnostic.Visible, "Скрыть подробности", "Технические подробности")
    form.pnlBody.ScrollTop = 0
    AtUiFit form
End Sub

Public Function AtUiOutcomeColor(ByVal status As String) As Long
    Select Case status
        Case "PASS": AtUiOutcomeColor = RGB(27, 119, 65)
        Case "FAIL": AtUiOutcomeColor = RGB(180, 40, 40)
        Case Else: AtUiOutcomeColor = RGB(151, 98, 0)
    End Select
End Function
