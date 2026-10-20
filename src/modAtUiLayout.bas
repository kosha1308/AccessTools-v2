Attribute VB_Name = "modAtUiLayout"
Option Explicit
Private arranging As Boolean

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

Public Sub AtUiRelayout(ByVal form As Object)
    If arranging Then Exit Sub
    AtUiLayout form, form.Width, form.Height
End Sub

' Native layout only: no hook, subclass, timer, topmost flag or window callback.
Public Sub AtUiLayout(ByVal form As Object, ByVal availableWidth As Double, ByVal availableHeight As Double)
    If arranging Then Exit Sub
    On Error GoTo Broken
    arranging = True
    Dim main As Boolean, names As Variant, name As Variant, control As Object
    Dim x As Double, rows As Long, row As Long, width As Double, height As Double, top As Double, footer As Double, extent As Double
    main = form.Name = "AtAdminMain"
    If main Then
        names = Array("cmdTest", "cmdPreview", "cmdLoad", "cmdSave", "cmdClear", "cmdClose")
        width = 1020: height = 720: top = 46
        If form.cmdSelectedPreview.Visible Then
            names = Array("cmdTest", "cmdSelectedPreview", "cmdPreview", "cmdLoad", "cmdSave", "cmdClear", "cmdClose")
            width = 900: height = 650
        End If
    Else
        names = Array("cmdCopy", "cmdDetails", "cmdClose")
        width = 710: height = 530: top = 56
        If form.txtDiagnostic.Visible Then height = 620
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
    footer = form.InsideHeight - 42 - rows * 38
    x = 12: row = 0
    For Each name In names
        Set control = form.Controls(CStr(name))
        If x + control.Width > form.InsideWidth - 12 And x > 12 Then row = row + 1: x = 12
        control.Move x, footer + row * 38, control.Width, 30
        x = x + control.Width + 8
    Next name
    With form.pnlBody
        .ScrollBars = 0: .KeepScrollBarsVisible = 0
        .Move 12, top, form.InsideWidth - 24, WorksheetFunction.Max(20, footer - top - 10)
        ' Reserve a small vertical gutter so showing it never causes x overflow.
        If main Then
            MainContent form, .InsideWidth - 18
        Else
            ResultContent form, .InsideWidth - 18
        End If
        extent = ContentBottom(form.pnlBody) + 8
        .ScrollHeight = WorksheetFunction.Max(extent, .InsideHeight)
        .ScrollWidth = .InsideWidth: .ScrollLeft = 0
        If extent > .InsideHeight + 1 Then
            .ScrollBars = 2
            .ScrollTop = WorksheetFunction.Min(.ScrollTop, WorksheetFunction.Max(0, .ScrollHeight - .InsideHeight))
        Else
            .ScrollTop = 0
        End If
    End With
    form.cmdPageUp.Visible = form.pnlBody.ScrollBars = 2
    form.cmdPageDown.Visible = form.cmdPageUp.Visible
    form.cmdPageUp.Move form.InsideWidth - 84, 10, 32, 25
    form.cmdPageDown.Move form.InsideWidth - 44, 10, 32, 25
    If main Then
        form.lblHeading.Width = form.InsideWidth - 112
        form.lblStatus.Move 12, form.InsideHeight - 37, form.InsideWidth - 24, 32
    Else
        form.lblResult.Width = form.InsideWidth - 112
        form.lblCopy.Move 12, form.InsideHeight - 37, form.InsideWidth - 24, 32
    End If
    arranging = False
    Exit Sub
Broken:
    arranging = False
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub Place(ByVal form As Object, ByVal name As String, ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double)
    form.Controls(name).Move x, y, width, height
End Sub
Private Function TextHeight(ByVal label As Object, ByVal x As Double, ByVal y As Double, ByVal width As Double) As Double
    label.AutoSize = False: label.WordWrap = True
    label.Move x, y, width, 16
    label.AutoSize = True
    TextHeight = WorksheetFunction.Max(16, label.Height)
End Function
Private Function ContentBottom(ByVal frame As Object) As Double
    Dim control As Object
    For Each control In frame.Controls
        If control.Visible Then ContentBottom = WorksheetFunction.Max(ContentBottom, control.Top + control.Height)
    Next control
End Function
Private Function Card(ByVal form As Object, ByVal suffix As String, ByVal x As Double, ByVal y As Double, ByVal width As Double) As Double
    Dim titleHeight As Double, bodyHeight As Double
    titleHeight = TextHeight(form.Controls("title" & suffix), x + 12, y + 9, width - 24)
    bodyHeight = TextHeight(form.Controls("lbl" & suffix), x + 12, y + titleHeight + 16, width - 24)
    Card = titleHeight + bodyHeight + 28
    Place form, "card" & suffix, x, y, width, Card
End Function
Private Function RuleCards(ByVal form As Object, ByVal x As Double, ByVal y As Double, ByVal width As Double) As Double
    Dim suffix As Variant, h As Double
    Place form, "lblRuleTitle", x, y, width, 24: y = y + 32
    For Each suffix In Array("When", "What", "Exceptions", "Action")
        h = Card(form, CStr(suffix), x, y, width): y = y + h + 10
    Next suffix
    If form.lblDetails.Visible Then y = y + TextHeight(form.lblDetails, x, y, width) + 8
    RuleCards = y
End Function
Private Function PropertyRow(ByVal form As Object, ByVal label As String, ByVal inputName As String, ByVal x As Double, ByVal y As Double, ByVal width As Double, Optional ByVal height As Double = 24) As Double
    Dim h As Double
    h = TextHeight(form.Controls(label), x, y, width)
    Place form, inputName, x, y + h + 2, width, height
    PropertyRow = y + h + height + 8
End Function
Private Function Properties(ByVal form As Object, ByVal x As Double, ByVal y As Double, ByVal width As Double) As Double
    Dim h As Double, leftWidth As Double, rightX As Double
    Place form, "lblPropsTitle", x, y, width, 24: y = y + 32
    If form.txtRuleName.Visible Then
        h = TextHeight(form.lblRuleName, x, y, width): y = y + h + 2
        Place form, "txtRuleName", x, y, width - 150, 24
        Place form, "chkRuleEnabled", x + width - 144, y, 144, 24: y = y + 32
        y = PropertyRow(form, "lblRuleField", "cboRuleField", x, y, width)
    End If
    h = TextHeight(form.lblBook, x, y, width): y = y + h + 2
    Place form, "cboBook", x, y, width - 42, 24: Place form, "cmdRefresh", x + width - 36, y, 36, 24: y = y + 32
    leftWidth = (width - 12) * 0.48: rightX = x + leftWidth + 12
    h = TextHeight(form.lblSheet, x, y, leftWidth)
    h = WorksheetFunction.Max(h, TextHeight(form.lblAddress, rightX, y, width - leftWidth - 12)): y = y + h + 2
    Place form, "cboSheet", x, y, leftWidth, 24
    Place form, "txtAddress", rightX, y, width - leftWidth - 90, 24
    Place form, "cmdPick", x + width - 72, y, 72, 24: y = y + 32
    h = TextHeight(form.lblField, x, y, width): y = y + h + 2
    If width >= 300 Then
        Place form, "txtField", x, y, width - 150, 24
        Place form, "cmdPickName", x + width - 144, y, 144, 24: y = y + 32
    Else
        Place form, "txtField", x, y, width, 24: y = y + 28
        Place form, "cmdPickName", x, y, width, 26: y = y + 34
    End If
    y = PropertyRow(form, "lblCondition", "cboCondition", x, y, width)
    y = PropertyRow(form, "lblExpected", "txtExpected", x, y, width)
    y = PropertyRow(form, "lblMessage", "txtMessage", x, y, width, 48)
    h = TextHeight(form.lblSeverity, x, y, leftWidth)
    h = WorksheetFunction.Max(h, TextHeight(form.lblColor, rightX, y, width - leftWidth - 12)): y = y + h + 2
    Place form, "cboSeverity", x, y, leftWidth, 24
    Place form, "cboColor", rightX, y, width - leftWidth - 36, 24
    Place form, "lblColorSample", x + width - 18, y, 18, 24: y = y + 32
    Place form, "chkNote", x, y, width, 28
    Properties = y + 36
End Function
Private Sub MainContent(ByVal form As Object, ByVal width As Double)
    Dim cardX As Double, cardWidth As Double, propX As Double, propWidth As Double, bottom As Double, y As Double
    Dim navHeight As Double
    If width >= 820 Then
        cardX = 156: cardWidth = (width - 168) * 0.48: propX = cardX + cardWidth + 12: propWidth = width - propX
        Place form, "txtTemplate", cardX, 4, cardWidth, 24
        Place form, "cboMode", propX, 4, propWidth, 24
        y = 42
        navHeight = TextHeight(form.lblNavTitle, 0, y, 144)
        navHeight = y + navHeight + 6
        If form.cboFieldFilter.Visible Then
            Place form, "cboFieldFilter", 0, navHeight, 144, 24: navHeight = navHeight + 30
            Place form, "lstNavigation", 0, navHeight, 144, 120: navHeight = navHeight + 128
            navHeight = NavigationActions(form, 0, navHeight, 144, False)
        Else
            Place form, "lstNavigation", 0, navHeight, 144, 180: navHeight = navHeight + 194
        End If
        form.lblNavHelp.Visible = True
        navHeight = navHeight + TextHeight(form.lblNavHelp, 0, navHeight, 144)
        bottom = WorksheetFunction.Max(RuleCards(form, cardX, y, cardWidth), Properties(form, propX, y, propWidth), navHeight)
    Else
        Place form, "txtTemplate", 0, 4, width * 0.6 - 6, 24
        Place form, "cboMode", width * 0.6 + 6, 4, width * 0.4 - 6, 24
        navHeight = TextHeight(form.lblNavTitle, 0, 36, width)
        y = 36 + navHeight + 4
        If form.cboFieldFilter.Visible Then Place form, "cboFieldFilter", 0, y, width, 24: y = y + 30
        Place form, "lstNavigation", 0, y, width, 40: y = y + 48
        If form.cmdAddField.Visible Then y = NavigationActions(form, 0, y, width, True)
        form.lblNavHelp.Visible = False
        bottom = Properties(form, 0, y, width)
        bottom = RuleCards(form, 0, bottom + 12, width)
    End If
    bottom = bottom + TextHeight(form.lblSaved, 0, bottom + 4, width) + 12
End Sub
Private Function NavigationActions(ByVal form As Object, ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal twoColumns As Boolean) As Double
    Dim names As Variant, i As Long, w As Double
    names = Array("cmdAddField", "cmdAddRule", "cmdCopyRule", "cmdDeleteRule")
    w = width
    If twoColumns Then w = (width - 8) / 2
    For i = 0 To 3
        If twoColumns Then
            Place form, CStr(names(i)), x + (i Mod 2) * (w + 8), y + (i \ 2) * 32, w, 26
        Else
            Place form, CStr(names(i)), x, y + i * 32, w, 26
        End If
    Next i
    NavigationActions = y + IIf(twoColumns, 72, 136)
End Function
Private Sub ResultContent(ByVal form As Object, ByVal width As Double)
    Dim y As Double
    If form.lstResults.Visible Then Place form, "lstResults", 0, 0, width, 90: y = 100
    y = PropertyRow(form, "lblTarget", "txtTarget", 0, y, width)
    y = PropertyRow(form, "lblActual", "txtActual", 0, y, width, 45)
    y = PropertyRow(form, "lblRemark", "txtRemark", 0, y, width, 60)
    y = PropertyRow(form, "lblActions", "txtActions", 0, y, width, 50)
    y = PropertyRow(form, "lblOutput", "txtOutput", 0, y, width, 45)
    If form.txtDiagnostic.Visible Then y = PropertyRow(form, "lblDiagnostic", "txtDiagnostic", 0, y, width, 80)
End Sub
Public Sub AtUiScroll(ByVal form As Object, ByVal direction As Long)
    With form.pnlBody
        .ScrollTop = WorksheetFunction.Max(0, WorksheetFunction.Min(.ScrollHeight - .InsideHeight, .ScrollTop + direction * WorksheetFunction.Max(40, .InsideHeight * 0.75)))
    End With
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
