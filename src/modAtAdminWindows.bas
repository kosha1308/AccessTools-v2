Attribute VB_Name = "modAtAdminWindows"
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal className As LongPtr, ByVal caption As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal window As LongPtr, ByRef process As Long) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal window As LongPtr, ByVal after As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
#End If
Private events As CAtAdminWindowEvents, wanted As Boolean, picking As Boolean, raising As Boolean, serial As Long

Public Sub AtUiWatchWindows()
    If events Is Nothing Then
        Set events = New CAtAdminWindowEvents
        events.Attach Application
    End If
    wanted = True
End Sub
Public Sub AtUiStopWatching()
    wanted = False: picking = False: Set events = Nothing
End Sub
Public Property Get AtUiWatchingWindows() As Boolean
    AtUiWatchingWindows = Not events Is Nothing
End Property
Public Sub AtUiPicking(ByVal value As Boolean)
    picking = value
End Sub
Public Sub AtUiHideMain()
    AtUiStopWatching
    AtAdminMain.Hide
End Sub
Public Sub AtUiKeepVisible()
    If Not wanted Or picking Or raising Then Exit Sub
    Dim form As Object
    For Each form In VBA.UserForms
        If TypeName(form) = "AtAdminMain" And form.Visible Then
            If Not AtUiRaiseWithinExcel(form) Then form.lblStatus.Caption = "Не удалось показать окно поверх Excel. Повторно нажмите «Открыть конструктор»."
        End If
    Next form
End Sub
Public Function AtUiRaiseWithinExcel(ByVal form As Object) As Boolean
    Dim original As String, token As String, className As String, process As Long
#If VBA7 Then
    Dim window As LongPtr
    On Error GoTo Restore
    If Not form.Visible Or raising Then Exit Function
    raising = True: serial = serial + 1
    original = form.Caption
    token = "AT_UI_" & Hex$(GetCurrentProcessId()) & "_" & serial
    form.Caption = token: className = "ThunderDFrame"
    window = FindWindowW(StrPtr(className), StrPtr(token))
    If window = 0 Then className = "ThunderXFrame": window = FindWindowW(StrPtr(className), StrPtr(token))
    form.Caption = original
    If window = 0 Then GoTo Restore
    GetWindowThreadProcessId window, process
    If process <> GetCurrentProcessId() Then GoTo Restore
    ' HWND_TOP, NOT HWND_TOPMOST. Preserve focus, size, position AND owner z-order.
    AtUiRaiseWithinExcel = SetWindowPos(window, 0, 0, 0, 0, 0, &H213) <> 0
Restore:
    If original <> "" Then form.Caption = original
    raising = False
#End If
End Function
