Attribute VB_Name = "modATClipboard"
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal window As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal format As Long, ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal bytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal memory As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal target As LongPtr, ByVal source As LongPtr, ByVal bytes As LongPtr)
#Else
Private Declare Function OpenClipboard Lib "user32" (ByVal window As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal format As Long, ByVal memory As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal bytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal memory As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal memory As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal memory As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal target As Long, ByVal source As Long, ByVal bytes As Long)
#End If
Public Sub AT_CopyUnicodeClipboard(ByVal text As String)
#If VBA7 Then
    Dim memory As LongPtr, pointer As LongPtr
#Else
    Dim memory As Long, pointer As Long
#End If
    Dim opened As Boolean, transferred As Boolean, code As String
    On Error GoTo Failed
    If Len(text) = 0 Or Len(text) > 2000000 Then Err.Raise vbObjectError + 3220, , "CLIPBOARD_TEXT_INVALID"
    memory = GlobalAlloc(66, LenB(text) + 2)
    If memory = 0 Then Err.Raise vbObjectError + 3220, , "CLIPBOARD_ALLOC_FAILED"
    pointer = GlobalLock(memory)
    If pointer = 0 Then Err.Raise vbObjectError + 3220, , "CLIPBOARD_ALLOC_FAILED"
    CopyMemory pointer, StrPtr(text), LenB(text)
    GlobalUnlock memory
    If OpenClipboard(Application.Hwnd) = 0 Then Err.Raise vbObjectError + 3220, , "CLIPBOARD_BUSY"
    opened = True
    If EmptyClipboard() = 0 Then Err.Raise vbObjectError + 3220, , "CLIPBOARD_WRITE_FAILED"
    If SetClipboardData(13, memory) = 0 Then Err.Raise vbObjectError + 3220, , "CLIPBOARD_WRITE_FAILED"
    transferred = True: CloseClipboard: Exit Sub
Failed:
    code = AT_StableErrorCode(Err.Description, "CLIPBOARD_WRITE_FAILED")
    If opened Then CloseClipboard
    If memory <> 0 And Not transferred Then GlobalFree memory
    Err.Raise vbObjectError + 3220, "AT_CopyUnicodeClipboard", code
End Sub
