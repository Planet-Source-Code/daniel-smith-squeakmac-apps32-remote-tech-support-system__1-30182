Attribute VB_Name = "modWinFunctions"
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long

Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4

Global sConnected As Boolean
Global sListening As Boolean

Public Sub Restart()
ForcedShutdown = ExitWindowsEx(EWX_REBOOT, 0&)
End Sub

Public Sub Shutdown()
StandardShutdown = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Sub

Public Function StayOnTop(frm As Form)
On Error Resume Next

Dim SetWinOnTop As Long
SetWinOnTop = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Function

Public Function NotOnTop(frm As Form)
On Error Resume Next

Dim SetWinOnTop As Long
SetWinOnTop = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Function

Public Sub HideMe()
On Error Resume Next
RegisterServiceProcess GetCurrentProcessId, 1
End Sub

Public Sub ShowMe()
On Error Resume Next
RegisterServiceProcess GetCurrentProcessId, 0
End Sub

Public Sub MouseClick(ByVal X As Long, ByVal Y As Long)
Dim cbuttons As Long
Dim dwExtraInfo As Long
Dim mevent As Long

SetCursorPos X, Y 'set the mouse pos

mevent = MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP 'set event

mouse_event mevent, 0&, 0&, cbuttons, dwExtraInfo 'click button
End Sub
