VERSION 5.00
Begin VB.Form frmSS 
   BorderStyle     =   0  'None
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label load 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label pos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   45
   End
   Begin VB.Label lblImage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image 0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   570
   End
   Begin VB.Image PicBack 
      Height          =   2775
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4095
   End
   Begin VB.Image ViewImage 
      Height          =   2775
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'IMPORTANT!
'For screenshots to work, the remote computer MUST have their C drive
'shared, or have another folder shared.  You can change the path.
'If you can't share the other computer's C drive, then you could
'use a regular transfer method, but it would be slower.


Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function GetTickCount Lib "kernel32" () As Long

Dim LX As Single, LY As Single, LShift As Integer
Dim num As Integer

Private Sub Form_Click()
frmClient.StopMonitor
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then frmClient.StopMonitor: Unload Me
If frmClient.Monitoring = False Then Exit Sub

frmClient.Client.SendData "sk " + CStr(KeyCode) & ";"
End Sub

Private Sub Form_Load()
'Note: See important note at the top
On Error GoTo noshow

ViewImage.Top = 0
ViewImage.Left = 0
ViewImage.Stretch = False
'ViewImage.Picture = LoadPicture("\\" & frmClient.txtHostName.Text & "\c\ss.jpg")
'ViewImage.Picture = LoadPicture("C:\apps32monitor.jpg")
ViewImage.Picture = LoadPicture("C:\ss.jpg")
lblImage.Caption = "Pic: " & MonCount

'StayOnTop Me
Exit Sub
noshow:
'MsgBox "Error showing picture!" & vbCrLf & vbCrLf & "You should be able to access it manually at \\Server\ss.jpg", vbCritical, "Error"
frmClient.AddText "Info> Error Displaying Screenshot: " & Err.Description
End Sub

Public Sub LoadPic()
On Error Resume Next
Show
Form_Load
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

Private Sub Form_Resize()
'PicBack.Height = ScaleHeight
'PicBack.Width = ScaleWidth
ViewImage.Height = ScaleHeight
ViewImage.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
NotOnTop Me
End Sub

Public Sub PicShow(ByVal PixPath As String, fForm As Form)
    On Error GoTo noshow
    Dim dHeight, dIHeight
    Dim dWidth, dIWidth
    Dim dPercent


    With fForm
        .ViewImage.Visible = False
        .ViewImage.Stretch = False
        .Caption = App.Title & " - " & UCase(PixPath)
        .ViewImage.Picture = LoadPicture(PixPath)


        If .ViewImage.Height < .PicBack.Height And .ViewImage.Width < .PicBack.Width Then
            .ViewImage.Visible = True
            Exit Sub
        End If
        dHeight = .ViewImage.Height
        dWidth = .ViewImage.Width
        dIHeight = .PicBack.Height - 1
        dIWidth = .PicBack.Width - 1
        .ViewImage.Stretch = True
        .ViewImage.Height = .PicBack.Height - 2
        dPercent = (.PicBack.Height - 2) / dHeight * 100
        .ViewImage.Width = dWidth / 100 * dPercent


        If .ViewImage.Width > (.PicBack.Width - 2) Then
            .ViewImage.Stretch = False
            dHeight = .ViewImage.Height
            dWidth = .ViewImage.Width
            dIHeight = .PicBack.Height - 1
            dIWidth = .PicBack.Width - 1
            .ViewImage.Stretch = True
            .ViewImage.Width = .PicBack.Width - 1
            dPercent = (.PicBack.Width - 1) / dWidth * 100
            .ViewImage.Height = dHeight / 100 * dPercent
        End If
        .ViewImage.Visible = True
        MidPic fForm
    End With
    Exit Sub
noshow:
End Sub

Public Sub MidPic(ByVal fForm As Form)
    fForm.ViewImage.Move (fForm.PicBack.Width - fForm.ViewImage.Width) / 2, (fForm.ViewImage.Height - fForm.ViewImage.Height) / 2
End Sub

Private Sub ViewImage_Click()
If frmClient.Monitoring Then
    'send one click
    frmClient.SendData "mclick " & Int(LX / 15) & "," & Int(LY / 15) & ";"
Else
    frmClient.StopMonitor
    Unload Me
End If
End Sub

Private Sub ViewImage_DblClick()
If frmClient.Monitoring Then
    'send two clicks
    frmClient.SendData "mclick " & Int(LX / 15) & "," & Int(LY / 15) & "; mclick " & Int(LX / 15) & "," & Int(LY / 15) & "; "
Else
    frmClient.StopMonitor
    Unload Me
End If
End Sub

Private Sub ViewImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    frmClient.AddText "Info> Image Saved"
SavePicture:
    If FileExist("C:\screenshot" & num & ".bmp") Then
        num = num + 1
        GoTo SavePicture
    End If
    SavePicture ViewImage.Picture, "C:\screenshot" & num & ".bmp"
    num = num + 1
End If
End Sub

Private Sub ViewImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'pos.Caption = Int(X / 10) & ", " & Int(Y / 10)
LX = X
LY = Y
LShift = Shift
End Sub
