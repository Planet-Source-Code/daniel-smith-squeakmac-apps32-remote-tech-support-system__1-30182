VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClient 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apps32 Remote"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      Height          =   405
      Left            =   120
      TabIndex        =   31
      Top             =   4080
      Width           =   5175
      Begin VB.Timer tmrPing 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   4080
         Top             =   120
      End
      Begin VB.Image imgPing 
         Height          =   255
         Left            =   4560
         ToolTipText     =   "Ping Status: "
         Top             =   120
         Width           =   615
      End
      Begin VB.Shape State 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   4680
         Shape           =   3  'Circle
         Top             =   180
         Width           =   195
      End
      Begin VB.Shape State 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   4920
         Shape           =   3  'Circle
         Top             =   180
         Width           =   135
      End
      Begin VB.Label lblRemote 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "0.0.0.0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1590
         TabIndex        =   35
         Top             =   150
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Remote Connection:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   90
         TabIndex        =   34
         Top             =   150
         Width           =   1470
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "On Port:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2910
         TabIndex        =   33
         Top             =   150
         Width           =   615
      End
      Begin VB.Label lblPort 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3570
         TabIndex        =   32
         Top             =   150
         Width           =   90
      End
   End
   Begin TabDlg.SSTab TabStrip 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   8388608
      TabCaption(0)   =   "Apps32"
      TabPicture(0)   =   "frmClient.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tmrApps"
      Tab(0).Control(1)=   "lstApps"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Chat"
      TabPicture(1)   =   "frmClient.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdChat"
      Tab(1).Control(1)=   "txtChat"
      Tab(1).Control(2)=   "txtSend"
      Tab(1).Control(3)=   "cmdSend"
      Tab(1).Control(4)=   "Image2"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Debug"
      TabPicture(2)   =   "frmClient.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtDebug"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtCommand"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdDebug"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Commands"
      TabPicture(3)   =   "frmClient.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label7"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "fCommands"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "fraDownload"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Connect"
      TabPicture(4)   =   "frmClient.frx":04B2
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label4"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label3"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label10"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Client"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "tmrTimeOut"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "fConnection"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "fConInfo"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).ControlCount=   8
      Begin VB.Frame fraDownload 
         Caption         =   "File Download Progress"
         Height          =   855
         Left            =   -74880
         TabIndex        =   36
         Top             =   2880
         Visible         =   0   'False
         Width           =   4935
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label lblProg 
            AutoSize        =   -1  'True
            Caption         =   "Progress: "
            Height          =   195
            Left            =   3600
            TabIndex        =   40
            Top             =   540
            Width           =   705
         End
         Begin VB.Label lblRec 
            AutoSize        =   -1  'True
            Caption         =   "Received:"
            Height          =   195
            Left            =   1800
            TabIndex        =   39
            Top             =   540
            Width           =   735
         End
         Begin VB.Label lblSize 
            AutoSize        =   -1  'True
            Caption         =   "File Size: "
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   540
            Width           =   675
         End
      End
      Begin VB.Frame fConInfo 
         Height          =   1575
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2175
         Begin VB.TextBox txtPort 
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Text            =   "505"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtHost 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Text            =   "192.168.0.4"
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remote Port: "
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remote Host: "
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1020
         End
      End
      Begin VB.Frame fConnection 
         Height          =   1215
         Left            =   3000
         TabIndex        =   10
         Top             =   480
         Width           =   1815
         Begin VB.CommandButton cmdDisconnect 
            Caption         =   "Disconnect"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdConnect 
            Caption         =   "Connect"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Timer tmrTimeOut 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   600
         Top             =   3240
      End
      Begin VB.Frame fCommands 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   9
         Top             =   360
         Width           =   4935
         Begin VB.CommandButton cmdTerminate 
            Caption         =   "Terminate App"
            Height          =   375
            Left            =   3240
            TabIndex        =   28
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CommandButton cmdSpawn 
            Caption         =   "Spawn App"
            Height          =   375
            Left            =   1680
            TabIndex        =   27
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdShow 
            Caption         =   "Show"
            Height          =   375
            Left            =   3240
            TabIndex        =   26
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmdHide 
            Caption         =   "Hide"
            Height          =   375
            Left            =   3240
            TabIndex        =   25
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdShutdown 
            Caption         =   "Shutdown"
            Height          =   375
            Left            =   1680
            TabIndex        =   24
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdRestart 
            Caption         =   "Restart"
            Height          =   375
            Left            =   1680
            TabIndex        =   23
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton cmdEndServer 
            Caption         =   "End Server"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdControl 
            Caption         =   "Control"
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton cmdScreenhot 
            Caption         =   "Screenshot"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdDebug 
         Caption         =   "Send"
         Default         =   -1  'True
         Height          =   285
         Left            =   -70920
         TabIndex        =   8
         Top             =   3450
         Width           =   975
      End
      Begin VB.TextBox txtCommand 
         Height          =   285
         Left            =   -74880
         TabIndex        =   7
         Top             =   3450
         Width           =   3855
      End
      Begin VB.CommandButton cmdChat 
         Caption         =   "Chat"
         Height          =   285
         Left            =   -70920
         TabIndex        =   6
         Top             =   3420
         Width           =   975
      End
      Begin VB.Timer tmrApps 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   -75000
         Top             =   3480
      End
      Begin VB.TextBox txtChat 
         Height          =   2775
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   540
         Width           =   4935
      End
      Begin VB.TextBox txtSend 
         Height          =   285
         Left            =   -74880
         TabIndex        =   4
         Top             =   3420
         Width           =   3855
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   285
         Left            =   -70920
         TabIndex        =   3
         Top             =   3420
         Width           =   975
      End
      Begin VB.TextBox txtDebug 
         Height          =   2895
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   480
         Width           =   4935
      End
      Begin VB.ListBox lstApps 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   4935
      End
      Begin MSWinsockLib.Winsock Client 
         Left            =   120
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemoteHost      =   "192.168.0.1"
         RemotePort      =   505
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://pagemac.cjb.net"
         BeginProperty Font 
            Name            =   "Microstyle Bold Extended ATT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1230
         TabIndex        =   42
         Top             =   3240
         Width           =   2565
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PageMac Programming"
         BeginProperty Font 
            Name            =   "Microstyle Bold Extended ATT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1230
         TabIndex        =   41
         Top             =   3120
         Width           =   2565
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Access"
         BeginProperty Font 
            Name            =   "Microstyle Bold Extended ATT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   -73200
         TabIndex        =   30
         Top             =   2640
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apps32"
         BeginProperty Font 
            Name            =   "Microstyle Bold Extended ATT"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   -73200
         TabIndex        =   29
         Top             =   2280
         Width           =   1605
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Access"
         BeginProperty Font 
            Name            =   "Microstyle Bold Extended ATT"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   1680
         TabIndex        =   19
         Top             =   2640
         Width           =   1605
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apps32"
         BeginProperty Font 
            Name            =   "Microstyle Bold Extended ATT"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1680
         TabIndex        =   18
         Top             =   2280
         Width           =   1605
      End
      Begin VB.Image Image2 
         Height          =   135
         Left            =   -74880
         Top             =   360
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   135
         Left            =   -74880
         Top             =   360
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Chat As Boolean
Public ChatUser As String, Ping As Boolean, BadPings As Integer
Public Monitoring As Boolean, StopLast As Boolean
Dim waiting As Boolean
Public bSendFile As Boolean

Private Sub Client_Close()
AddText "*** Disconnected ***"
tmrApps.Enabled = False
tmrPing.Enabled = False
lstApps.Clear
fConInfo.Enabled = True
LockButtons
cmdConnect.Enabled = True
cmdDisconnect.Enabled = False
imgPing.ToolTipText = "Ping Status: Disconnected"
End Sub

Private Sub Client_Connect()
AddText "*** Connected ***"
tmrApps.Enabled = True
tmrPing.Enabled = True
fConInfo.Enabled = False
UnlockButtons
BadPings = 0
txtDebug.Text = txtDebug.Text & vbCrLf & "<Apps32 Client Online>" & vbCrLf & vbCrLf
lblRemote.Caption = Client.RemoteHostIP
lblPort.Caption = Client.RemotePort
cmdConnect.Enabled = False
cmdDisconnect.Enabled = True
imgPing.ToolTipText = "Ping Status: Idle  Bad Pings: " & BadPings
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String
Dim AppsRunning() As String
Dim DataParsed As Boolean

If StopLast Then StopLast = False: Exit Sub
Client.GetData Data

If bSendFile Then GoTo FileTransfer

If InStr(1, Data, "success") <> 0 Then
    UnlockButtons
    AddText "Server> Success"
ElseIf LCase(Mid(Data, 1, 5)) = "apps:" Then
    AppsRunning = Split(Data, vbCrLf)
    lstApps.Clear
    For i = 1 To UBound(AppsRunning)
        lstApps.AddItem AppsRunning(i)
    Next i
ElseIf InStr(1, Data, "ssdone") <> 0 Then
    UnlockButtons
    If waiting = False Then
        AddText "Info> Unexpected screenshot arrival."
        Exit Sub
    End If
    If Monitoring = False Then
        AddText "Info> Screenshot taken"
        frmSS.Show
        tmrApps.Enabled = True
        tmrPing.Enabled = True
        waiting = False
    Else
        MonCount = MonCount + 1
        Client.SendData "ss 1"
        frmSS.LoadPic
    End If
ElseIf LCase(Mid(Data, 1, 7)) = "server>" Then
    AddChat Data
ElseIf Mid(LCase(Data), 1, 9) = "connected" Then
    If Data <> "Connected.  Authenticate." Then
        AddText "Info> Invalid Apps32 Server"
    Else
        AddText "Info> Connected"
        Client.SendData "apps"
    End If
ElseIf Mid(Data, 1, 4) = "sks " Then
    AddText "Info> Key Sent: " & Mid(Data, 5)
ElseIf InStr(1, Data, "Unreconized") <> 0 Then
    Beep
    AddText "Server> " & Data
ElseIf Data = "pong" Then
    State(1).FillColor = vbGreen
    Ping = False
    Pause 0.1
    State(0).FillColor = &H808080
    State(1).FillColor = &H808080
    imgPing.ToolTipText = "Last Ping: Success  Bad Pings: " & BadPings
ElseIf Mid(Data, 1, 8) = "OpenFile" Then
    FN = Mid(Data, 10, InStr(10, Data, "|") - 10)
    FS = Val(Mid(Data, InStr(10, Data, "|") + 1))
    
    Open FN For Binary Access Write As #1
    bSendFile = True
    AddText "Info> Downloading file..."
    
    fraDownload.Visible = True
    lblSize.Caption = "File Size: " & FS
    ProgressBar1.Max = FS
Else
    AddText "Server> " & Data
End If

Exit Sub

FileTransfer:
ii = InStr(1, Data, "ClosFile")
If ii <> 0 Then
    If ii <> 1 Then
        Put #1, , Left(Data, ii)
    End If
    Close #1
    bSendFile = False
    AddText "Info> Download complete"
    fraDownload.Visible = False
Else
    Put #1, , Data
    lblProg.Caption = "Progress: " & Int(100 / ProgressBar1.Max * ProgressBar1.Value) & " %"
    ProgressBar1.Value = ProgressBar1.Value + Len(Data)
    lblRec.Caption = "Received: " & ProgressBar1.Value
End If

DoEvents
End Sub

Private Sub Client_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Select Case Number
    Case 10061
        AddText "*** Error: Unable to connect to server application ***"
    Case 10053
        AddText "*** Error: Server application terminated or timeout ***"
    Case Else
        AddText "*** Error " & Number & ": " & Description & " ***"
End Select
tmrApps.Enabled = False
lstApps.Clear
LockButtons
cmdConnect.Enabled = True
End Sub

Private Sub cmdChat_Click()
Dim User As String
User = InputBox("Enter your chat handle: ", "Chat Handle")
If User = "" Then Exit Sub
ChatUser = User
Chat = True
Client.SendData "chat " & ChatUser
txtChat.Text = "Chat Connection Opened" & vbCrLf & vbCrLf
cmdChat.Visible = False
AddText "Info> Now chatting as " & ChatUser & "."
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdConnect_Click()
On Error Resume Next
Client.Close
Client.RemoteHost = txtHost.Text
Client.RemotePort = txtPort.Text
Client.Connect
TabStrip.Tab = 2
txtDebug.Text = ""
AddText "*** Connecting... ***"
End Sub

Private Sub cmdControl_Click()
    Client.SendData "ss 1"
    waiting = True
    AddText "Info> Monitoring server..."
    Monitoring = True
    tmrApps.Enabled = False
    tmrPing.Enabled = False
End Sub

Private Sub cmdDebug_Click()
If txtCommand.Text = "" Then Exit Sub
If LCase(txtCommand.Text) = "pad" Then
    SendData "spawn c:\windows\notepad.exe"
ElseIf LCase(txtCommand.Text) = "sol" Then
    SendData "spawn c:\windows\sol.exe"
ElseIf LCase(txtCommand.Text) = "monitor" Then
    Client.SendData "ss 1"
    waiting = True
    AddText "Info> Monitoring server..."
    Monitoring = True
    tmrApps.Enabled = False
    tmrPing.Enabled = False
ElseIf Left(LCase(txtCommand.Text), 2) = "ss" Then
    tmrApps.Enabled = False
    tmrPing.Enabled = False
    waiting = True
    SendData txtCommand.Text
Else
    SendData txtCommand.Text
End If
txtCommand.Text = ""
End Sub

Private Sub cmdDisconnect_Click()
Client.Close
Client_Close
End Sub

Private Sub cmdEndServer_Click()
r = MsgBox("Are you sure you want to end the Server?", vbQuestion + vbYesNo, "End Server")
If r = vbYes Then SendData "endserver"
End Sub

Private Sub cmdHide_Click()
SendData "hide"
End Sub

Private Sub cmdRestart_Click()
SendData "restart"
End Sub

Private Sub cmdShutdown_Click()
SendData "shutdown"
End Sub

Private Sub cmdScreenhot_Click()
tmrApps.Enabled = False
tmrPing.Enabled = False
waiting = True
SendData "ss 1"
End Sub

Private Sub cmdSend_Click()
If Chat = False Then Exit Sub
SendMessage txtSend.Text
txtSend.Text = ""
End Sub

Private Sub cmdSpawn_Click()
Dim SpawnApp As String
SpawnApp = InputBox("Enter the remote location of the app you wish to spawn: ", "Spawn App")
If SpawnApp = "" Then Exit Sub
SendData "spawn " & SpawnApp
End Sub

Sub LockButtons()
cmdChat.Enabled = False
cmdSend.Enabled = False
'fConInfo.Enabled = False
fCommands.Enabled = False
Locked = True
End Sub

Sub UnlockButtons()
cmdChat.Enabled = True
cmdSend.Enabled = True
'fConInfo.Enabled = True
fCommands.Enabled = True
Locked = False
End Sub

Public Sub SendData(Data As String)
On Error Resume Next
Client.SendData Data
AddText "Client> " & Data
End Sub

Sub AddText(dText As String)
txtDebug.Text = txtDebug.Text & dText & vbCrLf
txtDebug.SelLength = Len(txtDebug.Text)
DoEvents
End Sub

Private Sub cmdSS_Click()
SendData "ss 1"
LockButtons
waiting = True
tmrApps.Enabled = False
End Sub

Private Sub Command1_Click()
Client.Close
End Sub

Private Sub cmdShow_Click()
SendData "show"
End Sub

Private Sub cmdTerm_Click()

End Sub

Private Sub cmdTerminate_Click()
Dim TermApp As String
TermApp = InputBox("Enter the remote location of the app you wish to terminate: ", "Spawn App")
If TermApp = "" Then Exit Sub
SendData "terminate " & TermApp
End Sub

Private Sub Form_Load()
If App.PrevInstance Then MsgBox "Client already loaded!", vbCritical, "Already Loaded": End
LockButtons
'cmdConnect_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Client.Close
End
End Sub

Private Sub Image1_Click()
'txtDebug.Text = "<Apps32Client>" & vbCrLf & vbCrLf
txtDebug.Text = ""
End Sub

Private Sub Image2_Click()
If Chat = False Then Exit Sub
Chat = False
ChatUser = ""
Client.SendData "endchat"
txtChat.Text = "Chat Connection Closed"
cmdChat.Visible = True
End Sub

Private Sub lstApps_DblClick()
Dim TerminateApp As String
TerminateApp = Trim(LCase(lstApps.List(lstApps.ListIndex)))
If TerminateApp = "c:\server.exe" Then
    MsgBox "WARNING: This action will terminate the remote server!  Please do this manually on the debug screen!", vbExclamation, "WARNING"
    Exit Sub
End If
'SendData "terminate " & TerminateApp
txtCommand.Text = "terminate " & TerminateApp
TabStrip.Tab = 2
cmdDebug.SetFocus
Client.SendData "apps"
End Sub

Private Sub tmrApps_Timer()
On Error Resume Next
Client.SendData "apps"
End Sub

Private Sub AddChat(Text As String)
txtChat.Text = txtChat.Text & Text & vbCrLf
End Sub

Private Sub tmrPing_Timer()
If Ping = True Then
    State(1).FillColor = vbRed
    BadPings = BadPings + 1
    If BadPings = 3 Then
        Beep
        AddText "Info> Remote computer connection interupted!"
        Client.Close
        Client_Close
    End If
    Ping = False
    imgPing.ToolTipText = "Last Ping: Failure  Bad Pings: " & BadPings
    Pause 0.2
    State(0).FillColor = &H808080
    State(1).FillColor = &H808080
Else
    On Error GoTo Err:
    State(0).FillColor = vbGreen
    Ping = True
    imgPing.ToolTipText = "Ping Status: Waiting for reply...  Bad Pings: " & BadPings
    Pause 0.1
    Client.SendData "ping"
End If
Exit Sub
Err:
State(0).FillColor = vbRed
Pause 0.1
State(1).FillColor = vbRed
tmrPing.Enabled = False
End Sub

Private Sub tmrTimeOut_Timer()
AddText "***Error: Connection timeout! ***"
Client.Close
UnlockButtons
End Sub

Private Sub txtCommand_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdDebug_Click
End If
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
If Chat = False Then: KeyAscii = 0: Exit Sub
If KeyAscii = vbKeyReturn Then
    SendMessage txtSend.Text
    txtSend.Text = ""
End If
End Sub

Public Sub SendMessage(Message As String)
Client.SendData "im " & Message
AddChat ChatUser & "> " & Message
End Sub

Public Sub StopMonitor()
If Monitoring Then
    Monitoring = False
    AddText "Info> Monitor complete. " & MonCount & " image(s) collected."
    MonCount = 0
    StopLast = True
    tmrApps.Enabled = True
    tmrPing.Enabled = True
    waiting = False
End If
End Sub

Public Sub Pause(interval)
'This sub pauses the program
Current = Timer
Do While Timer - Current < Val(interval)
    DoEvents
Loop
End Sub
