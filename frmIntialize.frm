VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIntialize 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Apps32 Server"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   Icon            =   "frmIntialize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1560
      Top             =   960
   End
   Begin MSWinsockLib.Winsock wSock 
      Left            =   2040
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "You are not currently connected."
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   2325
      End
   End
   Begin VB.CommandButton cmdBegin 
      Caption         =   "Begin"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "255.255.255.255"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblApp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apps32 Server v1.1.1"
      Height          =   195
      Left            =   2880
      TabIndex        =   6
      Top             =   1920
      Width           =   1545
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Your IP Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1185
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmIntialize.frx":0442
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   $"frmIntialize.frx":0884
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmIntialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ForceUnload As Boolean

Private Sub cmdBegin_Click()
If cmdBegin.Caption = "Begin" Then
    cmdBegin.Caption = "Stop"
    Load frmServer
    WindowState = 1
Else
    frmServer.Server.Close
    Unload frmServer
    ForceUnload = True
    Unload Me
    End
End If
End Sub

Private Sub Form_Load()
Dim IP As String

lblApp.FontSize = 6
lblApp.Caption = "Apps32 Server v" & App.Major & "." & App.Minor & "." & App.Revision

IP = wSock.LocalIP
txtIP = IP
If IP = "127.0.0.1" Or IP = "0.0.0.0" Then
    Status "You are not currently connected to the internet!"
    cmdBegin.Enabled = False
ElseIf Mid(IP, 1, 3) = "192" Or Mid(IP, 1, 2) = "10" Then
    Status "You are on a network computer.  Apps32 may not work."
End If
End Sub

Public Sub Status(Text As String)
lblStatus.Caption = Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not ForceUnload And cmdBegin.Caption = "Stop" Then
    Cancel = 1
    WindowState = 1
End If
End Sub

Private Sub Timer1_Timer()
If sConnected And lblStatus.Caption <> "You are connected." Then
    Status "You are connected."
End If
If sListening And lblStatus.Caption <> "You are currently waiting for a connection." Then
    Status "You are currently waiting for a connection."
End If
End Sub
