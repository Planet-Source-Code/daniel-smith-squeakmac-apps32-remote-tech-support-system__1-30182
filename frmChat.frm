VERSION 5.00
Begin VB.Form frmChat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   285
      Left            =   3960
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   3735
   End
   Begin VB.TextBox txtChat 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
SendMessage txtSend.Text
txtSend.Text = ""
End Sub

Public Sub SendMessage(Message As String)
frmServer.Server.SendData "Server> " & Message
AddText "Server> " & Message
End Sub

Public Sub AddText(Text As String)
txtChat.Text = txtChat.Text & Text & vbCrLf
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmServer.Chat = False
frmServer.ChatUser = ""
frmServer.Server.SendData "Chat Connection Closed."
Unload Me
End Sub
