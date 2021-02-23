VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmIRC 
   Caption         =   "IRC "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmIRC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSession 
      Interval        =   10000
      Left            =   2130
      Top             =   1740
   End
   Begin VB.Timer tmrRecom 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1200
      Top             =   2190
   End
   Begin VB.Timer tmrNick 
      Interval        =   3000
      Left            =   1590
      Top             =   990
   End
   Begin VB.ListBox lstNick 
      Height          =   2880
      IntegralHeight  =   0   'False
      Left            =   3120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   30
      Width           =   1545
   End
   Begin VB.TextBox txtMS 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   2910
      Width           =   4665
   End
   Begin RichTextLib.RichTextBox rtbIRC 
      Height          =   2865
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   5054
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      BulletIndent    =   5
      TextRTF         =   $"frmIRC.frx":058A
   End
End
Attribute VB_Name = "frmIRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public WithEvents IRC As CLIRC
Attribute IRC.VB_VarHelpID = -1
Public WithEvents IRCsock As CSocketMaster
Attribute IRCsock.VB_VarHelpID = -1
Const AP = "#AP"

Function RandomString(strLength As Integer)
    Dim NewString As String
    Dim tmpNum As Integer
    
    For tmpNum = 1 To strLength
        NewString = NewString + Chr(Int(Rnd * 70) + 32)
    Next tmpNum
    
    RandomString = NewString
End Function

Private Sub Form_Load()
Set IRCsock = New CSocketMaster
Set IRC = New CLIRC
IRC.SocketHandle = Me.IRCsock
'ConnectToIRC
End Sub

Private Sub ConnectToIRC()
Exit Sub
Randomize Timer
IRCsock.CloseSck
IRC.UserFullName = "Aggressive Powered [LastChaos Bot]"
IRC.UserHostName = "APBOT"
IRC.UserName = "AP" & App.Major & App.Minor & App.Revision
IRC.UserNick = "AP" + Trim(str(Int(Rnd * 11000)))
IRC.UserServName = "APBOT"
IRC.Connect "irc.webchat.org"
End Sub

Private Sub Form_Resize()
On Error Resume Next
rtbIRC.Move 0, 0, Me.ScaleWidth - lstNick.Width, Me.ScaleHeight - txtMS.Height
lstNick.Move rtbIRC.Width, 0, lstNick.Width, Me.ScaleHeight - txtMS.Height
txtMS.Move 0, rtbIRC.Height, Me.ScaleWidth, txtMS.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set IRCsock = Nothing
Set IRC = Nothing
End Sub

Private Sub Irc_Action(nick As String, User As String, Host As String, Target As String, Message As String)
msIRC "[Action] " & nick & " " & Target & " " & Message
End Sub

Private Sub Irc_BAN(srcNick As String, srcUser As String, srcHost As String, Chan As String, tarNick As String, TarUser As String, TarHost As String)
msIRC "Ban: " & srcNick & " " & srcUser & " " & srcHost & " " & tarNick & " " & TarUser & " " & TarHost
End Sub

Private Sub Irc_Connected()
IRC.Join AP
End Sub

Private Sub Irc_DEOP(srcNick As String, srcUser As String, srcHost As String, Chan As String, Target As String)
If Chan = AP Then
    msIRC Chr(3) & "10[Deop]:---> " & srcNick & " set -o " & Target
End If
End Sub

Private Sub Irc_Join(nick As String, User As String, Host As String, Channel As String)
If Channel = AP Then
    msIRC Chr(3) & "10[Join]:---> " & nick & " (" & User & "@" & Host & ")  "
End If
End Sub

Private Sub Irc_Kick(srcNick As String, srcUser As String, srcHost As String, Chan As String, Target As String, Message As String)
If Chan = AP Then
    msIRC Chr(3) & "5[Kick]:---> " & srcNick & " kick " & Target & " (" & Message & ")  "
End If
End Sub

Private Sub IRC_nick(srcNick As String, srcUser As String, srcHost As String, newNick As String)
msIRC Chr(3) & "10[Nick]:---> " & srcNick & " is now known as " & newNick
End Sub

Private Sub Irc_Notice(nick As String, User As String, Host As String, Target As String, Message As String)
If nick = "NickServ" And Target = IRC.UserNick Then
    msIRC "[NickServ]: " & Message
ElseIf nick = "ChanServ" And Target = IRC.UserNick Then
    msIRC "[ChanServ]: " & Message
ElseIf InStr(1, nick, ".", vbTextCompare) > 0 And Target = IRC.UserNick Then
    msIRC "[" & nick & "]: " & Message
Else
    msIRC "[Notice]: " & nick & " " & User & " " & Host & " " & Target & " " & Message
End If
End Sub

Private Sub Irc_OP(srcNick As String, srcUser As String, srcHost As String, Chan As String, TargetNick As String)
If Chan = AP Then
    msIRC Chr(3) & "10[Op]:---> " & srcNick & " set +o " & TargetNick
End If
End Sub

Private Sub IRC_Part(nick As String, User As String, Host As String, Channel As String, Message As String)
If Channel = AP Then
    msIRC Chr(3) & "10[Part]:---> " & nick & " (" & User & "@" & Host & ")  "
    If nick = IRC.UserNick Then
     IRC.CLChan.RemoveChan AP
     IRC.Join AP
    End If
End If
End Sub

Private Sub Irc_PrivateMessage(nick As String, User As String, Host As String, Target As String, Message As String)
If Target = AP Then
    msIRC nick & ": " & Message
ElseIf (Left(Target, 1) <> "#") Then
    msIRC nick & " [" & Target & "]: " & Message
Else
    msIRC nick & " [" & Target & "]: " & Message
End If
End Sub


Private Sub IRC_Quit(nick As String, User As String, Host As String, Message As String)
msIRC Chr(3) & "10[Quit]:---> " & nick & " (" & User & "@" & Host & ")  "
End Sub

Private Sub IRC_Topic(nick As String, User As String, Host As String, Chan As String, Topic As String)
If Chan = AP Then
    msIRC Chr(3) & "10[Topic]:---> " & nick & " set Topic to " & Topic
End If
End Sub

Private Sub Irc_UNBAN(srcNick As String, srcUser As String, srcHost As String, Chan As String, tarNick As String, TarUser As String, TarHost As String)
If Chan = AP Then
    msIRC Chr(3) & "10[Unban]:---> " & srcNick & " set -b " & tarNick
End If
End Sub

Private Sub IRCsock_Close()
msIRC "Close"
End Sub

Private Sub IRC_VOICE(srcNick As String, srcUser As String, srcHost As String, Chan As String, Target As String)
If Chan = AP Then
    msIRC Chr(3) & "10[Voice]:---> " & srcNick & " set +v " & Target
End If
End Sub

Private Sub IRCsock_CloseSck()
tmrRecom.Enabled = False
tmrRecom.Enabled = True
End Sub

Private Sub IRCsock_DataArrival(ByVal bytesTotal As Long)
    Dim tmp As String
    IRCsock.GetData tmp
    IRC.RecvData tmp
End Sub

Private Sub IRCsock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
msIRC Description
tmrRecom.Enabled = False
tmrRecom.Enabled = True
Err.Clear
End Sub

Private Sub IRCsock_Connect()
    IRC.Connected
End Sub

Private Sub ListNick()
Dim u() As String, i As Integer
u = Split(IRC.CLChan.NickList(AP), " ")
lstNick.Clear
For i = 0 To UBound(u)
    If Trim(u(i)) <> "" Then lstNick.AddItem u(i)
Next
End Sub

Private Sub tmrNick_Timer()
ListNick
End Sub

Private Sub tmrRecom_Timer()
ConnectToIRC
tmrRecom.Enabled = False
End Sub

Private Sub tmrSession_Timer()
If IRC.isConnected And Not IRC.IsChan(AP) Then
    IRC.Join AP
ElseIf Not IRC.isConnected Then
    ConnectToIRC
End If
End Sub

Private Sub txtMS_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) And (Trim(txtMS.Text) <> "") Then
    If Left(Trim(txtMS.Text), 1) = "/" Then
        IRC.SendData Mid(Trim(txtMS.Text), 2)
    Else
        If IRC.PrivMSG(AP, Trim(txtMS.Text)) Then msIRC IRC.UserNick & ": " & Trim(txtMS.Text)
    End If
    KeyAscii = 0
    txtMS.Text = ""
End If
End Sub
