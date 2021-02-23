VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   Caption         =   "Chat"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstsay 
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.ComboBox lstChatMenu 
      Appearance      =   0  'Flat
      Height          =   315
      IntegralHeight  =   0   'False
      ItemData        =   "frmChat.frx":058A
      Left            =   0
      List            =   "frmChat.frx":058C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtsay 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1260
      MaxLength       =   100
      TabIndex        =   2
      Top             =   2880
      Width           =   3555
   End
   Begin VB.ComboBox cbWhisper 
      Height          =   315
      Left            =   1110
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   1155
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   4154
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmChat.frx":058E
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
CreateChat
End Sub

Private Sub Form_Resize()
On Error Resume Next
rtbChat.Width = Me.ScaleWidth
rtbChat.Height = Me.ScaleHeight - lstChatMenu.Height
lstChatMenu.Top = rtbChat.Height + 15
cbWhisper.Top = lstChatMenu.Top
txtsay.Top = lstChatMenu.Top
If cbWhisper.Visible Then
txtsay.Width = rtbChat.Width - cbWhisper.Left - cbWhisper.Width
Else
txtsay.Width = rtbChat.Width - lstChatMenu.Left - lstChatMenu.Width
End If
End Sub

Private Sub lstChatMenu_Change()
lstChatMenu_Click
End Sub

Private Sub lstChatMenu_Click()
If lstChatMenu.ListIndex = 4 Then
cbWhisper.Visible = True
txtsay.Left = cbWhisper.Left + cbWhisper.Width
txtsay.Width = rtbChat.Width - cbWhisper.Left - cbWhisper.Width
Else
cbWhisper.Visible = False
txtsay.Left = lstChatMenu.Left + lstChatMenu.Width
txtsay.Width = rtbChat.Width - lstChatMenu.Left - lstChatMenu.Width
End If
End Sub

Public Sub SendChat()
SayChat txtsay.Text, lstChatMenu.ListIndex
End Sub

Private Sub lstChatMenu_Validate(Cancel As Boolean)
lstChatMenu_Click
End Sub

Private Sub txtsay_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) And (Trim(txtsay.Text) <> "") Then
    lstsay.AddItem txtsay.Text
KeyAscii = 0
Dim towhisper As String, numsay As Integer
towhisper = cbWhisper.Text
Dim i As Integer
If towhisper <> "" Then
    For i = 0 To cbWhisper.ListCount
        If cbWhisper.List(i) = towhisper Then GoTo z
    Next
    cbWhisper.AddItem towhisper
End If
z:
SendChat
If (lstsay.ListCount > 20) Then lstsay.RemoveItem (0)
    numsay = lstsay.ListCount
    txtsay.Text = ""
ElseIf (KeyAscii = 38) Then
    If (numsay > 0) Then
         txtsay.Text = lstsay.List(numsay - 1)
         numsay = numsay - 1
    End If
ElseIf (KeyAscii = 40) Then
    If (numsay < lstsay.ListCount) Then
         txtsay.Text = lstsay.List(numsay + 1)
         numsay = numsay + 1
    End If
End If
End Sub
