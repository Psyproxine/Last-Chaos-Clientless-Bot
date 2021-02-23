VERSION 5.00
Begin VB.Form frmInvite 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Party Invite"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrdelay 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   1980
      Top             =   0
   End
   Begin Aggressive.chameleonButton btnYes 
      Height          =   315
      Left            =   420
      TabIndex        =   1
      Top             =   600
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "ตอบรับ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmInvite.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Aggressive.chameleonButton btnNo 
      Height          =   315
      Left            =   1890
      TabIndex        =   2
      Top             =   600
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "ยกเลิก"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmInvite.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label txt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1650
      TabIndex        =   0
      Top             =   90
      Width           =   75
   End
End
Attribute VB_Name = "frmInvite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnNo_Click()
SendPacket Chr(&H98) & Chr(&H2), 1
tmrdelay.Enabled = False
Unload Me
End Sub

Private Sub btnYes_Click()
SendPacket Chr(&H98) & Chr(&H1), 1
tmrdelay.Enabled = False
Chat Chr(&H3) & "8ตอบParty : รับปาร์ตี้แล้ว "
Unload Me
End Sub

Private Sub Form_Load()
MakeTopMost hwnd
End Sub

Private Sub tmrdelay_Timer()
SendPacket Chr(&H98) & Chr(&H2), 1
tmrdelay.Enabled = False
Unload Me
End Sub
