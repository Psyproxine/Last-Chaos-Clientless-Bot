VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About AP"
   ClientHeight    =   3015
   ClientLeft      =   2340
   ClientTop       =   1890
   ClientWidth     =   4365
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2081.007
   ScaleMode       =   0  'User
   ScaleWidth      =   4098.96
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Aggressive.chameleonButton cmdOK 
      Height          =   345
      Left            =   3360
      TabIndex        =   3
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   49152
      MPTR            =   1
      MICON           =   "frmAbout.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2006 killzone :: http://curse-x.com/"
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   90
      Picture         =   "frmAbout.frx":001C
      Top             =   750
      Width           =   1200
   End
   Begin VB.Label lblDescription 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LastChaos Bot"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1500
      TabIndex        =   0
      Top             =   1110
      Width           =   1035
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aggressive Powered"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   390
      TabIndex        =   1
      Top             =   180
      Width           =   3615
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1530
      TabIndex        =   2
      Top             =   780
      Width           =   2775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
