VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAccount 
   Caption         =   "Account"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2325
   Icon            =   "frmAccount.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   2325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fAcc 
      Caption         =   "User"
      Height          =   2145
      Left            =   0
      TabIndex        =   1
      Top             =   1860
      Width           =   2325
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   960
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   1410
         Width           =   1305
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   960
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1305
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   960
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   750
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   870
         TabIndex        =   11
         Top             =   450
         Width           =   1395
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   870
         TabIndex        =   10
         Top             =   150
         Width           =   1395
      End
      Begin L2Plus.chameleonButton chbDel 
         Height          =   315
         Left            =   1530
         TabIndex        =   9
         Top             =   1770
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   "Del"
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
         MICON           =   "frmAccount.frx":058A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin L2Plus.chameleonButton chbEdit 
         Height          =   315
         Left            =   810
         TabIndex        =   8
         Top             =   1770
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   "Edit"
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
         MICON           =   "frmAccount.frx":05A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin L2Plus.chameleonButton chbAdd 
         Height          =   315
         Left            =   90
         TabIndex        =   7
         Top             =   1770
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   "Add"
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
         MICON           =   "frmAccount.frx":05C2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "L2 Char:"
         Height          =   195
         Left            =   330
         TabIndex        =   6
         Top             =   1500
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "L2 Server:"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   1170
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "L2 Location:"
         Height          =   195
         Left            =   60
         TabIndex        =   4
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   510
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         Height          =   195
         Left            =   60
         TabIndex        =   2
         Top             =   240
         Width           =   765
      End
   End
   Begin MSComctlLib.TreeView tvAcc 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   3201
      _Version        =   393217
      Indentation     =   18
      Style           =   7
      Appearance      =   0
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
