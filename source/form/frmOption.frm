VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Path"
      Height          =   3540
      Left            =   1950
      TabIndex        =   38
      Top             =   2970
      Width           =   4920
      Begin VB.CheckBox Check5 
         Caption         =   "Auto re-connect                  second."
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   300
         Width           =   4245
      End
   End
   Begin VB.Frame fPath 
      Caption         =   "Path"
      Height          =   3540
      Left            =   1950
      TabIndex        =   36
      Top             =   600
      Width           =   4920
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   4380
         TabIndex        =   44
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   3390
         TabIndex        =   43
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   2460
         TabIndex        =   42
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1800
         TabIndex        =   41
         Top             =   240
         Width           =   405
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1200
         TabIndex        =   40
         Top             =   240
         Width           =   345
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Lock, map:        , x:          , y:          , width:         , height:   "
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   300
         Width           =   4695
      End
   End
   Begin VB.ListBox lstmenu 
      Appearance      =   0  'Flat
      Height          =   2580
      IntegralHeight  =   0   'False
      ItemData        =   "frmOption.frx":0000
      Left            =   0
      List            =   "frmOption.frx":0007
      TabIndex        =   2
      Top             =   270
      Width           =   1935
   End
   Begin VB.Frame fBasic 
      Caption         =   "Basic"
      Height          =   3540
      Left            =   1950
      TabIndex        =   24
      Top             =   600
      Width           =   4920
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1140
         TabIndex        =   28
         Top             =   570
         Width           =   645
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Time Out                  second"
         Height          =   225
         Left            =   120
         TabIndex        =   27
         Top             =   630
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1620
         TabIndex        =   26
         Top             =   240
         Width           =   675
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Auto re-connect                  second."
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   300
         Width           =   4245
      End
   End
   Begin VB.Frame fHP 
      Appearance      =   0  'Flat
      Caption         =   "HP /SP"
      ForeColor       =   &H80000008&
      Height          =   3540
      Left            =   1950
      TabIndex        =   3
      Top             =   600
      Width           =   4920
      Begin VB.TextBox txtSitUntilSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3870
         MaxLength       =   3
         TabIndex        =   23
         Text            =   "48"
         Top             =   2160
         Width           =   345
      End
      Begin VB.TextBox txtAutoSitSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1710
         MaxLength       =   3
         TabIndex        =   22
         Text            =   "48"
         Top             =   2160
         Width           =   345
      End
      Begin VB.CheckBox cSitMP 
         Appearance      =   0  'Flat
         Caption         =   "Auto sit when mp <          % and stand when mp >         %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         TabIndex        =   21
         Top             =   2130
         Width           =   4755
      End
      Begin VB.TextBox txtSitUntilHP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3825
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "48"
         Top             =   1905
         Width           =   315
      End
      Begin VB.TextBox txtAutositHP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   20
         Text            =   "48"
         Top             =   1890
         Width           =   465
      End
      Begin VB.CheckBox cSitHP 
         Appearance      =   0  'Flat
         Caption         =   "Auto sit when hp <          % and stand when hp >         %"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   1890
         Width           =   4305
      End
      Begin VB.PictureBox ImgSitNomons 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   60
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   14
         Top             =   1050
         Width           =   165
      End
      Begin VB.TextBox txtHealLv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1230
         MaxLength       =   2
         TabIndex        =   13
         Text            =   "10"
         Top             =   1665
         Width           =   255
      End
      Begin VB.PictureBox imgHeal 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   60
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   12
         Top             =   1680
         Width           =   165
      End
      Begin VB.TextBox txtHealHP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2715
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "48"
         Top             =   1665
         Width           =   255
      End
      Begin VB.TextBox txtItem2HP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "48"
         Top             =   1455
         Width           =   255
      End
      Begin VB.PictureBox imgItem2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   60
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   9
         Top             =   1470
         Width           =   165
      End
      Begin VB.TextBox txtItem2Name 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1065
         TabIndex        =   8
         Text            =   "Red_Potion"
         Top             =   1455
         Width           =   1080
      End
      Begin VB.TextBox txtItem1HP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "48"
         Top             =   1245
         Width           =   255
      End
      Begin VB.PictureBox imgUseItem1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   60
         ScaleHeight     =   165
         ScaleWidth      =   165
         TabIndex        =   5
         Top             =   1260
         Width           =   165
      End
      Begin VB.TextBox txtItem1Name 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1110
         TabIndex        =   4
         Text            =   "Red_Herb"
         Top             =   1245
         Width           =   1035
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Sit when no monster"
         ForeColor       =   &H00FF6600&
         Height          =   180
         Left            =   285
         TabIndex        =   18
         Top             =   1035
         Width           =   1815
      End
      Begin VB.Label LabHeal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto heal lv.       when HP below       % "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF6600&
         Height          =   180
         Left            =   285
         TabIndex        =   17
         Top             =   1665
         Width           =   2880
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto use [                         ] when HP below      % "
         ForeColor       =   &H00FF6600&
         Height          =   195
         Left            =   285
         TabIndex        =   16
         Top             =   1455
         Width           =   3495
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Auto use [                         ] when HP below      % "
         ForeColor       =   &H00FF6600&
         Height          =   195
         Left            =   285
         TabIndex        =   15
         Top             =   1245
         Width           =   3495
      End
   End
   Begin VB.Frame fAttack 
      Caption         =   "Attack"
      Height          =   3540
      Left            =   1950
      TabIndex        =   29
      Top             =   600
      Width           =   4920
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   3990
         TabIndex        =   35
         Top             =   600
         Width           =   405
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2910
         TabIndex        =   34
         Top             =   600
         Width           =   405
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1740
         TabIndex        =   33
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Auto skill, skill ID                 , HP <          % , MP >          %"
         Height          =   225
         Left            =   210
         TabIndex        =   32
         Top             =   660
         Width           =   4395
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2430
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Auto Attack, attack speed                   millisec"
         Height          =   225
         Left            =   210
         TabIndex        =   30
         Top             =   300
         Width           =   3645
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "  Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   1950
      TabIndex        =   0
      Top             =   0
      Width           =   4905
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
