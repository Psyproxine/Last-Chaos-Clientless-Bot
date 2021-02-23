VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Aggressive Powered"
   ClientHeight    =   5760
   ClientLeft      =   2850
   ClientTop       =   2865
   ClientWidth     =   10500
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   10500
   Begin Aggressive.ucPerspective ucMain 
      Align           =   1  'Align Top
      Height          =   2235
      Left            =   0
      TabIndex        =   43
      Top             =   1440
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   3942
      ViewCaptionIcons=   0   'False
   End
   Begin VB.PictureBox TRAY 
      CausesValidation=   0   'False
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   16
      ToolTipText     =   "Aggressive Powered"
      Top             =   5040
      Visible         =   0   'False
      Width           =   300
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   635
      ButtonWidth     =   1138
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         Height          =   240
         IMEMode         =   3  'DISABLE
         Left            =   1950
         PasswordChar    =   "*"
         TabIndex        =   32
         Top             =   60
         Width           =   1125
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   420
         TabIndex        =   31
         Top             =   60
         Width           =   1065
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3570
         ScaleHeight     =   225
         ScaleWidth      =   765
         TabIndex        =   28
         Top             =   45
         Width           =   795
         Begin VB.ComboBox cbLocation 
            Appearance      =   0  'Flat
            Height          =   315
            IntegralHeight  =   0   'False
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   -45
            Width           =   840
         End
      End
      Begin Aggressive.chameleonButton btnSit 
         Height          =   300
         Left            =   9720
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   15
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   529
         BTYPE           =   5
         TX              =   "Stand"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":2D12
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Aggressive.chameleonButton btnActive 
         Height          =   300
         Left            =   8760
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   15
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   529
         BTYPE           =   5
         TX              =   "Active"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":2D2E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6900
         ScaleHeight     =   225
         ScaleWidth      =   615
         TabIndex        =   22
         Top             =   45
         Width           =   645
         Begin VB.ComboBox cbChar 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmMain.frx":2D4A
            Left            =   -30
            List            =   "frmMain.frx":2D4C
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   -45
            Width           =   690
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5880
         ScaleHeight     =   225
         ScaleWidth      =   675
         TabIndex        =   20
         Top             =   45
         Width           =   705
         Begin VB.ComboBox cbWorld 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   -45
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4740
         ScaleHeight     =   225
         ScaleWidth      =   765
         TabIndex        =   18
         Top             =   45
         Width           =   795
         Begin VB.ComboBox cbServer 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   -30
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   -45
            Width           =   840
         End
      End
      Begin Aggressive.chameleonButton btnConnect 
         Height          =   300
         Left            =   7590
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Connect / Disconnect"
         Top             =   15
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         BTYPE           =   5
         TX              =   "Connect"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmMain.frx":2D4E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   1
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   $"frmMain.frx":3068
         Top             =   90
         Width           =   7815
      End
   End
   Begin MSComctlLib.Toolbar tbChar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   635
      ButtonWidth     =   1138
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      Begin VB.TextBox txtWalk 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   9270
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Walking"
         Top             =   75
         Width           =   735
      End
      Begin VB.TextBox txtLv 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Lv:"
         Top             =   75
         Width           =   585
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6330
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "EXP:"
         Top             =   75
         Width           =   360
      End
      Begin Aggressive.vbalProgressBar pbCEXP 
         Height          =   210
         Left            =   6720
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   75
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   370
         Picture         =   "frmMain.frx":30F1
         BackColor       =   16777215
         ForeColor       =   49152
         BarColor        =   49152
         BarPicture      =   "frmMain.frx":310D
         BarPictureMode  =   0
         BackPictureMode =   0
         Value           =   50
         ShowText        =   -1  'True
         Text            =   "50%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Aggressive.vbalProgressBar pbCMP 
         Height          =   210
         Left            =   4650
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   75
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   370
         Picture         =   "frmMain.frx":3129
         BackColor       =   16777215
         ForeColor       =   16737792
         BarColor        =   16711680
         BarPicture      =   "frmMain.frx":3145
         BarPictureMode  =   0
         BackPictureMode =   0
         Value           =   50
         ShowText        =   -1  'True
         Text            =   "100 / 200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Aggressive.vbalProgressBar pbCHP 
         Height          =   210
         Left            =   2700
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   75
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   370
         Picture         =   "frmMain.frx":3161
         BackColor       =   16777215
         ForeColor       =   255
         BarColor        =   255
         BarPicture      =   "frmMain.frx":317D
         BarPictureMode  =   0
         BackPictureMode =   0
         Value           =   50
         ShowText        =   -1  'True
         Text            =   "50%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtHP 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "HP:"
         Top             =   75
         Width           =   270
      End
      Begin VB.TextBox txtChar 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Char:"
         Top             =   75
         Width           =   1965
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4350
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "MP:"
         Top             =   75
         Width           =   270
      End
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   5475
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
            Text            =   "Aggressive Powered"
            TextSave        =   "Aggressive Powered"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4480
            MinWidth        =   4480
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   88194
            MinWidth        =   88194
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMon 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   635
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      Begin VB.TextBox txtPos 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   8760
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   90
         Width           =   1695
      End
      Begin MSComctlLib.Toolbar tbAction 
         Height          =   330
         Left            =   5700
         TabIndex        =   26
         Top             =   15
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   582
         ButtonWidth     =   2223
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Basic Action"
               Object.ToolTipText     =   "Basic Action"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   7
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "ทักทาย"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Text            =   "Party Invite"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Text            =   "Trade Invite"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Text            =   "Move Type-Run/Walk"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Emotion"
               Object.ToolTipText     =   "Emotion"
               Object.Tag             =   "Emotion"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   11
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "ทักทาย"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "ร่าเริง"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "เสียใจ"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "ยอดเยี่ยม"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "ปรบมือ"
                  EndProperty
                  BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "ปฎิเสธ"
                  EndProperty
                  BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "โอ้อวด"
                  EndProperty
                  BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "ตำหนิ"
                  EndProperty
                  BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "เชียร์ "
                  EndProperty
                  BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "ท้าทาย"
                  EndProperty
                  BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "เคารพ"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3630
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "HP:"
         Top             =   75
         Width           =   270
      End
      Begin Aggressive.vbalProgressBar pbCMon 
         Height          =   210
         Left            =   3930
         TabIndex        =   11
         Top             =   60
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   370
         Picture         =   "frmMain.frx":3199
         BackColor       =   16777215
         ForeColor       =   255
         BarColor        =   255
         BarPicture      =   "frmMain.frx":31B5
         BarPictureMode  =   0
         BackPictureMode =   0
         Value           =   50
         ShowText        =   -1  'True
         Text            =   "50%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtTarget 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   0
         TabIndex        =   4
         Text            =   "Target:"
         Top             =   90
         Width           =   3555
      End
   End
   Begin MSComctlLib.Toolbar tbPet 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   33
      Top             =   720
      Visible         =   0   'False
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   635
      ButtonWidth     =   1138
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      Begin Aggressive.vbalProgressBar pPetFL 
         Height          =   210
         Left            =   9330
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   370
         Picture         =   "frmMain.frx":31D1
         BackColor       =   16777215
         ForeColor       =   8388736
         BarColor        =   12583104
         BarPicture      =   "frmMain.frx":31ED
         BarPictureMode  =   0
         BackPictureMode =   0
         Value           =   50
         ShowText        =   -1  'True
         Text            =   "50%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9030
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "TM:"
         Top             =   90
         Width           =   270
      End
      Begin VB.TextBox txtPet 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "Pet:"
         Top             =   90
         Width           =   2775
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "HG:"
         Top             =   75
         Width           =   270
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "HP:"
         Top             =   75
         Width           =   270
      End
      Begin Aggressive.vbalProgressBar pPetHP 
         Height          =   210
         Left            =   3150
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   75
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   370
         Picture         =   "frmMain.frx":3209
         BackColor       =   16777215
         ForeColor       =   255
         BarColor        =   255
         BarPicture      =   "frmMain.frx":3225
         BarPictureMode  =   0
         BackPictureMode =   0
         Value           =   50
         ShowText        =   -1  'True
         Text            =   "50%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Aggressive.vbalProgressBar pPetEN 
         Height          =   210
         Left            =   5100
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   75
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   370
         Picture         =   "frmMain.frx":3241
         BackColor       =   16777215
         ForeColor       =   32896
         BarColor        =   49344
         BarPicture      =   "frmMain.frx":325D
         BarPictureMode  =   0
         BackPictureMode =   0
         Value           =   50
         ShowText        =   -1  'True
         Text            =   "100 / 200"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Aggressive.vbalProgressBar pPetEXP 
         Height          =   210
         Left            =   7170
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   75
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   370
         Picture         =   "frmMain.frx":3279
         BackColor       =   16777215
         ForeColor       =   49152
         BarColor        =   49152
         BarPicture      =   "frmMain.frx":3295
         BarPictureMode  =   0
         BackPictureMode =   0
         Value           =   50
         ShowText        =   -1  'True
         Text            =   "50%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "EXP:"
         Top             =   75
         Width           =   360
      End
   End
   Begin VB.Menu mnfile 
      Caption         =   "&File"
   End
   Begin VB.Menu mntool 
      Caption         =   "&Tool"
      Begin VB.Menu mnuoption 
         Caption         =   "&Options"
         Begin VB.Menu mnurefresh 
            Caption         =   "Refresh"
         End
         Begin VB.Menu mnuedit 
            Caption         =   "Edit"
         End
      End
      Begin VB.Menu mnuhide 
         Caption         =   "Hide"
      End
   End
   Begin VB.Menu mnunpc 
      Caption         =   "&Window"
      Begin VB.Menu mnuontop 
         Caption         =   "OnTop"
      End
      Begin VB.Menu mnudebug 
         Caption         =   "Debug"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnnpc 
         Caption         =   "NPC"
      End
      Begin VB.Menu mnuchat 
         Caption         =   "Chat"
      End
      Begin VB.Menu mnuinfo 
         Caption         =   "Info"
      End
      Begin VB.Menu mnuinv 
         Caption         =   "Inv"
      End
      Begin VB.Menu mnuitem 
         Caption         =   "Item"
      End
      Begin VB.Menu mnupeople 
         Caption         =   "People"
      End
      Begin VB.Menu mnuparty 
         Caption         =   "Party"
      End
      Begin VB.Menu mnumonster 
         Caption         =   "Monster"
      End
      Begin VB.Menu mnuskill 
         Caption         =   "Skill"
      End
      Begin VB.Menu mnustat 
         Caption         =   "Stat"
      End
   End
   Begin VB.Menu mnColor 
      Caption         =   "ColorScheme"
      Begin VB.Menu mnStyle 
         Caption         =   "WindowXP"
         Index           =   0
      End
      Begin VB.Menu mnStyle 
         Caption         =   "WindowVista"
         Index           =   1
      End
      Begin VB.Menu mnStyle 
         Caption         =   "Office2003"
         Index           =   2
      End
      Begin VB.Menu mnStyle 
         Caption         =   "Eclipse3"
         Index           =   3
      End
      Begin VB.Menu mnStyle 
         Caption         =   "VS2005"
         Index           =   4
      End
   End
   Begin VB.Menu mnuwarp 
      Caption         =   "Warp"
      Begin VB.Menu mnuwarp1 
         Caption         =   "juno"
      End
      Begin VB.Menu mnudun1 
         Caption         =   "Dun1"
      End
      Begin VB.Menu mnuwarp2 
         Caption         =   "dratan"
      End
      Begin VB.Menu mnudun2 
         Caption         =   "Dun2"
      End
      Begin VB.Menu mnuwarp3 
         Caption         =   "murruc"
      End
   End
   Begin VB.Menu mnupeo 
      Caption         =   "People"
      Visible         =   0   'False
      Begin VB.Menu mnurparty 
         Caption         =   "Request Party"
         Begin VB.Menu mnuparty0 
            Caption         =   "Share All"
         End
         Begin VB.Menu mnuparty1 
            Caption         =   "Share Exp"
         End
         Begin VB.Menu mnuparty2 
            Caption         =   "Share Money"
         End
      End
   End
   Begin VB.Menu mnu2inv 
      Caption         =   "inv"
      Visible         =   0   'False
      Begin VB.Menu mnuuse 
         Caption         =   "Use"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEquip 
         Caption         =   "Equip/UnEquip"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnudrop 
         Caption         =   "Drop"
      End
      Begin VB.Menu mnitemre 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mns 
         Caption         =   "-"
      End
      Begin VB.Menu mnPet 
         Caption         =   "ปล่อยสัตว์เลี้ยง"
         Index           =   1
      End
      Begin VB.Menu mnPet 
         Caption         =   "เก็บสัตว์เลี้ยง"
         Index           =   2
      End
      Begin VB.Menu mngetfeeden 
         Caption         =   "ให้อาหารสัตว์ (หิว)"
      End
      Begin VB.Menu mngetfeedhp 
         Caption         =   "ให้อาหารสัตว์ (เลือด)"
      End
   End
   Begin VB.Menu mnhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnabout 
         Caption         =   "About AP"
      End
   End
   Begin VB.Menu mnutask 
      Caption         =   "&task"
      Visible         =   0   'False
      Begin VB.Menu mnushow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub btnConnect_Click()
    Dim Check As Boolean
    If (Trim(txtUser.Text) = "" Or Trim(txtPass.Text) = "" Or Trim(cbLocation.Text) = "") Then Exit Sub
    If Trim(frmPermission.txtPer_User.Text) = "" Or Trim(frmPermission.txtPer_Pass.Text) = "" Then
        ucMain.ShowView "Permission"
        Connecting False
        frmBot.Winsock.CloseSck
        MsgBox "กรุณากรอก User / Pass ที่ใช้ในเวบ positron.in.th"
        frmPermission.txtPer_User.SetFocus
        Exit Sub
    ElseIf (btnConnect.Caption = "Connect") Then
        Connecting True
        AddAccount
        Bot_Connect
    Else
            Connecting False
            frmBot.Winsock.CloseSck
    End If
    frmBot.tmrRecon.Enabled = False
End Sub

Private Sub btnSit_Click()
If btnSit.Caption = "Stand" Then
    SitDown
Else
    StandUp
End If
End Sub

Private Sub cbLocation_Click()
Dim i As Integer
cbServer.Clear
cbServer.AddItem "Manual"
cbServer.ListIndex = 0
For i = 1 To Location(cbLocation.ListIndex).Serv
    cbServer.AddItem CStr(i)
    If Server = i Then cbServer.ListIndex = i
Next
cbWorld.Clear
cbWorld.AddItem "Manual"
cbWorld.ListIndex = 0
For i = 1 To 6
    cbWorld.AddItem CStr(i)
    If World = i Then cbWorld.ListIndex = i
Next
cbChar.Clear
cbChar.AddItem "Manual"
cbChar.ListIndex = 0
For i = 1 To 5
    cbChar.AddItem CStr(i)
    If Char = i Then cbChar.ListIndex = i
Next
End Sub

Private Sub btnActive_Click()
    If (btnActive.Caption = "Active") Then
        InActive True
    Else
        InActive False
    End If
End Sub



Private Sub Form_Load()
'LimitSERV = True
'LimitEXE = True
LimitLEVEL = 999


Check_AP
ReDim tmppacket(0)
CreateStyle
Load_Account
Form_Resize
AddTray
Clear_Array
Load_Item
Load_Monster
Load_Skill
LoadColors
Load_Option
InActive True
frmBot.tmrSession.Enabled = True
MapScale = 3
frmMap.slMap_Change
Caption = Version
frmStat.wbLoad.Navigate "http://curse-x.com/"
tmrRep = 600
End Sub

Public Sub Form_Resize()
On Error Resume Next
ucMain.Move ucMain.Left, ucMain.Top, Me.ScaleWidth, Me.ScaleHeight - (tbMon.Top + tbMon.Height) - sbMain.Height
ucMain.Refresh
End Sub

Public Sub Form_Unload(Cancel As Integer)
    If MsgBox("Exit AP?", vbYesNo, "AP") <> vbYes Then
        Cancel = 1
        Exit Sub
    End If
     Clear_Array
    Set frmBot.Winsock = Nothing
    Set frmIRC.IRC = Nothing
    Set frmIRC.IRCsock = Nothing
    ucMain.Terminate
    DelTray
    Cancel = 0
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
    End
    End
    End
End Sub

Private Sub mnabout_Click()
frmAbout.show 1, Me
End Sub

Private Sub mngetfeeden_Click()
'90 00 00 07 02 00 08 AC 02 00 00 00 00
SendFeedPet_EN
End Sub

Private Sub mngetfeedhp_Click()
SendFeedPet_HP
End Sub

Private Sub mnitemre_Click()
Send_Refresh_INV
End Sub

Private Sub mnPet_Click(Index As Integer)
ChangPet (Index)
End Sub

Public Sub mnStyle_Click(Index As Integer)
Dim i As Integer
ucMain.ColorScheme = Index
For i = 0 To mnStyle.Count - 1
    mnStyle.Item(i).Checked = False
Next
mnStyle.Item(Index).Checked = True
End Sub

Private Sub mnuchat_Click()
   frmMain.ucMain.ShowView ("Chat")
End Sub

Private Sub mnudebug_Click()
frmDebug.Visible = Not frmDebug.Visible
End Sub

Private Sub mnnpc_Click()
   frmMain.ucMain.ShowView ("NPC")
End Sub

Private Sub mnudrop_Click()
'90 02 00 0E 02 00 37 F8 B2 00 00 00 00 00 00 00 01
On Error GoTo z
Dim i As Integer, tmpVar As Long, str As String
For i = 0 To UBound(Inv) - 1
If ChrtoHex(Inv(i).id) = frmInv.lvInv.ListItems(frmInv.lvInv.SelectedItem.Index).Text Then
    tmpVar = CLng(InputBox("ต้องการทิ้ง " & Inv(i).name & " จำนวน?", "Drop ITEM", 0))
    If tmpVar > 0 Then
        str = Chr(&H90) & Chr(&H2) & Chr(&H0) & Chr(Inv(i).row) & Chr(Inv(i).col) & Inv(i).id & _
                                  Chr(&H0) & Chr(&H0) & Chr(&H0) & Chr(&H0) & LngToChr(tmpVar)
        SendPacket str, 1
        'MsgBox ChrtoHex(str)
    End If
End If
Next
z:
End Sub

Private Sub mnudun1_Click()
    Gomap 1
End Sub

Private Sub mnudun2_Click()
Gomap 3
End Sub

Private Sub mnuedit_Click()
Shell "notepad " & App.path & "/config/options.ini", vbNormalFocus
End Sub

Private Sub mnuEquip_Click()
Dim i As Integer
For i = 0 To UBound(Inv) - 1
If ChrtoHex(Inv(i).id) = frmInv.lvInv.ListItems(frmInv.lvInv.SelectedItem.Index).Text Then
    'SendPacket ChrtoHex(&H10) & ChrtoHex(&H5) & ChrtoHex(&H1) & ChrtoHex(&H0) & ChrtoHex(&H0) & ChrtoHex(&H1) & Inv(i).ID
    Exit Sub
End If
Next
End Sub

Private Sub mnuexit_Click()
Form_Unload 0
End Sub

Private Sub mnuhide_Click()
frmMain.Hide
End Sub

Private Sub mnuinfo_Click()
   frmMain.ucMain.ShowView ("Info")
End Sub

Private Sub mnuinv_Click()
   frmMain.ucMain.ShowView ("Inv")
End Sub

Private Sub mnuitem_Click()
   frmMain.ucMain.ShowView ("Item")
End Sub

Private Sub mnumonster_Click()
   frmMain.ucMain.ShowView ("Mons")
End Sub

Private Sub mnuontop_Click()
If mnuontop.Checked Then
    mnuontop.Checked = False
    MakeNormal hwnd
Else
    mnuontop.Checked = True
    MakeTopMost hwnd
End If
End Sub

Private Sub mnuparty_Click()
   frmMain.ucMain.ShowView ("Party")
End Sub

Private Sub mnuparty0_Click()
If frmPeople.lvPlayer.SelectedItem.Index > 0 Then InviteParty frmPeople.lvPlayer.ListItems(frmPeople.lvPlayer.SelectedItem.Index).SubItems(1), HextoChr(frmPeople.lvPlayer.ListItems(frmPeople.lvPlayer.SelectedItem.Index).Text), 0
End Sub

Private Sub mnuparty1_Click()
If frmPeople.lvPlayer.SelectedItem.Index > 0 Then InviteParty frmPeople.lvPlayer.ListItems(frmPeople.lvPlayer.SelectedItem.Index).SubItems(1), HextoChr(frmPeople.lvPlayer.ListItems(frmPeople.lvPlayer.SelectedItem.Index).Text), 1
End Sub

Private Sub mnuparty2_Click()
If frmPeople.lvPlayer.SelectedItem.Index > 0 Then InviteParty frmPeople.lvPlayer.ListItems(frmPeople.lvPlayer.SelectedItem.Index).SubItems(1), HextoChr(frmPeople.lvPlayer.ListItems(frmPeople.lvPlayer.SelectedItem.Index).Text), 2
End Sub

Private Sub mnupeople_Click()
   frmMain.ucMain.ShowView ("People")
End Sub

Private Sub mnurefresh_Click()
Load_Option
End Sub

Private Sub mnushow_Click()
    Me.show
End Sub

Private Sub mnuskill_Click()
   frmMain.ucMain.ShowView ("Skill")
End Sub

Private Sub mnustat_Click()
   frmMain.ucMain.ShowView ("Stat")
End Sub

Private Sub mnuwarp1_Click()
    Gomap 0
End Sub

Private Sub mnuwarp2_Click()
    Gomap 4
End Sub

Private Sub mnuwarp3_Click()
    Gomap 7
End Sub

Private Sub tbAction_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim X As String
    If ButtonMenu.Parent.Index = 1 Then

    ElseIf ButtonMenu.Parent.Index = 3 Then
        Select Case ButtonMenu.Index
            Case 1
                X = Chr(&HA)
            Case 2
                X = Chr(&HB)
            Case 3
                X = Chr(&HD)
            Case 4
                X = Chr(&HE)
            Case 5
                X = Chr(&HF)
            Case 6
                X = Chr(&H10)
            Case 7
                X = Chr(&H11)
            Case 8
                X = Chr(&H13)
            Case 9
                X = Chr(&H14)
            Case 10
                X = Chr(&H15)
            Case 11
                X = Chr(&H16)
        End Select
        '21 00 02 BA C6 00 0A 00
        SendPacket Chr(&HA1) & LoginChar(Char).id & Chr(&H0) & X & Chr(&H0), 1
    End If
End Sub

Private Sub TRAY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
    Dim Message As Long, SlockPass2 As String
    Message = X / Screen.TwipsPerPixelX
    Select Case Message
        Case WM_LBUTTONDBLCLK
            frmMain.show
            frmMain.Visible = True
            frmMain.Visible = True
        Case WM_RBUTTONDOWN
            PopupMenu mnutask
    End Select
End Sub

Private Sub ucMain_CloseView(ByVal ViewId As String, Cancel As Boolean)
If (ViewId = "Stat" Or ViewId = "Chat" Or ViewId = "IRC") Then MsgBox "Can't Close: " & ViewId, vbOKOnly, "Window":       Cancel = True
End Sub
