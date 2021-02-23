VERSION 5.00
Begin VB.Form frmServ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Server"
   ClientHeight    =   2955
   ClientLeft      =   6435
   ClientTop       =   4905
   ClientWidth     =   4665
   Icon            =   "frmserv.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4665
   Begin VB.ListBox lstServ 
      Appearance      =   0  'Flat
      Height          =   2550
      IntegralHeight  =   0   'False
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   2055
   End
   Begin VB.ListBox lstSub 
      Appearance      =   0  'Flat
      Height          =   2550
      IntegralHeight  =   0   'False
      Left            =   2130
      TabIndex        =   0
      Top             =   30
      Width           =   2505
   End
   Begin Aggressive.chameleonButton chbSelect 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   2610
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      BTYPE           =   5
      TX              =   "Select"
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
      MICON           =   "frmserv.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chbSelect_Click()
    If (lstServ.ListIndex > -1 And lstSub.ListIndex > -1) Then lstSub_DblClick
End Sub

Private Sub Form_Load()
MakeTopMost hwnd
End Sub

Private Sub lstServ_Click()
Dim i As Integer, x As Integer
tmpServ = CInt(Left$(lstServ.List(lstServ.ListIndex), InStr(lstServ.List(lstServ.ListIndex), ".")))
lstSub.Clear
For x = 1 To UBound(Chan_List(tmpServ).ChannelX)
 lstSub.AddItem Chan_List(tmpServ).ChannelX(x).id & ". " & Chan_List(tmpServ).ChannelX(x).name & ":" & Chan_List(tmpServ).ChannelX(x).Connection
Next
End Sub

Public Sub lstSub_DblClick()
tmpServ = CInt(Left$(lstServ.List(lstServ.ListIndex), InStr(lstServ.List(lstServ.ListIndex), ".")))

If Not LimitSERV Then
'ไม่จำกัดเซิฟย่อยที่เล่น
    tmpSub = lstSub.ListIndex + 1
    frmBot.Winsock.CloseSck
    Stat "Change Server to " & Chan_List(tmpServ).ChannelX(tmpSub).name & ":" & Chan_List(tmpServ).ChannelX(tmpSub).Connection
    frmBot.Winsock.Connect Chan_List(tmpServ).ChannelX(tmpSub).name, Chan_List(tmpServ).ChannelX(tmpSub).Connection
Else
'่จำกัดเซิฟย่อยที่เล่น
    ConnectToServ tmpServ
End If
Unload Me
End Sub
