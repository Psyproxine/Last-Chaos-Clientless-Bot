VERSION 5.00
Begin VB.Form frmChar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Character"
   ClientHeight    =   2220
   ClientLeft      =   5835
   ClientTop       =   4110
   ClientWidth     =   3765
   Icon            =   "frmChar.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   3765
   Begin VB.ListBox lstChar 
      Appearance      =   0  'Flat
      Height          =   1830
      IntegralHeight  =   0   'False
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   3705
   End
   Begin Aggressive.chameleonButton chbSelect 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   1890
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
      MICON           =   "frmChar.frx":058A
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
Attribute VB_Name = "frmChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chbSelect_Click()
If lstChar.ListIndex > -1 Then lstChar_DblClick
End Sub

Private Sub Form_Load()
MakeTopMost hwnd
End Sub

Private Sub lstChar_Click()
If lstChar.ListIndex > -1 Then
    Char = CInt(Left$(lstChar.List(lstChar.ListIndex), InStr(lstChar.List(lstChar.ListIndex), "."))) - 1
    Update_Stats
End If
End Sub

Public Sub lstChar_DblClick()

Char = CInt(Left$(lstChar.List(lstChar.ListIndex), InStr(lstChar.List(lstChar.ListIndex), "."))) - 1
SendPacket Chr(&H84) & Chr(&H2) & LoginChar(Char).ID, 1
Stat "Select Character [" & LoginChar(Char).name & " Lv." & LoginChar(Char).level & "]"
Constate = 3
'Update_Stats
Unload Me
End Sub
