VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMap 
   Caption         =   "Map"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7440
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   427
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox fBGMap 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   0
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   0
      Width           =   3375
      Begin VB.PictureBox PicMain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   0
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   187
         TabIndex        =   9
         Top             =   0
         Width           =   2805
         Begin VB.Shape blPeo 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            Height          =   60
            Index           =   0
            Left            =   930
            Shape           =   1  'Square
            Top             =   1230
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Shape sPos 
            BorderColor     =   &H000000FF&
            Height          =   750
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Shape blmons 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            Height          =   60
            Index           =   0
            Left            =   1890
            Top             =   1140
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Shape block 
            BackColor       =   &H00FF0000&
            BackStyle       =   1  'Opaque
            Height          =   60
            Left            =   810
            Top             =   1410
            Width           =   60
         End
         Begin VB.Label lbMons 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   0
            Left            =   1650
            TabIndex        =   14
            Top             =   930
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lbPeo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   0
            Left            =   660
            TabIndex        =   13
            Top             =   930
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Shape blNPC 
            BackColor       =   &H00FF00FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            Height          =   60
            Index           =   0
            Left            =   300
            Top             =   1230
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lbNPC 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            ForeColor       =   &H00FF00FF&
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   12
            Top             =   1020
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Shape blITEM 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00000000&
            Height          =   60
            Index           =   0
            Left            =   240
            Top             =   600
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label lbITEM 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            ForeColor       =   &H00808080&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   11
            Top             =   390
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label lbShop 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label2"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Index           =   0
            Left            =   750
            TabIndex        =   10
            Top             =   690
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   2610
      Width           =   8265
      Begin VB.CheckBox cPeople 
         Caption         =   "People"
         Height          =   195
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   795
      End
      Begin VB.CheckBox cMonster 
         Caption         =   "Monster"
         Height          =   195
         Left            =   870
         TabIndex        =   4
         Top             =   30
         Width           =   915
      End
      Begin VB.CheckBox cNPC 
         Caption         =   "NPC"
         Height          =   195
         Left            =   1770
         TabIndex        =   3
         Top             =   30
         Width           =   675
      End
      Begin VB.CheckBox cTool 
         Caption         =   "Lock"
         Height          =   195
         Left            =   3180
         TabIndex        =   2
         Top             =   30
         Width           =   675
      End
      Begin VB.CheckBox cItem 
         Caption         =   "ITEM"
         Height          =   195
         Left            =   2460
         TabIndex        =   1
         Top             =   30
         Width           =   735
      End
      Begin MSComctlLib.Slider slMap 
         Height          =   225
         Left            =   3900
         TabIndex        =   6
         Top             =   30
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   5
         SelStart        =   3
         Value           =   3
      End
      Begin VB.Label lbMouse 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   5940
         TabIndex        =   16
         Top             =   30
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0F0F0&
         BackStyle       =   0  'Transparent
         Caption         =   "scale"
         Height          =   195
         Left            =   4980
         TabIndex        =   7
         Top             =   30
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long


Private Sub cMonster_Click()
Dim i As Integer
For i = 1 To lbMons.Count - 1
    Monster_Move i
Next
End Sub

Private Sub cNPC_Click()
Dim i As Integer
For i = 1 To lbNPC.Count - 1
    NPC_Move i
Next
End Sub

Private Sub cPeople_Click()
Dim i As Integer
For i = 1 To lbPeo.Count - 1
    People_Move i
Next
End Sub

Public Sub cTool_Click()
sPos.Visible = IIf(Map = opt.path.lock.m, True, False)
If cTool.Value = 1 Then
'sPos.Left = (cur.x / MapScale) - 27
'sPos.Top = (cur.y / MapScale) - 27
opt.path.lock.auto = True
sPos.Visible = True
Else
opt.path.lock.auto = False
sPos.Visible = False
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
fBGMap.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - Frame1.Height
Frame1.Top = fBGMap.Top + fBGMap.Height
slMap_Change
End Sub

Private Sub PicMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Start_Walk CInt(x / MapScale) * 3, CInt(y / MapScale) * 3
'toPoint CInt(x * MapScale, CInt(y * MapScale, 0
End Sub

Private Sub PicMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Pt As Coord
Pt.x = CInt(x / MapScale) * 3
Pt.y = CInt(y / MapScale) * 3
lbMouse.Caption = "Mouse: " & Pt.x & ":" & Pt.y & "  [" & Distant(cur, Pt) & "]"
End Sub

Public Sub slMap_Change()
On Error Resume Next
    MapScale = slMap.Value / 3
    Dim a As Boolean, CurPos As Coord, w As Integer, h As Integer
    If FileExists(App.path & "/map/" & Map & ".jpg") Then
        PicMap.Picture = LoadPicture(App.path & "/map/" & Map & ".jpg")
    Else
        PicMap.Picture = LoadPicture(App.path & "/map/none.jpg")
    End If
    PicMap.Height = CLng(PicMap.Picture.Height / Screen.TwipsPerPixelY)
    
    PicMap.Width = CLng(PicMap.Picture.Width / Screen.TwipsPerPixelX)
    Label1.Caption = "scale 1:" & MapScale
    PicMain.Width = PicMap.Width * MapScale
    PicMain.Height = PicMap.Height * MapScale
    'Clear_All_Lives
    PicMain.AutoRedraw = True '
       a = StretchBlt(PicMain.hdc, 0, 0, _
                          PicMap.Width * MapScale, PicMap.Height * MapScale, PicMap.hdc, _
                          0, 0, PicMap.Width, PicMap.Height, vbSrcCopy)
                          PicMain.Refresh
    Plot_Dot LoginChar(Char).Pos
   'PicMain.AutoRedraw = True
    sPos.Width = CInt(opt.path.lock.w * MapScale) \ 3
    sPos.Height = CInt(opt.path.lock.h * MapScale) \ 3
    sPos.Left = CInt(((opt.path.lock.x) - (opt.path.lock.w / 2)) * MapScale) \ 3
    sPos.Top = CInt(((opt.path.lock.y) - (opt.path.lock.h / 2)) * MapScale) \ 3
    sPos.Visible = IIf(Map = opt.path.lock.m And cTool.Value, True, False)
   cPeople_Click
   cMonster_Click
End Sub
