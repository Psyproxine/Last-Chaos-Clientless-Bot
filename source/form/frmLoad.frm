VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmLoad 
   BorderStyle     =   0  'None
   Caption         =   "ygPlus"
   ClientHeight    =   4740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser wbLoad 
      Height          =   2970
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   6285
      ExtentX         =   11086
      ExtentY         =   5239
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image imgtmp 
      Height          =   270
      Index           =   1
      Left            =   0
      Picture         =   "frmLoad.frx":058A
      Top             =   4230
      Width           =   2700
   End
   Begin VB.Image imgtmp 
      Height          =   270
      Index           =   0
      Left            =   0
      Picture         =   "frmLoad.frx":0B12
      Top             =   3900
      Width           =   2700
   End
   Begin VB.Image imgclick 
      Height          =   270
      Left            =   2160
      MouseIcon       =   "frmLoad.frx":10A3
      MousePointer    =   99  'Custom
      Picture         =   "frmLoad.frx":11F5
      Top             =   3390
      Width           =   2700
   End
   Begin VB.Image imgbg 
      Appearance      =   0  'Flat
      Height          =   3750
      Left            =   0
      Picture         =   "frmLoad.frx":1786
      Top             =   0
      Width           =   6750
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Width = imgbg.Width
Me.Height = imgbg.Height

wbLoad.Navigate "http://www.positron.in.th/ygplus/update.php"
End Sub

Private Sub imgbg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub wbLoad_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call ReleaseCapture
Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Private Sub imgclick_Click()
Unload Me
frmMain.Visible = True
End Sub

Private Sub imgclick_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgclick.Picture = imgtmp(1).Picture
End Sub

Private Sub imgclick_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgclick.Picture = imgtmp(0).Picture
End Sub

Private Sub wbLoad_DownloadComplete()
wbLoad.Visible = True
End Sub

Private Sub wbLoad_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
MsgBox "Network Error!"
End Sub
