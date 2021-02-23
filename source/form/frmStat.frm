VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmStat 
   Caption         =   "Stat"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "frmStat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wbLoad 
      Height          =   3090
      Left            =   3450
      TabIndex        =   1
      Top             =   30
      Width           =   3015
      ExtentX         =   5318
      ExtentY         =   5450
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
   Begin RichTextLib.RichTextBox rtbStat 
      Height          =   3165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   5583
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmStat.frx":058A
   End
End
Attribute VB_Name = "frmStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
rtbStat.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
wbLoad.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub wbLoad_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
wbLoad.Visible = False
rtbStat.Visible = True
End Sub
