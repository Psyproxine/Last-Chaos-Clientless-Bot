VERSION 5.00
Begin VB.Form frmToolWin 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "ToolWindow"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Menu mnuPopup 
      Caption         =   "&Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuMaximize 
         Caption         =   "&Maximize"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmToolWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_ucFolder As ucFolder
Private m_Folder As Folder

Public Property Get Folder() As ucFolder
   
   Set Folder = m_ucFolder
   
End Property
Public Property Set Folder(ucFolder As ucFolder)
   
   Set m_ucFolder = ucFolder
   
   If Not m_ucFolder Is Nothing Then
      RefreshFolder
   Else
      Unload Me
   End If
   
End Property

Public Property Set FolderModel(ByRef Folder As Folder)
   
   Set m_Folder = Folder
      
End Property

Private Sub Form_Resize()
   
   Dim r As RECT
   Dim w As Long
   Dim h As Long
   
   GetWindowRect Me.hWnd, r
       
   If Me.Visible Then
   If Not m_Folder Is Nothing Then
      With m_Folder.Position
         .Left = r.Left
         .Right = r.Right
         .Top = r.Top
         .Bottom = r.Bottom
      End With
   End If
   End If
   w = r.Right - r.Left
   h = r.Bottom - r.Top
   
   ' If a floating window gets smaller than 100 then resize it.
   If w < 100 Or r.Bottom - r.Top < 100 Then
      
      If w < 100 Then w = 100
      If h < 100 Then h = 100
      
      SetWindowPos Me.hWnd, 0&, r.Left, r.Top, w, h, SWP_FRAMECHANGED
      
   End If
   
   RefreshFolder
   
End Sub

Public Sub RefreshFolder()
   
   On Error Resume Next
   
   Dim r As RECT
   
   If Not m_ucFolder Is Nothing Then
      
      GetWindowRect hWnd, r
      
      SetWindowPos m_ucFolder.hWnd, 0&, 1, 1, r.Right - r.Left - 7, r.Bottom - r.Top - 7, SWP_FRAMECHANGED
      
      m_ucFolder.Refresh
      
      With m_ucFolder
         .LeftPos = 1
         .TopPos = 1
         .RightPos = Me.Width
         .BottomPos = Me.Height
      End With
           
   End If
   
End Sub
