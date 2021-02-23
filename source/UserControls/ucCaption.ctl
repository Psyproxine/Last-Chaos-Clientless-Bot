VERSION 5.00
Begin VB.UserControl ucCaption 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8160
   ScaleHeight     =   270
   ScaleWidth      =   8160
   ToolboxBitmap   =   "ucCaption.ctx":0000
   Begin VB.PictureBox picCaption 
      Align           =   3  'Links ausrichten
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   7455
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.Image imgIcon 
         Height          =   255
         Left            =   20
         Top             =   0
         Width           =   255
      End
   End
   Begin Aggressive.ucButton btnCloseView 
      Align           =   4  'Rechts ausrichten
      Height          =   270
      Left            =   7905
      TabIndex        =   1
      ToolTipText     =   "Close"
      Top             =   0
      Width           =   255
      _extentx        =   450
      _extenty        =   476
      picture         =   "ucCaption.ctx":0312
      image           =   "VIEW_CLOSE"
   End
   Begin Aggressive.ucButton btnMaximizeView 
      Align           =   4  'Rechts ausrichten
      Height          =   270
      Left            =   7650
      TabIndex        =   2
      ToolTipText     =   "Maximize"
      Top             =   0
      Width           =   255
      _extentx        =   450
      _extenty        =   476
      picture         =   "ucCaption.ctx":062C
      image           =   "VIEW_MAXIMIZE"
   End
End
Attribute VB_Name = "ucCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Active As Boolean ' State of this folder (true = active / false = inactive)
Private m_Caption As String
Private m_MaximizeButton As Boolean

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Public Event CloseView()
Public Event MaximizeView()
Public Event RestoreView()

Public Property Get Active() As Boolean
   Active = m_Active
End Property
Public Property Let Active(ByVal NewActive As Boolean)
   m_Active = NewActive
   btnCloseView.Active = NewActive
   btnMaximizeView.Active = NewActive
End Property

Public Property Get MaximizeButton() As Boolean
   MaximizeButton = m_MaximizeButton
End Property
Public Property Let MaximizeButton(ByVal NewMaximizeButton As Boolean)
   m_MaximizeButton = NewMaximizeButton
   btnMaximizeView.Visible = NewMaximizeButton
End Property

Public Property Get Caption() As String
   Caption = m_Caption
End Property
Public Property Let Caption(ByVal NewCaption As String)
   m_Caption = NewCaption
   Refresh
End Property
Public Property Get Icon() As Picture
   Set Icon = imgIcon.Picture
End Property
Public Property Set Icon(ByVal NewIcon As Picture)
   Set imgIcon.Picture = NewIcon
End Property
Public Function hWnd() As Long
   hWnd = UserControl.hWnd
End Function
Public Function hDc() As Long
   hDc = UserControl.hDc
End Function


Private Sub UserControl_Initialize()
   btnCloseView.ToolTip = "Close"
   btnMaximizeView.ToolTip = "Maximize"
End Sub

Private Sub UserControl_Resize()
      
   On Error Resume Next
      
   Dim w As Long
   
   If btnCloseView.Visible Then w = w + btnCloseView.Width
   If btnMaximizeView.Visible Then w = w + btnMaximizeView.Width
      
   picCaption.Width = ScaleWidth - w
   
   Refresh
   
End Sub

Public Sub Refresh()

   Dim l_Gradient As Gradient
      
   If btnMaximizeView.Visible Then btnMaximizeView.Refresh
   If btnCloseView.Visible Then btnCloseView.Refresh
   
   picCaption.Cls
      
   ' Draw the caption styles
   Select Case m_Scheme.CaptionStyle
      
      Case vbHorizontalGradient, vbVerticalGradient:  ' Draw a gradient caption
      
         Set l_Gradient = New Gradient
         
         With l_Gradient
            
            If m_Active Then
               .Color1 = m_Scheme.ActiveCaptionGradient1
               .Color2 = m_Scheme.ActiveCaptionGradient2
            Else
               .Color1 = m_Scheme.InactiveCaptionGradient1
               .Color2 = m_Scheme.InactiveCaptionGradient2
            End If
              
            If m_Scheme.CaptionStyle = vbHorizontalGradient Then
               .Angle = 0
            Else
               .Angle = 90
            End If
              
            .Draw picCaption
        
        End With
      
   End Select
   
   imgIcon.Visible = m_Scheme.ViewCaptionIcons
   
   If m_Scheme.ViewCaptionIcons Then
      PrintText m_Caption, picCaption.hDc, 19, 1, , 11, fwNormal, IIf(Active, m_Scheme.ActiveCaptionForeColor, m_Scheme.InactiveCaptionForeColor)
   Else
      PrintText m_Caption, picCaption.hDc, 4, 1, , 11, fwNormal, IIf(Active, m_Scheme.ActiveCaptionForeColor, m_Scheme.InactiveCaptionForeColor)
   End If
   
   picCaption.Refresh
      
   Set l_Gradient = Nothing

End Sub

Private Sub picCaption_Click()
   RaiseEvent Click
End Sub

Private Sub picCaption_DblClick()
   RaiseEvent DblClick
End Sub
Private Sub imgIcon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  ' Dim r As RECT
   
  ' GetWindowRect Me.hWnd, r
   
  ' PopupMenu frmToolWin.mnuPopup, , r.Left, r.Top
  'picCaption_MouseDown Button, Shift, x, y
  RaiseEvent Click
End Sub
Private Sub picCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   m_CursorPos.x = x / Screen.TwipsPerPixelX
   m_CursorPos.y = y / Screen.TwipsPerPixelY
   RaiseEvent MouseDown(Button, Shift, x, y)
End Sub
Private Sub picCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub btnCloseView_Click()
   RaiseEvent CloseView
End Sub
Private Sub btnMaximizeView_Click()
   
   If StrComp(btnMaximizeView.Image, "VIEW_MAXIMIZE") = 0 Then
      btnMaximizeView.Image = "VIEW_RESTORE"
      btnMaximizeView.ToolTip = "Restore"
   Else
      btnMaximizeView.Image = "VIEW_MAXIMIZE"
      btnMaximizeView.ToolTip = "Maximize"
   End If
   btnMaximizeView.Refresh
   RaiseEvent MaximizeView
End Sub
