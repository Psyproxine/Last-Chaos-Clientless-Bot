VERSION 5.00
Begin VB.UserControl ucTab 
   AutoRedraw      =   -1  'True
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1410
   ScaleHeight     =   315
   ScaleWidth      =   1410
   ToolboxBitmap   =   "ucTab.ctx":0000
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   40
      Top             =   30
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   50
      Width           =   540
   End
End
Attribute VB_Name = "ucTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum vbTabState
   STATE_INACTIVE = 0
   STATE_ACTIVE = 1
   STATE_FOCUS = 2
End Enum

Private m_ViewId As String
Private m_Orientation As VbOrientation
Private m_State As vbTabState

Public Property Get Caption() As String
   Caption = lblCaption.Caption
End Property
Public Property Let Caption(ByVal NewCaption As String)
   lblCaption.Caption = NewCaption
End Property

Public Property Get ToolTip() As String
   ToolTip = lblCaption.ToolTipText
End Property
Public Property Let ToolTip(ByVal NewToolTip As String)
   lblCaption.ToolTipText = NewToolTip
   imgIcon.ToolTipText = NewToolTip
End Property

Public Property Get Icon() As Picture
   Set Icon = imgIcon.Picture
End Property
Public Property Set Icon(ByVal NewIcon As Picture)
   Set imgIcon.Picture = NewIcon
End Property
Public Property Get ViewId() As String
   ViewId = m_ViewId
End Property
Public Property Let ViewId(ByVal NewViewId As String)
   m_ViewId = NewViewId
End Property

Public Property Get Orientation() As VbOrientation
   Orientation = m_Orientation
End Property
Public Property Let Orientation(ByVal NewOrientation As VbOrientation)
   m_Orientation = NewOrientation
End Property
Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property
Public Property Get HasDC() As Boolean
   HasDC = UserControl.HasDC
End Property
Public Property Get hDc() As Long
   hDc = UserControl.hDc
End Property
Public Property Get State() As vbTabState
   State = m_State
End Property
Public Property Let State(ByVal NewState As vbTabState)
   m_State = NewState
End Property

Public Sub Refresh()
   
   Dim l_Gradient As Gradient
   Set l_Gradient = New Gradient
   
   'UserControl.AutoRedraw = True
   
   imgIcon.Visible = (State > STATE_INACTIVE Or Orientation = VbOrientationTop)
   
   Select Case State
      
      Case STATE_INACTIVE:
      
           lblCaption.ForeColor = m_Scheme.InactiveTabForeColor
           If Orientation = VbOrientationBottom Then
              lblCaption.Left = 40
           Else
              lblCaption.Left = 300
           End If
      
      Case STATE_ACTIVE:
      
           lblCaption.ForeColor = m_Scheme.ActiveTabForeColor
           lblCaption.Left = 300
           
      Case Else:
      
           lblCaption.ForeColor = m_Scheme.FocusTabForeColor
           lblCaption.Left = 300
            
   End Select
   
   With lblCaption
      UserControl.Width = .Left + .Width + 90
   End With
   
   UserControl.Cls
   Select Case State
      
      Case STATE_INACTIVE:
                      
           'UserControl.BackColor = m_Scheme.BackColor
                      
           If m_Scheme.InactiveTabGradient1 <> m_Scheme.InactiveTabGradient2 Then
              With l_Gradient
                 .Color1 = m_Scheme.InactiveTabGradient1
                 .Color2 = m_Scheme.InactiveTabGradient2
                 .Angle = m_Scheme.InactiveTabGradientAngle
                 If Orientation = VbOrientationTop Then
                    If .Angle >= 180 Then
                       .Angle = .Angle - 180
                    Else
                       .Angle = .Angle + 180
                    End If
                 End If
                 .Draw Me
              End With
           Else
              UserControl.BackColor = m_Scheme.InactiveTabGradient1
           End If
                      
                      
      Case STATE_ACTIVE:
           
           If m_Scheme.ActiveTabGradient1 <> m_Scheme.ActiveTabGradient2 Then
              With l_Gradient
                 .Color1 = m_Scheme.ActiveTabGradient1
                 .Color2 = m_Scheme.ActiveTabGradient2
                 .Angle = m_Scheme.ActiveTabGradientAngle
                 If Orientation = VbOrientationTop Then
                    If .Angle >= 180 Then
                       .Angle = .Angle - 180
                    Else
                       .Angle = .Angle + 180
                    End If
                 End If
                 .Draw Me
              End With
           Else
              UserControl.BackColor = m_Scheme.ActiveTabGradient1
           End If
            
      Case STATE_FOCUS:
                      
           UserControl.Cls
                      
           If m_Scheme.FocusTabGradient1 <> m_Scheme.FocusTabGradient2 Then
              With l_Gradient
                 .Color1 = m_Scheme.FocusTabGradient1
                 .Color2 = m_Scheme.FocusTabGradient2
                 .Angle = m_Scheme.FocusTabGradientAngle
                 If Orientation = VbOrientationTop Then
                    If .Angle >= 180 Then
                       .Angle = .Angle - 180
                    Else
                       .Angle = .Angle + 180
                    End If
                 End If
                 .Draw Me
              End With
           Else
              UserControl.BackColor = m_Scheme.FocusTabGradient1
           End If
           
            Dim r As RECT
            GetWindowRect Me.hWnd, r
            r.Right = r.Right - r.Left - 2
            r.Bottom = r.Bottom - r.Top
            r.Left = 2
            r.Top = 0
            If Orientation = VbOrientationTop Then
               r.Top = r.Top + 2
               r.Bottom = r.Bottom - 1
            Else
               r.Top = r.Top
               r.Bottom = r.Bottom - 2
            End If
            DrawFocusRect Me.hDc, r
                      
           'UserControl.AutoRedraw = False
           
   End Select
   
   If Orientation = VbOrientationTop Then
      If State = STATE_INACTIVE Then
         UserControl.Line (1, 1)-(ScaleWidth - 10, ScaleHeight - 30), m_Scheme.FrameColor, B
         
         UserControl.Line (1, ScaleHeight - 10)-(ScaleWidth, ScaleHeight - 10), m_Scheme.FrameColor
      Else
         UserControl.Line (1, 1)-(ScaleWidth - 10, ScaleHeight), m_Scheme.FrameColor, B
         
         If State = STATE_FOCUS Then
            UserControl.Line (-10, ScaleHeight - 20)-(1, ScaleHeight - 20), m_Scheme.FocusTabGradient2, B
            UserControl.Line (ScaleWidth - 10, ScaleHeight - 20)-(ScaleWidth, ScaleHeight - 20), m_Scheme.FocusTabGradient2, B
         Else
            UserControl.Line (-10, ScaleHeight - 20)-(1, ScaleHeight - 20), m_Scheme.ActiveTabGradient2, B
            UserControl.Line (ScaleWidth - 10, ScaleHeight - 20)-(ScaleWidth, ScaleHeight - 20), m_Scheme.ActiveTabGradient2, B
         End If
      End If
   
      UserControl.Line (-10, 1)-(1, 1), m_Scheme.BackColor, B
      UserControl.Line (ScaleWidth - 10, 1)-(ScaleWidth, 1), m_Scheme.BackColor, B
   
   Else
      If State = STATE_INACTIVE Then
         'UserControl.Line (1, 1)-(ScaleWidth - 10, ScaleHeight - 10), vbWhite, BF
         UserControl.Line (1, 1)-(ScaleWidth - 10, ScaleHeight - 10), m_Scheme.FrameColor, B
         
         'UserControl.Line (1, 1)-(ScaleWidth, 1), m_Scheme.FrameColor
      Else
         UserControl.Line (1, -10)-(ScaleWidth - 10, ScaleHeight - 10), m_Scheme.FrameColor, B
         
         If State = STATE_FOCUS Then
            UserControl.Line (-10, 1)-(1, 1), m_Scheme.FocusTabGradient2, B
            UserControl.Line (ScaleWidth - 10, 1)-(ScaleWidth, 1), m_Scheme.FocusTabGradient2, B
         Else
            UserControl.Line (-10, 1)-(1, 1), m_Scheme.ActiveTabGradient2, B
            UserControl.Line (ScaleWidth - 10, 1)-(ScaleWidth, 1), m_Scheme.ActiveTabGradient2, B
         End If
      End If
   
      UserControl.Line (-10, ScaleHeight - 20)-(1, ScaleHeight - 20), m_Scheme.BackColor, B
      UserControl.Line (ScaleWidth - 10, ScaleHeight - 20)-(ScaleWidth, ScaleHeight - 20), m_Scheme.BackColor, B
   
   End If
   
  ' UserControl.AutoRedraw = False
   
   Set l_Gradient = Nothing
   
End Sub



Private Sub imgIcon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   UserControl_MouseDown Button, Shift, x, y
End Sub
Private Sub imgIcon_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   UserControl_MouseUp Button, Shift, x, y
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   UserControl_MouseDown Button, Shift, x, y
End Sub
Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   UserControl_MouseUp Button, Shift, x, y
End Sub

Private Sub imgIcon_DblClick()
   UserControl_DblClick
End Sub

Private Sub lblCaption_DblClick()
   UserControl_DblClick
End Sub

Private Sub UserControl_DblClick()
   Parent.EventRaise "DblClick", Me.ViewId
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      Parent.EventRaise "DragStart", Me.ViewId
   End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next
   If Button = vbLeftButton Then
      Parent.EventRaise "DragEnd", Me.ViewId
      Parent.EventRaise "Click", Me.ViewId
   End If
End Sub
