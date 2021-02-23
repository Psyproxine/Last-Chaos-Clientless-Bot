VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SchemeVS2005"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSysColor Lib "user32.dll" ( _
   ByVal nIndex As Long _
) As Long

Implements IScheme

' See IScheme Interface.
Private Property Get IScheme_ActiveCaptionForeColor() As Long
   IScheme_ActiveCaptionForeColor = GetSysColor(9)
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveCaptionGradient1() As Long
   IScheme_ActiveCaptionGradient1 = GetSysColor(2)
End Function

' See IScheme Interface.
Private Property Get IScheme_ActiveCaptionGradient2() As Long
   IScheme_ActiveCaptionGradient2 = GetSysColor(27)
End Property

' See IScheme Interface.
Private Property Get IScheme_FocusTabGradientAngle() As Long
   IScheme_FocusTabGradientAngle = 90
End Property

' See IScheme Interface.
Private Property Get IScheme_InactiveCaptionForeColor() As Long
   IScheme_InactiveCaptionForeColor = vbBlack
End Property

' See IScheme Interface.
Private Property Get IScheme_InactiveCaptionGradient1() As Long
   IScheme_InactiveCaptionGradient1 = RGB(204, 199, 186)
End Property

' See IScheme Interface.
Private Property Get IScheme_InactiveCaptionGradient2() As Long
   IScheme_InactiveCaptionGradient2 = RGB(204, 199, 186)
End Property

' See IScheme Interface.
Private Property Get IScheme_FocusTabForeColor() As Long
   IScheme_FocusTabForeColor = vbBlack
End Property

' See IScheme Interface.
Private Property Get IScheme_FocusTabGradient1() As Long
   IScheme_FocusTabGradient1 = vbWhite
End Property

' See IScheme Interface.
Private Property Get IScheme_FocusTabGradient2() As Long
   IScheme_FocusTabGradient2 = vbWhite
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveTabForeColor() As Long
   IScheme_ActiveTabForeColor = vbBlack
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveTabGradient1() As Long
   IScheme_ActiveTabGradient1 = vbWhite
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveTabGradient2() As Long
   IScheme_ActiveTabGradient2 = vbWhite
End Property

' See IScheme Interface.
Private Property Get IScheme_FrameColor() As Long
   IScheme_FrameColor = RGB(172, 168, 153)
End Property

' See IScheme Interface.
Private Property Get IScheme_FrameWidth() As Long
   IScheme_FrameWidth = 0
End Property

' See IScheme Interface.
Private Property Get IScheme_BackColor() As Long
   If Scheme = VbClassic Then
      IScheme_BackColor = RGB(230, 230, 230)
   Else
      IScheme_BackColor = RGB(234, 233, 220)
   End If
End Property

' See IScheme Interface.
Private Property Get IScheme_EditorAreaBackColor() As Long
   IScheme_EditorAreaBackColor = vbApplicationWorkspace
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveTabGradientAngle() As Long
   IScheme_ActiveTabGradientAngle = 270
End Property

Private Property Get IScheme_InactiveTabForeColor() As Long
   IScheme_InactiveTabForeColor = vbBlack
End Property

Private Property Get IScheme_InactiveTabGradient1() As Long
   IScheme_InactiveTabGradient1 = RGB(236, 232, 215) 'RGB(252, 252, 254)
End Property

Private Property Get IScheme_InactiveTabGradient2() As Long
   IScheme_InactiveTabGradient2 = vbWhite
End Property

Private Property Get IScheme_InactiveTabGradientAngle() As Long
   IScheme_InactiveTabGradientAngle = IScheme_ActiveTabGradientAngle
End Property

' See IScheme Interface.
Private Property Get IScheme_ViewCaptions() As Boolean
   IScheme_ViewCaptions = True
End Property

' See IScheme Interface.
Private Property Get IScheme_ViewCaptionIcons() As Boolean
   IScheme_ViewCaptionIcons = True
End Property

' See IScheme Interface.
Private Property Get IScheme_CaptionStyle() As vbCaptionStyle
   IScheme_CaptionStyle = vbVerticalGradient
End Property
