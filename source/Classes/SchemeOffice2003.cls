VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SchemeOffice2003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IScheme

' See IScheme Interface.
Private Property Get IScheme_ActiveCaptionForeColor() As Long
   IScheme_ActiveCaptionForeColor = vbWhite
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveCaptionGradient1() As Long
   
   Select Case Scheme
      Case VbNormalColor:
           IScheme_ActiveCaptionGradient1 = RGB(0, 84, 227)
      Case VbMetallic:
           IScheme_ActiveCaptionGradient1 = RGB(155, 153, 182)
      Case VbHomeStead:
           IScheme_ActiveCaptionGradient1 = RGB(139, 161, 105)
      Case Else:
           IScheme_ActiveCaptionGradient1 = RGB(10, 36, 106)
   End Select
   
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveCaptionGradient2() As Long
   
   Select Case Scheme
      Case VbNormalColor:
           IScheme_ActiveCaptionGradient2 = RGB(135, 172, 227)
      Case VbMetallic:
           IScheme_ActiveCaptionGradient2 = RGB(237, 237, 243)
      Case VbHomeStead:
           IScheme_ActiveCaptionGradient2 = RGB(197, 209, 161)
      Case Else:
           IScheme_ActiveCaptionGradient2 = RGB(166, 202, 240)
   End Select
   
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveCaptionGradientAngle() As Long
   IScheme_ActiveCaptionGradientAngle = 270
End Property

' See IScheme Interface.
Private Property Get IScheme_InactiveCaptionForeColor() As Long
   IScheme_InactiveCaptionForeColor = vbBlack
End Property

' See IScheme Interface.
Private Property Get IScheme_InactiveCaptionGradient1() As Long
   
   Select Case Scheme
      Case VbNormalColor:
           IScheme_InactiveCaptionGradient1 = RGB(135, 172, 227)
      Case VbMetallic:
           IScheme_InactiveCaptionGradient1 = vbButtonFace
      Case VbHomeStead:
           IScheme_InactiveCaptionGradient1 = RGB(216, 215, 198)
      Case Else:
           IScheme_InactiveCaptionGradient1 = RGB(160, 160, 160)
   End Select

End Property

' See IScheme Interface.
Private Property Get IScheme_InactiveCaptionGradient2() As Long

   Select Case Scheme
      Case VbNormalColor:
           IScheme_InactiveCaptionGradient2 = RGB(227, 239, 252)
      Case VbMetallic:
           IScheme_InactiveCaptionGradient2 = vbWhite
      Case VbHomeStead:
           IScheme_InactiveCaptionGradient2 = RGB(239, 238, 230)
      Case Else:
           IScheme_InactiveCaptionGradient2 = vbButtonFace
   End Select
   
End Property

' See IScheme Interface.
Private Property Get IScheme_InactiveCaptionGradientAngle() As Long
   Select Case Scheme
      Case VbNormalColor, VbMetallic, VbHomeStead:
           IScheme_InactiveCaptionGradientAngle = 90
      Case Else:
           IScheme_InactiveCaptionGradientAngle = 180
   End Select
End Property

' See IScheme Interface.
Private Property Get IScheme_FocusTabForeColor() As Long
   IScheme_FocusTabForeColor = vbBlack
End Property

' See IScheme Interface.
Private Property Get IScheme_FocusTabGradient1() As Long
   
   IScheme_FocusTabGradient1 = RGB(252, 240, 204)
   
End Property

' See IScheme Interface.
Private Property Get IScheme_FocusTabGradient2() As Long
   IScheme_FocusTabGradient2 = RGB(241, 168, 55)
End Property

' See IScheme Interface.
Private Property Get IScheme_FocusTabGradientAngle() As Long
   IScheme_FocusTabGradientAngle = 90
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveTabForeColor() As Long
   IScheme_ActiveTabForeColor = vbBlack
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveTabGradient1() As Long
   
   Select Case Scheme
      Case VbNormalColor:
           IScheme_ActiveTabGradient1 = RGB(227, 239, 252)
      Case VbMetallic:
           IScheme_ActiveTabGradient1 = vbWhite
      Case VbHomeStead:
           IScheme_ActiveTabGradient1 = RGB(239, 238, 230)
      Case Else:
           IScheme_ActiveTabGradient1 = vbButtonFace
   End Select

End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveTabGradient2() As Long

   Select Case Scheme
      Case VbNormalColor:
           IScheme_ActiveTabGradient2 = RGB(135, 172, 227)
      Case VbMetallic:
           IScheme_ActiveTabGradient2 = vbButtonFace
      Case VbHomeStead:
           IScheme_ActiveTabGradient2 = RGB(216, 215, 198)
      Case Else:
           IScheme_ActiveTabGradient2 = RGB(160, 160, 160)
   End Select

End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveTabGradientAngle() As Long
   IScheme_ActiveTabGradientAngle = 90
End Property

' See IScheme Interface.
Private Property Get IScheme_FrameColor() As Long
   IScheme_FrameColor = vbHighlight
End Property

' See IScheme Interface.
Private Property Get IScheme_FrameWidth() As Long
   IScheme_FrameWidth = 20
End Property

' See IScheme Interface.
Private Property Get IScheme_BackColor() As Long
   
IScheme_BackColor = IScheme_InactiveCaptionGradient2

End Property

' See IScheme Interface.
Private Property Get IScheme_EditorAreaBackColor() As Long
   IScheme_EditorAreaBackColor = vbApplicationWorkspace
End Property

Private Property Get IScheme_InactiveTabForeColor() As Long
   IScheme_InactiveTabForeColor = vbBlack
End Property

Private Property Get IScheme_InactiveTabGradient1() As Long
   IScheme_InactiveTabGradient1 = vbWhite
End Property

Private Property Get IScheme_InactiveTabGradient2() As Long
   IScheme_InactiveTabGradient2 = IScheme_BackColor  ' RGB(189, 212, 242)
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
