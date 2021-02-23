VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SchemeCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IScheme

' Caption
Private m_ActiveCaptionForeColor As Long
Private m_ActiveCaptionGradient1 As Long
Private m_ActiveCaptionGradient2 As Long
Private m_InactiveCaptionForeColor As Long
Private m_InactiveCaptionGradient1 As Long
Private m_InactiveCaptionGradient2 As Long

' Tabs
Private m_FocusTabForeColor As Long
Private m_FocusTabGradient1 As Long
Private m_FocusTabGradient2 As Long
Private m_FocusTabGradientAngle As Long
Private m_ActiveTabForeColor As Long
Private m_ActiveTabGradient1 As Long
Private m_ActiveTabGradient2 As Long
Private m_ActiveTabGradientAngle As Long
Private m_InactiveTabForeColor As Long
Private m_InactiveTabGradient1 As Long
Private m_InactiveTabGradient2 As Long
Private m_InactiveTabGradientAngle As Long

' Generals
Private m_FrameColor As Long
Private m_BackColor As Long
Private m_FrameWidth As Long
Private m_EditorAreaBackColor As OLE_COLOR
Private m_ViewCaptions As Boolean
Private m_ViewCaptionIcons As Boolean
Private m_CaptionStyle As Long

' See IScheme Interface.
Private Property Get IScheme_ActiveCaptionForeColor() As Long
   IScheme_ActiveCaptionForeColor = m_ActiveCaptionForeColor
End Property
Public Property Let ActiveCaptionForeColor(ByVal New_ActiveCaptionForeColor As OLE_COLOR)
   m_ActiveCaptionForeColor = New_ActiveCaptionForeColor
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveCaptionGradient1() As Long
   IScheme_ActiveCaptionGradient1 = m_ActiveCaptionGradient1
End Property
Public Property Let ActiveCaptionGradient1(ByVal New_ActiveCaptionGradient1 As OLE_COLOR)
   m_ActiveCaptionGradient1 = New_ActiveCaptionGradient1
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveCaptionGradient2() As Long
   IScheme_ActiveCaptionGradient2 = m_ActiveCaptionGradient2
End Property
Public Property Let ActiveCaptionGradient2(ByVal New_ActiveCaptionGradient2 As OLE_COLOR)
   m_ActiveCaptionGradient2 = New_ActiveCaptionGradient2
End Property

' See IScheme Interface.
Private Property Get IScheme_InactiveCaptionForeColor() As Long
   IScheme_InactiveCaptionForeColor = m_InactiveCaptionForeColor
End Property
Public Property Let InactiveCaptionForeColor(ByVal New_InactiveCaptionForeColor As OLE_COLOR)
   m_InactiveCaptionForeColor = New_InactiveCaptionForeColor
End Property

' See IScheme Interface.
Private Property Get IScheme_InactiveCaptionGradient1() As Long
   IScheme_InactiveCaptionGradient1 = m_InactiveCaptionGradient1
End Property
Public Property Let InactiveCaptionGradient1(ByVal New_InactiveCaptionGradient1 As OLE_COLOR)
   m_InactiveCaptionGradient1 = New_InactiveCaptionGradient1
End Property

' See IScheme Interface.
Private Property Get IScheme_InactiveCaptionGradient2() As Long
   IScheme_InactiveCaptionGradient2 = m_InactiveCaptionGradient2
End Property
Public Property Let InactiveCaptionGradient2(ByVal New_InactiveCaptionGradient2 As OLE_COLOR)
   m_InactiveCaptionGradient2 = New_InactiveCaptionGradient2
End Property

' See IScheme Interface.
Private Property Get IScheme_FocusTabForeColor() As Long
   IScheme_FocusTabForeColor = m_FocusTabForeColor
End Property
Public Property Let FocusTabForeColor(ByVal New_FocusTabForeColor As OLE_COLOR)
   m_FocusTabForeColor = New_FocusTabForeColor
End Property

' See IScheme Interface.
Private Property Get IScheme_FocusTabGradient1() As Long
   IScheme_FocusTabGradient1 = m_FocusTabGradient1
End Property
Public Property Let FocusTabGradient1(ByVal New_FocusTabGradient1 As Long)
   m_FocusTabGradient1 = New_FocusTabGradient1
End Property

' See IScheme Interface.
Private Property Get IScheme_FocusTabGradient2() As Long
   IScheme_FocusTabGradient2 = m_FocusTabGradient2
End Property
Public Property Let FocusTabGradient2(ByVal New_FocusTabGradient2 As Long)
   m_FocusTabGradient2 = New_FocusTabGradient2
End Property

' See IScheme Interface.
Private Property Get IScheme_FocusTabGradientAngle() As Long
   IScheme_FocusTabGradientAngle = m_FocusTabGradientAngle
End Property
Public Property Let FocusTabGradientAngle(ByVal New_FocusTabGradientAngle As Long)
   m_FocusTabGradientAngle = New_FocusTabGradientAngle
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveTabForeColor() As Long
   IScheme_ActiveTabForeColor = m_ActiveTabForeColor
End Property
Public Property Let ActiveTabForeColor(ByVal New_ActiveTabForeColor As OLE_COLOR)
   m_ActiveTabForeColor = New_ActiveTabForeColor
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveTabGradient1() As Long
   IScheme_ActiveTabGradient1 = m_ActiveTabGradient1
End Property
Public Property Let ActiveTabGradient1(ByVal New_ActiveTabGradient1 As OLE_COLOR)
   m_ActiveTabGradient1 = New_ActiveTabGradient1
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveTabGradient2() As Long
   IScheme_ActiveTabGradient2 = m_ActiveTabGradient2
End Property
Public Property Let ActiveTabGradient2(ByVal New_ActiveTabGradient2 As OLE_COLOR)
   m_ActiveTabGradient2 = New_ActiveTabGradient2
End Property

' See IScheme Interface.
Private Property Get IScheme_ActiveTabGradientAngle() As Long
   IScheme_ActiveTabGradientAngle = m_ActiveTabGradientAngle
End Property
Public Property Let ActiveTabGradientAngle(ByVal New_ActiveTabGradientAngle As Long)
   m_ActiveTabGradientAngle = New_ActiveTabGradientAngle
End Property

' See IScheme Interface.
Private Property Get IScheme_BackColor() As Long
   IScheme_BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   m_BackColor = New_BackColor
End Property

' See IScheme Interface.
Private Property Get IScheme_FrameColor() As Long
   IScheme_FrameColor = m_FrameColor
End Property
Public Property Let FrameColor(ByVal New_FrameColor As OLE_COLOR)
   m_FrameColor = New_FrameColor
End Property

' See IScheme Interface.
Private Property Get IScheme_FrameWidth() As Long
   IScheme_FrameWidth = m_FrameWidth
End Property
Public Property Let FrameWidth(ByVal New_FrameWidth As Long)
   m_FrameWidth = New_FrameWidth
End Property

' See IScheme Interface.
Private Property Get IScheme_EditorAreaBackColor() As Long
   IScheme_EditorAreaBackColor = m_EditorAreaBackColor
End Property
Public Property Let EditorAreaBackColor(ByVal New_EditorAreaBackColor As OLE_COLOR)
   m_EditorAreaBackColor = New_EditorAreaBackColor
End Property





' See IScheme Interface.
Private Property Get IScheme_InactiveTabForeColor() As Long
   IScheme_InactiveTabForeColor = m_InactiveTabForeColor
End Property
Public Property Let InactiveTabForeColor(ByVal New_InactiveTabForeColor As OLE_COLOR)
   m_InactiveTabForeColor = New_InactiveTabForeColor
End Property

' See IScheme Interface.
Private Property Get IScheme_InactiveTabGradient1() As Long
   IScheme_InactiveTabGradient1 = m_InactiveTabGradient1
End Property
Public Property Let InactiveTabGradient1(ByVal New_InactiveTabGradient1 As OLE_COLOR)
   m_InactiveTabGradient1 = New_InactiveTabGradient1
End Property

' See IScheme Interface.
Private Property Get IScheme_InactiveTabGradient2() As Long
   IScheme_InactiveTabGradient2 = m_InactiveTabGradient2
End Property
Public Property Let InactiveTabGradient2(ByVal New_InactiveTabGradient2 As OLE_COLOR)
   m_InactiveTabGradient2 = New_InactiveTabGradient2
End Property

' See IScheme Interface.
Private Property Get IScheme_InactiveTabGradientAngle() As Long
   IScheme_InactiveTabGradientAngle = m_InactiveTabGradientAngle
End Property
Public Property Let InactiveTabGradientAngle(ByVal New_InactiveTabGradientAngle As Long)
   m_InactiveTabGradientAngle = New_InactiveTabGradientAngle
End Property

' See IScheme Interface.
Private Property Get IScheme_ViewCaptions() As Boolean
   IScheme_ViewCaptions = m_ViewCaptions
End Property
Public Property Let ViewCaptions(ByVal New_ViewCaptions As Boolean)
   m_ViewCaptions = New_ViewCaptions
End Property

' See IScheme Interface.
Private Property Get IScheme_ViewCaptionIcons() As Boolean
   IScheme_ViewCaptionIcons = m_ViewCaptionIcons
End Property
Private Property Let IScheme_ViewCaptionIcons(ByVal New_ViewCaptionIcons As Boolean)
   m_ViewCaptionIcons = New_ViewCaptionIcons
End Property

' See IScheme Interface.
Private Property Get IScheme_CaptionStyle() As vbCaptionStyle
   IScheme_CaptionStyle = m_CaptionStyle
End Property
Public Property Let CaptionStyle(ByVal New_CaptionStyle As vbCaptionStyle)
   m_CaptionStyle = New_CaptionStyle
End Property
