VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' <B>Description</B><BR />
' <BR />
' This interface could be used to create custom color schemes. The
' implemented properties should return the different colors or constants
' to draw the controls GUI elements. You can set the scheme by calling
' the SetScheme() function of the perspective control.<BR />
' <BR />
' <B>Terms of Use</B><BR />
' <BR />
' You may freely distribute the VbEclipse project and the
' test sample project. You're free to use the VbEclipse control
' in your applications and distribute it with the source code
' of your applications.<BR />
' <BR />
' However, you may not claim that you've written the VbEclipse
' or the Test project.<BR />
' <BR />
' For questions, contact me at vbeclipseproject ab-software.com<BR />
' <BR />
' For more information about the VbEclipse Control, visit the
' following URL: http://www.ab-software.com<BR />
' <BR />
' @See http://www.ab-software.com

Option Explicit


' Returns the text color of an active view caption.
'
' @ActiveCaptionForeColor The text color of an active caption.
Public Property Get ActiveCaptionForeColor() As Long: End Property

' Returns the beginning gradient color of an active view caption.
'
' @ActiveCaptionGradient1 The beginning gradient color of an active caption.
Public Property Get ActiveCaptionGradient1() As Long: End Property

' Returns the ending gradient color of an active view caption.
'
' @ActiveCaptionGradient2 The ending gradient color of an active caption.
Public Property Get ActiveCaptionGradient2() As Long: End Property

' Returns the text color of an inactive view caption.
'
' @InactiveCaptionForeColor The text color of an inactive caption.
Public Property Get InactiveCaptionForeColor() As Long: End Property

' Returns the beginning gradient color of an inactive view caption.
'
' @InactiveCaptionGradient1 The beginning gradient color of an inactive caption.
Public Property Get InactiveCaptionGradient1() As Long: End Property

' Returns the ending gradient color of an inactive view caption.
'
' @InactiveCaptionGradient2 The ending gradient color of an inactive caption.
Public Property Get InactiveCaptionGradient2() As Long: End Property

' Returns the text color of an active tab.
'
' @FocusTabForeColor The text color of an active tab.
Public Property Get FocusTabForeColor() As Long: End Property

' Returns the beginning gradient color of an active tab.
'
' @FocusTabGradient1 The beginning gradient color of an active tab.
Public Property Get FocusTabGradient1() As Long: End Property

' Returns the ending gradient color of an active tab.
'
' @FocusTabGradient2 The ending gradient color of an active tab.
Public Property Get FocusTabGradient2() As Long: End Property

' Returns the gradient angle (numeric value between 0 and 360) of an active tab.
'
' @FocusTabGradientAngle The gradient angle of an active tab.
Public Property Get FocusTabGradientAngle() As Long: End Property

' Returns the text color of an inactive tab.
'
' @ActiveTabForeColor The text color of an inactive tab.
Public Property Get ActiveTabForeColor() As Long: End Property

' Returns the beginning gradient color of an inactive tab.
'
' @ActiveTabGradient1 The beginning gradient color of an inactive tab.
Public Property Get ActiveTabGradient1() As Long: End Property

' Returns the ending gradient color of an inactive tab.
'
' @ActiveTabGradient2 The beginning gradient color of an inactive tab.
Public Property Get ActiveTabGradient2() As Long: End Property

' Returns the gradient angle (numeric value between 0 and 360) of an inactive tab.
'
' @ActiveTabGradientAngle The gradient angle of an inactive tab.
Public Property Get ActiveTabGradientAngle() As Long: End Property


' Returns the text color of an inactive tab.
'
' @InactiveTabForeColor The text color of an inactive tab.
Public Property Get InactiveTabForeColor() As Long: End Property

' Returns the beginning gradient color of an inactive tab.
'
' @InactiveTabGradient1 The beginning gradient color of an inactive tab.
Public Property Get InactiveTabGradient1() As Long: End Property

' Returns the ending gradient color of an inactive tab.
'
' @InactiveTabGradient2 The beginning gradient color of an inactive tab.
Public Property Get InactiveTabGradient2() As Long: End Property

' Returns the gradient angle (numeric value between 0 and 360) of an inactive tab.
'
' @InactiveTabGradientAngle The gradient angle of an inactive tab.
Public Property Get InactiveTabGradientAngle() As Long: End Property





' You can choose a caption style like gradient caption. Just return a
' constant of the vbCaptionStyle enumeration.
'
' @CaptionStyle Returns the captions style.
Public Property Get CaptionStyle() As vbCaptionStyle: End Property

' Returns true to show view captions, false otherwise. If the captions are
' hidden the drag & drop feature is supported by the tabs. Double click a
' tab to maximize / restore a view.
'
' @ViewCaption Returns true if the views captions should be displayed.
Public Property Get ViewCaptions() As Boolean: End Property

' You can show the views icon on the left side of the view captions if you
' return true.
'
' @ViewCaptionIcons Returns true if the views captions should be displayed.
Public Property Get ViewCaptionIcons() As Boolean: End Property

' The background color of the control.
'
' @BackColor The background color of the control.
Public Property Get BackColor() As Long: End Property

' The background color of the editor area.
'
' @EditorAreaBackColor The background color of the editor area.
Public Property Get EditorAreaBackColor() As Long: End Property

' The frame color of the control.
'
' @FrameColor The frame color of the control.
Public Property Get FrameColor() As Long: End Property

' The width of the frame.
'
' @FrameWidth The frames width.
Public Property Get FrameWidth() As Long: End Property
