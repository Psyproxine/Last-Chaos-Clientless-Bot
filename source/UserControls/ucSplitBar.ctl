VERSION 5.00
Begin VB.UserControl ucSplitBar 
   BackColor       =   &H00F5F9FA&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "ucSplitBar.ctx":0000
End
Attribute VB_Name = "ucSplitBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Rect As RECT
Private m_RectLeft As Long
Private m_RectRight As Long
Private m_RectTop As Long

Private m_TopPos As Long ' The top position
Private m_BottomPos As Long ' The bottom position
Private m_LeftPos As Long ' The left position
Private m_RightPos As Long ' The right position

Private m_RectBottom As Long
Private m_Folder As Folder
Private m_Orientation As eOrientationConstants

Private WithEvents SplitBar As SplitBar
Attribute SplitBar.VB_VarHelpID = -1

Private Sub UserControl_Initialize()
   Set SplitBar = New SplitBar
   Orientation = Orientation
   BackColor = vbButtonFace
End Sub

Private Sub UserControl_Terminate()
   Set SplitBar = Nothing
End Sub

Public Property Get Folder() As Folder
   Set Folder = m_Folder
End Property
Public Property Set Folder(ByRef NewFolder As Folder)
   Set m_Folder = NewFolder
End Property

Public Property Let TopPos(ByVal NewTopPos As Long)
   m_TopPos = NewTopPos
End Property

Public Property Get TopPos() As Long
   TopPos = m_TopPos
End Property

Public Property Let BottomPos(ByVal NewBottomPos As Long)
   m_BottomPos = NewBottomPos
End Property

Public Property Get BottomPos() As Long
   BottomPos = m_BottomPos
End Property

Public Property Let LeftPos(ByVal NewLeftPos As Long)
   m_LeftPos = NewLeftPos
End Property

Public Property Get LeftPos() As Long
   LeftPos = m_LeftPos
End Property

Public Property Let RightPos(ByVal NewRightPos As Long)
   m_RightPos = NewRightPos
End Property

Public Property Get RightPos() As Long
   RightPos = m_RightPos
End Property

Public Property Get RectLeft() As Long
   RectLeft = m_RectLeft
End Property
Public Property Let RectLeft(ByVal NewRectLeft As Long)
   m_RectLeft = NewRectLeft
End Property

Public Property Get RectRight() As Long
   RectRight = m_RectRight
End Property
Public Property Let RectRight(ByVal NewRectRight As Long)
   m_RectRight = NewRectRight
End Property

Public Property Get RectTop() As Long
   RectTop = m_RectTop
End Property
Public Property Let RectTop(ByVal NewRectTop As Long)
   m_RectTop = NewRectTop
End Property

Public Property Get RectBottom() As Long
   RectBottom = m_RectBottom
End Property
Public Property Let RectBottom(ByVal NewRectBottom As Long)
   m_RectBottom = NewRectBottom
End Property

Public Property Get Orientation() As eOrientationConstants
   Orientation = m_Orientation
End Property
Public Property Let Orientation(ByVal NewOrientation As eOrientationConstants)
   
   m_Orientation = NewOrientation
   
   UserControl.MousePointer = 99
   If NewOrientation = espHorizontal Then
      UserControl.MouseIcon = getResourceCursor(CURSOR_ARROW_HORIZONTAL_SPLITTER)
   Else
      UserControl.MouseIcon = getResourceCursor(CURSOR_ARROW_VERTICAL_SPLITTER)
   End If
   
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   On Error Resume Next
   
   Parent.Refresh
   
   With m_Rect
      .Left = m_RectLeft
      .Top = m_RectTop
      .Bottom = m_RectBottom
      .Right = m_RectRight
   End With
   
   With SplitBar
      .Orientation = m_Orientation
      .SplitterMouseDown UserControl.hWnd, m_Rect, x, y
   End With
  
End Sub

Private Sub SplitBar_AfterResize(ByVal NewSize As Long)
   
   Dim NewRatio As Double

   ' ----------------------------------------------------------------------
   ' Calculate new ratio
   ' ----------------------------------------------------------------------
   Select Case Folder.Relationship
      Case vbRelLeft:     NewSize = NewSize + 4
                         NewRatio = (NewSize - m_Rect.Left) / (m_Rect.Right - m_Rect.Left)
      Case vbRelRight:    NewRatio = 1 - ((NewSize - m_Rect.Left) / (m_Rect.Right - m_Rect.Left))
      Case vbRelTop:      NewSize = NewSize - 4
                         NewRatio = (NewSize - m_Rect.Top) / (m_Rect.Bottom - m_Rect.Top)
      Case vbRelBottom:   NewRatio = (NewSize - m_Rect.Top) / (m_Rect.Bottom - m_Rect.Top)
   End Select

   If NewRatio < 0.015 Then NewRatio = 0.05 ' Minimal size
   If NewRatio > 0.985 Then NewRatio = 0.95 ' Maximal size
   
   ' ----------------------------------------------------------------------
   ' Set new ratio and refresh perspective
   ' ----------------------------------------------------------------------
   Folder.Ratio = NewRatio
   
   Parent.Refresh
      
End Sub

Public Sub Refresh()
   UserControl.BackColor = m_Scheme.BackColor
End Sub
