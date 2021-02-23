VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SplitBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' <B>Description</B><BR />
' <BR />
' A splitter class uses the Desktop window to draw a
' splitter bar, therefore allowing splitting of MDI forms
' as well as standard forms.<BR />
' <BR />
' <B>Remarks</B><BR />
' <BR />
' This code was extract from vbaccelarator web site and modified
' to acomplish the requirements of this project
'
' Thanks to the author SP McMahon!

Option Explicit

Private Const SPLITBAR_SIZE As Long = 2

Private bDraw As Boolean
Private rcCurrent As RECT
Private rcNew As RECT
Private rcWindow As RECT
Private m_Offset As Long

Private Const VK_LBUTTON = &H1

Public Enum eOrientationConstants
   espVertical = 1
   espHorizontal = 2
End Enum

' A simple point with x and y coordinates.
Private Type PointAPI
   x As Long ' The x coordinate of the point.
   y As Long ' The y coordinate of the point.
End Type

'Private Type LOGBRUSH
'   lbStyle As Long
'   lbColor As Long
'   lbHatch As Long
'End Type

Private Declare Function SelectObject Lib "gdi32" ( _
   ByVal hDc As Long, _
   ByVal hObject As Long _
) As Long

Private Declare Function CreatePen Lib "gdi32" ( _
   ByVal nPenStyle As Long, _
   ByVal nWidth As Long, _
   ByVal crColor As Long _
) As Long

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
   ByVal hWnd As Long, _
   ByVal lpClassName As String, _
   ByVal nMaxCount As Long _
) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function Rectangle Lib "gdi32" ( _
   ByVal hDc As Long, _
   ByVal X1 As Long, _
   ByVal Y1 As Long, _
   ByVal X2 As Long, _
   ByVal Y2 As Long _
) As Long

Private Declare Function GetKeyState Lib "user32" ( _
   ByVal nVirtKey As Long _
) As Integer

Private Declare Function SetCapture Lib "user32" ( _
   ByVal hWnd As Long _
) As Long

Private Declare Function GetCursorPos Lib "user32" ( _
   lpPoint As PointAPI _
) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" _
   Alias "CreateDCA" ( _
   ByVal lpDriverName As String, _
   lpDeviceName As Any, _
   lpOutput As Any, _
   lpInitData As Any _
) As Long

Private Declare Function DeleteDC Lib "gdi32" ( _
   ByVal hDc As Long _
) As Long

Private Declare Function SetROP2 Lib "gdi32" ( _
   ByVal hDc As Long, _
   ByVal nDrawMode As Long _
) As Long

Private m_hWnd As Long
Private m_eOrientation As eOrientationConstants
Private m_bSplitting As Boolean

Public Event AfterResize(ByVal NewSize As Long)

' Initialization of the class.
Private Sub Class_Initialize()
   m_eOrientation = espVertical
End Sub

Public Property Get Offset() As Long
    Offset = m_Offset
End Property

Public Property Get Orientation() As eOrientationConstants
    Orientation = m_eOrientation
End Property

Public Property Let Orientation(ByVal eOrientation As eOrientationConstants)
    m_eOrientation = eOrientation
End Property

' Start dragging the splitter bar.
'
' @hWnd The window handle.
' @Rc The rectangle which bounds the maximal splitting area.
' @x The x position of the mouse cursor.
' @y The y position of the mouse cursor.
Friend Sub SplitterMouseDown(ByVal hWnd As Long, Rc As RECT, ByVal x As Long, ByVal y As Long)
    
    Dim tP As PointAPI
    Dim tpPrev As PointAPI
    Dim hDc As Long
    Dim hPen As Long
    
    m_hWnd = hWnd
    ' Send subsequent mouse messages to the owner window
    SetCapture m_hWnd
    ' get window rect
    rcWindow = Rc

    GetCursorPos tP
    ' Store the initial cursor position
    tpPrev.x = tP.x
    tpPrev.y = tP.y
    If (m_eOrientation = espHorizontal) Then
        m_Offset = tP.y
    Else
        m_Offset = tP.x
    End If
    bDraw = True  ' start actual drawing from next move message
    rcCurrent.Left = 0: rcCurrent.Top = 0: rcCurrent.Right = 0: rcCurrent.Bottom = 0
    SplitterFormMouseMove tP.x, tP.y
    Do While GetKeyState(VK_LBUTTON) < 0
        GetCursorPos tP
        If tP.x <> tpPrev.x Or tP.y <> tpPrev.y Then
            tpPrev.x = tP.x
            tpPrev.y = tP.y
            SplitterFormMouseMove tP.x, tP.y
        End If
        DoEvents
    Loop
    SplitterFormMouseUp tP.x, tP.y
End Sub

' Redraw the splitter while dragging.
'
' @x The x position of the mouse cursor.
' @y The y position of the mouse cursor.
Private Sub SplitterFormMouseMove(ByVal x As Long, ByVal y As Long)
    Dim hDc As Long
    Dim Pt As PointAPI
    
    If (bDraw) Then
        
        DrawDragRect rcCurrent, 1
        ' It is simpler to use the mouse cursor position than try to translate
        ' X,Y to screen coordinates!
        GetCursorPos Pt
        ' Determine where to draw the splitter:
        If (m_eOrientation = espHorizontal) Then
            rcNew.Left = rcWindow.Left
            rcNew.Right = rcWindow.Right
            If (Pt.y >= rcWindow.Top) And (Pt.y < rcWindow.Bottom) Then
                rcNew.Top = Pt.y - SPLITBAR_SIZE
                rcNew.Bottom = Pt.y + SPLITBAR_SIZE
            Else
                If (Pt.y < rcWindow.Top) Then
                    rcNew.Top = rcWindow.Top - SPLITBAR_SIZE
                    rcNew.Bottom = rcNew.Top + SPLITBAR_SIZE
                Else
                    rcNew.Top = rcWindow.Bottom - SPLITBAR_SIZE
                    rcNew.Bottom = rcNew.Top + SPLITBAR_SIZE
                End If
            End If
        Else
            rcNew.Top = rcWindow.Top
            rcNew.Bottom = rcWindow.Bottom
            If (Pt.x >= rcWindow.Left) And (Pt.x <= rcWindow.Right) Then
                rcNew.Left = Pt.x - SPLITBAR_SIZE
                rcNew.Right = Pt.x + SPLITBAR_SIZE
            Else
                If (Pt.x < rcWindow.Left) Then
                    rcNew.Left = rcWindow.Left - SPLITBAR_SIZE
                    rcNew.Right = rcNew.Left + SPLITBAR_SIZE

                End If
            End If
        End If
        ' Draw the new rectangle
        'Rectangle hDc, rcNew.Left, rcNew.Top, rcNew.Right, rcNew.Bottom
        DrawDragRect rcNew, 1
        ' Store this position so we can erase it next time:
        LSet rcCurrent = rcNew
        ' Free the reference to the Desktop DC we got (make sure you do this!)
        'DeleteDC hDc
        'DeleteObject hPen
    End If
End Sub

' Undraw the splitter user drops it by releasing the mouse button.
'
' @x The x position of the mouse cursor.
' @y The y position of the mouse cursor.
Private Function SplitterFormMouseUp(ByVal x As Long, ByVal y As Long) As Boolean
    Dim hDc As Long
    Dim tP As PointAPI
    Dim hWndClient As Long

    ' Release mouse capture:
    ReleaseCapture
    ' Don't leave orphaned rectangle on desktop; erase last rectangle.
    If (bDraw) Then
        bDraw = False
        ' Get the Desktop DC:
        'hDc = CreateDCAsNull("DISPLAY", 0, 0, 0)
        ' Set to XOR drawing mode:
        'SetROP2 hDc, R2_XORPEN
        ' Erase the last rectangle:
        'Rectangle hDc, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom
        DrawDragRect rcCurrent, 1
        ' Clear up the desktop DC:
        'DeleteDC hDc
        ' Here we ensure the splitter is within bounds before releasing:
        GetCursorPos tP
        If (tP.x < rcWindow.Left) Then
            tP.x = rcWindow.Left
        End If
        If (tP.x > rcWindow.Right) Then
            tP.x = rcWindow.Right
        End If
        If (tP.y < rcWindow.Top) Then
            tP.y = rcWindow.Top
        End If
        If (tP.y > rcWindow.Bottom) Then
            tP.y = rcWindow.Bottom
        End If
        ' Move the splitter to the validated final position:
        If (m_eOrientation = espHorizontal) Then
            m_Offset = (tP.y - m_Offset) * Screen.TwipsPerPixelY
        Else
            m_Offset = (tP.x - m_Offset) * Screen.TwipsPerPixelX
        End If
        ' Return true to tell the owner we have completed splitting:
        SplitterFormMouseUp = True
    End If
    
    If Orientation = espVertical Then
       RaiseEvent AfterResize(rcNew.Left)
    Else
       RaiseEvent AfterResize(rcNew.Top)
    End If
    
End Function

Private Function ClassName(ByVal lhWnd As Long) As String
    Dim lLen As Long
    Dim sBuf As String
    
    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lhWnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left$(sBuf, lLen)
    End If
End Function

