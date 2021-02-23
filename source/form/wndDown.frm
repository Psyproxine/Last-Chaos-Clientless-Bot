VERSION 5.00
Begin VB.Form wndDown 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   3270
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pScroller 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   180
      ScaleHeight     =   2655
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   360
      Width           =   2715
      Begin VB.VScrollBar vsb 
         Height          =   1575
         Left            =   2160
         Max             =   115
         SmallChange     =   100
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         ScaleHeight     =   1935
         ScaleWidth      =   1875
         TabIndex        =   1
         Top             =   0
         Width           =   1875
         Begin VB.Timer timUpdate 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   1920
            Top             =   720
         End
         Begin VB.Image ImgItem 
            Height          =   240
            Index           =   0
            Left            =   60
            Picture         =   "wndDown.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   240
         End
         Begin VB.Label lblCaption 
            BackColor       =   &H80000005&
            Caption         =   "Item-0"
            Height          =   255
            Index           =   0
            Left            =   420
            TabIndex        =   2
            Top             =   60
            Visible         =   0   'False
            Width           =   3315
         End
      End
   End
End
Attribute VB_Name = "wndDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''      ControlName:    ISCombo.
''      Filename:       wndDown.frm( Don't modify this form ! !)
''
''      Author:         Alfredo Córdova Pérez ( fred_cpp )
''      e-mail:         fred_cpp@hotmail.com
''                      fred_cpp@yahoo.com.mx
''
''      Description:
''
''      I've Got a lot of problemas with the VB' combo, I couldn't detect
''      when the user changes the selection, and, those combos are relly ugly :(
''      so, I decided made one better.
''      you know, you can use this freely, just give me credit.
''      Votes and suggestions are wellcome.
''



Option Explicit

Dim iPos As Integer
Dim iItems As Integer
Dim IsInside As Boolean
Dim iPrevPos As Integer
Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const WM_SIZE = &H5
Private Const WM_MOVE = &H3
Private Const WM_WINDOWPOSCHANGED = &H47
Private Const WM_KILLFOCUS = &H8
Private Const GWL_WNDPROC = (-4)
Private OriginalWndProc As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



Public m_Items As New Collection
Public m_Images As New Collection
Public m_ShowingList As Boolean
Public ItemClick As Integer

Event ItemClick(iItem As Integer, sText As String)

Dim nValue As Long

'' Detect if the Mouse cursor is inside a Window
Private Function InBox(ObjectHWnd As Long) As Boolean
    Dim mpos As PointAPI
    Dim oRect As RECT
    GetCursorPos mpos
    GetWindowRect ObjectHWnd, oRect
    If mpos.X >= oRect.Left And mpos.X <= oRect.Right And _
        mpos.Y >= oRect.Top And mpos.Y <= oRect.Bottom Then
        InBox = True
    Else
        InBox = False
   End If
End Function

Private Sub Form_Paint()
    '' Draw The Border of the Window
    ScaleMode = 3
    Line (0, 0)-(ScaleWidth - 1, 0), vb3DHighlight
    Line (0, 0)-(0, ScaleHeight - 1), vb3DHighlight
    Line (ScaleWidth - 1, 0)-(ScaleWidth - 1, ScaleHeight - 1), vb3DShadow
    Line (0, ScaleHeight - 1)-(ScaleWidth - 1, ScaleHeight - 1), vb3DShadow
End Sub

'' Draw All items
Private Sub DrawAll(ActiveItem As Integer)
    lblCaption(iPrevPos).BackColor = vbWindowBackground
    lblCaption(iPrevPos).ForeColor = vbButtonText
    lblCaption(ActiveItem).BackColor = vbHighlight
    lblCaption(ActiveItem).ForeColor = vbHighlightText
    If ActiveItem <= 0 Then
        iPrevPos = 0
    Else
        iPrevPos = ActiveItem
    End If
    Debug.Print "DrawAll"
End Sub

'' Raise the ItemClick event
Private Sub lblCaption_Click(Index As Integer)
    Reset
    RaiseEvent ItemClick(Index, lblCaption(Index).Caption)
End Sub

'' Detect the mouse movement
Private Sub lblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        timUpdate.Enabled = True
    End If
    iPos = Index
End Sub

''  Hide and unload if the window lost the focus
Private Sub picGroup_LostFocus()
    Reset
End Sub

''  Hide and unload if the window lost the focus
Private Sub Form_LostFocus()
    Reset
End Sub

'' Activate the TipUpdate Timer
Private Sub picGroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then
        timUpdate.Enabled = True
    End If
End Sub

''  Hide and unload if the window lost the focus
Private Sub pScroller_LostFocus()
    'If vsb.Visible Then Exit Sub
    Reset
End Sub

'' Detect the position of the cursor
''  (Only if the cursor is in the Window)
Private Sub timUpdate_Timer()
Static temiPOs As Integer
    If InBox(picGroup.hwnd) Then
        If IsInside Then
            If temiPOs <> iPos Then
                DrawAll iPos
            End If
        Else
            IsInside = True
        End If
    Else
        timUpdate.Enabled = False
        DrawAll 0
        IsInside = False
    End If
    temiPOs = iPos
End Sub

'' Change the position of the items when the ScrollBar changes
Private Sub vsb_Change()
    On Error Resume Next
    picGroup.Move 1, 1 - 255 * vsb.Value
    Me.SetFocus
End Sub

'' Hide Window and Save state in Variable
Private Sub Reset()
    Hide
    m_ShowingList = False
End Sub

'' This function Show the cDown Window, And adds the items
Public Function PopUp(X As Long, Y As Long, lWidth As Single, parent As Object) As Boolean
    Dim ni As Integer
    Dim ht As Single
    Dim lHeight As Single
    m_ShowingList = True
    ht = 255 * (m_Items.Count) + 60
    For ni = 1 To m_Items.Count + 2
        Load lblCaption(ni)
        Load imgItem(ni)
    Next ni
    On Error GoTo LimitOfItems
    For ni = 1 To m_Items.Count
        lblCaption(ni - 1).Visible = True
        lblCaption(ni - 1).Caption = m_Items.Item(ni)
        lblCaption(ni - 1).Move 360, 255 * (ni - 1)
        imgItem(ni - 1).Visible = True
        Set imgItem(ni - 1).Picture = m_Images(ni)
        imgItem(ni - 1).Move 30, 255 * (ni - 1)
    Next ni
LimitOfItems:
    If m_Items.Count <= 8 Then
        lHeight = ht
        vsb.Visible = False
    Else
        lHeight = 8 * 255 + 60
        vsb.Visible = True
        vsb.Min = 1
        vsb.Max = m_Items.Count - 8
    End If
    Visible = True
    Move X, Y, lWidth, lHeight
    Show ', 'parent
    picGroup.Move 15, 15, Width, ht
    pScroller.Move 30, 30, Width - 60, lHeight - 60
    vsb.Move Width - vsb.Width - 2 * Screen.TwipsPerPixelX, 0, vsb.Width, lHeight - 30
    pScroller.Move 15, 15, Width - 30, lHeight - 30
    iPrevPos = 0
'    If vsb.Visible Then
'        'vsb.SetFocus
'        pScroller.SetFocus
'    Else
'        picGroup.SetFocus
'    End If
    Me.SetFocus
End Function

