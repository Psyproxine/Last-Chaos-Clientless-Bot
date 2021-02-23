VERSION 5.00
Begin VB.UserControl ISCombo 
   BackColor       =   &H80000005&
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3660
   LockControls    =   -1  'True
   ScaleHeight     =   1995
   ScaleWidth      =   3660
   ToolboxBitmap   =   "ISCombo.ctx":0000
   Begin VB.PictureBox picButton 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1800
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2640
      Width           =   435
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo"
      Top             =   2760
      Width           =   1875
   End
   Begin VB.Timer timUpdate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1560
      Top             =   180
   End
   Begin VB.Image imgItem 
      Height          =   195
      Left            =   240
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   375
   End
   Begin VB.Image imgDown 
      Height          =   480
      Left            =   1140
      Picture         =   "ISCombo.ctx":0312
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "ISCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''      ControlName:    ISCombo.
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

' Type Declarations
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

Private Enum State
    Normal
    Hover
    pushed
End Enum

Private InOut As Boolean
Private iState As State
Private OnClicking As Boolean
Private OnFocus As Boolean

Private gScaleX As Single '= Screen.TwipsPerPixelX
Private gScaleY As Single '= Screen.TwipsPerPixelY

Private WithEvents cDown As wndDown
Attribute cDown.VB_VarHelpID = -1

'Default Property Values:
Const m_def_Enabled = True
Const m_def_FontColor = 0
Const m_def_FontHighlightColor = 0
Const m_def_IconAlign = 0
Const m_def_IconSize = 0
Const m_def_TextAlign = 4
Const m_def_BackColor = &HE0E0E0
Const m_def_HoverColor = &HFFF0B8
Const m_def_Default = False
'Property Variables:
Dim m_Enabled As Boolean
Dim m_FontColor As OLE_COLOR
Dim m_FontHighlightColor As OLE_COLOR
Dim m_IconSize As Integer
Dim m_TextAlign As Integer
Dim m_IconAlign As Integer
Dim m_Icon As Picture
Dim m_HoverIcon As Picture
Dim m_BackColor As OLE_COLOR
Dim m_HoverColor As OLE_COLOR
Dim m_Default As Boolean
Dim m_Focused As Boolean
Dim m_ImageSize As Integer
Dim m_Items As New Collection
Dim m_Images As New Collection
Dim m_ItemsCount As Integer

'Event Declarations:

Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseOut()
Event MouseHover()
Event KeyPress(KeyAscii As Integer)
Event ButtonClick()
Event ItemClick(iItem As Integer)
Event Change()
Const pBorderColor = &HC08080
' API Declarations

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' Support Functions
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

Private Sub DrawFlat()
    If OnFocus Then
        DrawFace 4
    Else
        DrawFace 0
    End If
End Sub

Private Sub DrawRaised()
    DrawFace 3
End Sub

Private Sub DrawPushed()
    DrawFace 1
End Sub


Private Sub cDown_ItemClick(iItem As Integer, sText As String)
    UserControl.imgItem.Picture = m_Images(iItem + 1)
    txtText.Text = sText
    txtText.SelStart = 0
    txtText.SelLength = Len(sText)
    RaiseEvent ItemClick(iItem)
End Sub

Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OnClicking = True
    If Button = vbLeftButton Then
        picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(picButton.ScaleWidth - 1, 0), vbWindowBackground
        picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(0, picButton.ScaleHeight - 1), vbWindowBackground
        picButton.Line (0, 0)-(0, picButton.ScaleHeight - 1), vb3DShadow
        picButton.Line (1, 0)-(picButton.ScaleWidth - 1, 0), vb3DShadow
    End If
End Sub

Private Sub picButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not OnClicking Then
        UserControl_MouseMove Button, Shift, X, Y
    End If
End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OnClicking = False
End Sub

Private Sub picButton_Paint()
    picButton.PaintPicture imgDown.Picture, (picButton.ScaleWidth - imgDown.Width) / 2, (picButton.ScaleHeight - imgDown.Height) / 2
End Sub

Private Sub timUpdate_Timer()
    If InBox(UserControl.hwnd) Then
        If InOut = False Then
            iState = Hover
            DrawRaised
            RaiseEvent MouseHover
        End If
        InOut = True
    Else
        If InOut Then
            timUpdate.Enabled = False
            iState = Normal
            DrawFlat
            RaiseEvent MouseOut
        End If
        InOut = False
    End If
End Sub

Private Sub txtText_Change()
    RaiseEvent Change
End Sub

Private Sub txtText_GotFocus()
    txtText.SelStart = 0
    txtText.SelLength = Len(txtText.Text)
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyUp, vbKeyLeft
        Case vbKeyDown, vbKeyRight
    End Select
End Sub

Private Sub txtText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_Click()
    'If m_Enabled Then RaiseEvent Click
End Sub

Private Sub UserControl_EnterFocus()
    OnFocus = True
    If m_Enabled Then
        If Not InBox(UserControl.hwnd) Then
            DrawFace 3
        End If
    End If
End Sub

Private Sub UserControl_ExitFocus()
    OnFocus = False
    If m_Enabled Then
        If Not InBox(UserControl.hwnd) Then
            If Not OnFocus Then
                DrawFace 0
            Else
                DrawFace 4
            End If
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    gScaleX = Screen.TwipsPerPixelX
    gScaleY = Screen.TwipsPerPixelY
    m_ImageSize = 16 * gScaleX
    'Set ImgIcon.Picture = LoadPicture()
    'UserControl.ImgIcon = m_Icon
    UserControl_Resize
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If m_Enabled Then DrawFace 1
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If m_Enabled Then
        RaiseEvent Click
        RaiseEvent KeyPress(KeyAscii)
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If m_Enabled Then DrawFace 0
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        iState = pushed
        UserControl_Paint
        timUpdate.Enabled = False
        OnClicking = True
        RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        If Button = 0 Then timUpdate.Enabled = True
        RaiseEvent MouseMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        iState = Hover
        UserControl_Paint
        timUpdate.Enabled = True
        RaiseEvent MouseUp(Button, Shift, X, Y)
        If InBox(UserControl.hwnd) Then  'If OnClicking Then
            RaiseEvent Click
        End If
        OnClicking = False
    End If
End Sub

Private Sub UserControl_Resize()
    '   Text Position
    UserControl.ScaleMode = 1
    imgItem.Move 75, (UserControl.Height - m_ImageSize) / 2, m_ImageSize, m_ImageSize
    txtText.Move 105 + m_ImageSize, (UserControl.Height - m_ImageSize) / 2, Width - m_ImageSize - 150
    Select Case m_TextAlign
        Case 0  '   Left
            txtText.Alignment = 0
        Case 1  '   Right
            txtText.Alignment = 1
        Case 2  '   Top
            txtText.Alignment = 2
    End Select
    'Locate Button
    picButton.Move Width - 270, 30, 240, Height - 60
    'imgDown.Move (picButton.ScaleWidth - imgDown.Width) / 2, (picButton.ScaleHeight - imgDown.Height) / 2
End Sub

Private Sub UserControl_Paint()
    '
    If m_Enabled Then
        Select Case iState
            Case Hover
                DrawFace 4
            Case pushed
                DrawFace 1
            Case Normal
                DrawFace 0
        End Select
    Else
        DrawFace 2
    End If

End Sub

Private Sub UserControl_DblClick()
    If m_Enabled Then RaiseEvent DblClick
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txttext,txttext,-1,Caption
Public Property Get Caption() As String
    Caption = txtText.Text
End Property

Public Property Let Caption(ByVal New_Caption As String)
    txtText.Text() = New_Caption
    PropertyChanged "Caption"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txttext,txttext,-1,Font
Public Property Get Font() As Font
    Set Font = txtText.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtText.Font = New_Font
    PropertyChanged "Font"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=picCmd,picCmd,-1,ToolTipText
Public Property Get ToolTipText() As String
    ToolTipText = txtText.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    txtText.ToolTipText() = New_ToolTipText
    'ImgIcon.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    'Global Constants Initialization.
    Set m_Icon = LoadPicture("")
    m_TextAlign = m_def_TextAlign
    m_FontColor = m_def_FontColor
    m_FontHighlightColor = m_def_FontHighlightColor
    m_Default = m_def_Default
    m_Enabled = m_def_Enabled
    txtText.Text = UserControl.Extender.Name
    
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim picNormal As Picture
    With PropBag
        Set picNormal = PropBag.ReadProperty("Icon", Nothing)
        If Not (picNormal Is Nothing) Then Set Icon = picNormal
    End With

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    txtText.Text = PropBag.ReadProperty("Caption", "Caption")
    Set txtText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtText.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_HoverColor = PropBag.ReadProperty("HoverColor", m_def_HoverColor)
    m_TextAlign = PropBag.ReadProperty("TextAlign", m_def_TextAlign)
    m_IconAlign = PropBag.ReadProperty("IconAlign", m_def_IconAlign)
    m_IconSize = PropBag.ReadProperty("IconSize", m_def_IconSize)
    m_FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
    m_FontHighlightColor = PropBag.ReadProperty("FontHighlightColor", m_def_FontHighlightColor)
    m_Default = PropBag.ReadProperty("Default", False)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
End Sub

Private Sub UserControl_Show()
    UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    If Not cDown Is Nothing Then
        Unload cDown
        Set cDown = Nothing
    End If
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", txtText.Text, "Caption")
    Call PropBag.WriteProperty("Font", txtText.Font, Ambient.Font)
    Call PropBag.WriteProperty("ToolTipText", txtText.ToolTipText, "")
    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("TextAlign", m_TextAlign, m_def_TextAlign)
    Call PropBag.WriteProperty("FontColor", m_FontColor, m_def_FontColor)
    Call PropBag.WriteProperty("FontHighlightColor", m_FontHighlightColor, m_def_FontHighlightColor)
    Call PropBag.WriteProperty("Default", m_Default, m_def_Default)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=11,0,0,0
Public Property Get Icon() As Picture
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
'    Set m_Icon = New_Icon
'    Set ImgIcon.Picture = New_Icon
    PropertyChanged "Icon"
End Property


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,4
Public Property Get TextAlign() As Integer
    TextAlign = m_TextAlign
End Property

Public Property Let TextAlign(ByVal New_TextAlign As Integer)
    m_TextAlign = New_TextAlign
    UserControl_Resize
    PropertyChanged "TextAlign"
End Property


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=10,0,0,0
Public Property Get FontColor() As OLE_COLOR
    FontColor = m_FontColor
End Property

Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
    m_FontColor = New_FontColor
    txtText.ForeColor = New_FontColor
    PropertyChanged "FontColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=10,0,0,0
Public Property Get FontHighlightColor() As OLE_COLOR
    FontHighlightColor = m_FontHighlightColor
End Property

Public Property Let FontHighlightColor(ByVal New_FontHighlightColor As OLE_COLOR)
    m_FontHighlightColor = New_FontHighlightColor
    PropertyChanged "FontHighlightColor"
End Property


Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    UserControl.Enabled = New_Enabled
    If New_Enabled Then
        txtText.BackColor = vbWindowBackground
        UserControl.BackColor = vbWindowBackground
        txtText.ForeColor = vbButtonText
        txtText.Locked = False
    Else
        txtText.BackColor = vb3DFace
        UserControl.BackColor = vb3DFace
        txtText.ForeColor = vbGrayText
        txtText.Locked = True
    End If
    UserControl_Paint
    PropertyChanged "Enabled"
End Property


Private Sub picButton_Click()
    ' Show de Auxiliar Window
    'On Error GoTo NoItemsToShow
    Dim ni As Integer

    If m_Enabled Then
        Set cDown = New wndDown
        For ni = 1 To m_Items.Count
            cDown.m_Items.Add m_Items(ni)
            cDown.m_Images.Add m_Images(ni)
        Next ni
        RaiseEvent ButtonClick
        Dim rt As RECT
        GetWindowRect UserControl.hwnd, rt
        'cDown.Visible = True
        'cDown.Move rt.Left * gScaleX, rt.Bottom * gScaleY
        'cDown.Show , UserControl.Extender.parent
        cDown.PopUp rt.Left * gScaleX, rt.Bottom * gScaleY, UserControl.Width, UserControl.Extender.parent
    End If
NoItemsToShow:
End Sub

Private Sub DrawFace(iState As Integer)
    '' This is the drawing code, I know there are better ways to do this,
    '' but I writte this 6 months and, and I don't want to work on this :)
    UserControl.ScaleMode = 3
    Select Case iState
        Case 0: 'Normal
            UserControl.Cls
            UserControl.DrawWidth = 2
            UserControl.ForeColor = vb3DFace
            UserControl.Line (1, 1)-(ScaleWidth + 1, 1)
            UserControl.Line (1, 1)-(1, ScaleHeight + 1)
            UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(ScaleWidth - 1, -1)
            UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(-1, ScaleHeight - 1)
            picButton.Line (0, 0)-(picButton.ScaleWidth - 1, picButton.ScaleHeight - 1), vbWindowBackground, B
        Case 1: 'Pushed
            UserControl.ForeColor = vb3DLight
            UserControl.Line (ScaleWidth - 2, 2)-(ScaleWidth - 2, ScaleHeight - 2)
            UserControl.Line (ScaleWidth - 2, ScaleHeight - 2)-(2, ScaleHeight - 2)
            UserControl.Line (1, 1)-(ScaleWidth + 1, 1)
            UserControl.Line (1, 1)-(1, ScaleHeight + 1)
            UserControl.ForeColor = vb3DShadow
            UserControl.Line (0, 0)-(ScaleWidth, 0)
            UserControl.Line (0, 0)-(0, ScaleHeight)
            UserControl.ForeColor = vbWindowBackground
            UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(ScaleWidth - 1, -1)
            UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(-1, ScaleHeight - 1)
            picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(picButton.ScaleWidth - 1, 0), vb3DShadow
            picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(0, picButton.ScaleHeight - 1), vb3DShadow
            picButton.Line (0, 0)-(0, picButton.ScaleHeight - 1), vb3DFace
            picButton.Line (1, 0)-(1, picButton.ScaleHeight - 1), vbWindowBackground
        
        Case 2: 'Disabled
            picButton.Cls
            picButton_Paint
            txtText.BackColor = vb3DFace
            UserControl.BackColor = vb3DFace
            UserControl.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), vbWindowBackground, B
        Case 3: 'Highlight
        UserControl.DrawWidth = 1
            UserControl.ForeColor = vb3DLight
            UserControl.Line (ScaleWidth - 2, 2)-(ScaleWidth - 2, ScaleHeight - 2)
            UserControl.Line (ScaleWidth - 2, ScaleHeight - 2)-(2, ScaleHeight - 2)
            UserControl.Line (1, 1)-(ScaleWidth + 1, 1)
            UserControl.Line (1, 1)-(1, ScaleHeight + 1)
            UserControl.ForeColor = vb3DShadow
            UserControl.Line (0, 0)-(ScaleWidth, 0)
            UserControl.Line (0, 0)-(0, ScaleHeight)
            UserControl.ForeColor = vbWindowBackground
            UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(ScaleWidth - 1, -1)
            UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(-1, ScaleHeight - 1)
            picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(picButton.ScaleWidth - 1, 0), vb3DShadow
            picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(0, picButton.ScaleHeight - 1), vb3DShadow
            picButton.Line (0, 0)-(0, picButton.ScaleHeight - 1), vb3DFace
            picButton.Line (1, 0)-(1, picButton.ScaleHeight - 1), vbWindowBackground
        Case 4: 'Focused
            UserControl.ForeColor = vb3DLight
            UserControl.Line (ScaleWidth - 2, 2)-(ScaleWidth - 2, ScaleHeight - 2)
            UserControl.Line (ScaleWidth - 2, ScaleHeight - 2)-(2, ScaleHeight - 2)
            UserControl.Line (1, 1)-(ScaleWidth + 1, 1)
            UserControl.Line (1, 1)-(1, ScaleHeight + 1)
            UserControl.ForeColor = vb3DShadow
            UserControl.Line (0, 0)-(ScaleWidth, 0)
            UserControl.Line (0, 0)-(0, ScaleHeight)
            UserControl.ForeColor = vbWindowBackground
            UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(ScaleWidth - 1, -1)
            UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(-1, ScaleHeight - 1)
            picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(picButton.ScaleWidth - 1, 0), vb3DShadow
            picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(0, picButton.ScaleHeight - 1), vb3DShadow
            picButton.Line (0, 0)-(0, picButton.ScaleHeight - 1), vb3DFace
            picButton.Line (1, 0)-(1, picButton.ScaleHeight - 1), vbWindowBackground
    End Select
End Sub

'' Add a new Item to the Combo List
Public Sub AddItem(Text As String, Optional Index As Integer, Optional iImage As Picture)
    Dim ImageTemp As Picture
    If IsMissing(iImage) Then
        Set ImageTemp = LoadPicture()
    Else
        Set ImageTemp = iImage
    End If
    m_Items.Add Text
    m_Images.Add ImageTemp
End Sub

