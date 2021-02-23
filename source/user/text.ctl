VERSION 5.00
Begin VB.UserControl text 
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   ScaleHeight     =   1095
   ScaleWidth      =   3495
   Begin VB.TextBox textbox1 
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   3240
   End
   Begin VB.Shape cadre 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00B1AA54&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINT_API) As Long

Public Enum State_b
    Normal_ = 0
    Default_ = 1
End Enum

Dim m_State As State_b
Dim m_Font As Font

Const m_Def_State = State_b.Normal_

Private Type POINT_API
    X As Long
    Y As Long
End Type
Const m_def_PasswordChar = ""
Dim m_PasswordChar As String

Dim s As Integer
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = textbox1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    textbox1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = textbox1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    textbox1.BackColor() = New_BackColor
    cadre.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
Private Sub textbox1_Change()
RaiseEvent Change
End Sub
Public Property Get Locked() As Boolean
    Locked = textbox1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    textbox1.Locked() = New_Locked
    PropertyChanged "Locked"
End Property
Public Property Get MaxLength() As Long
    MaxLength = textbox1.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    textbox1.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Public Property Get SelStart() As Long
    SelStart = textbox1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    textbox1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property
Public Property Get SelText() As String
    SelText = textbox1.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    textbox1.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

Public Property Get SelLength() As Long
    SelLength = textbox1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    textbox1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property
Public Property Get PasswordChar() As String
On Error Resume Next
    textbox1.PasswordChar = m_PasswordChar
End Property
Public Property Let PasswordChar(ByVal New_PasswordChar As String)
On Error Resume Next
    textbox1.PasswordChar = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property
Public Property Get FontBold() As Boolean
    FontBold = textbox1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    textbox1.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property
Public Property Get FontItalic() As Boolean
    FontItalic = textbox1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    textbox1.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property
Public Property Get Font() As Font
    Set Font = textbox1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set textbox1.Font = New_Font
    PropertyChanged "Font"
End Property

'SHAPE
Public Property Get LineColor() As OLE_COLOR
    LineColor = cadre.BorderColor
End Property

Public Property Let LineColor(ByVal New_LineColor As OLE_COLOR)
    cadre.BorderColor() = New_LineColor
    PropertyChanged "LineColor"
End Property
'*******

Public Property Get text() As String
    text = textbox1.text
End Property

Public Property Let text(ByVal New_Text As String)
    textbox1.text() = New_Text
    PropertyChanged "Text"
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Sub Refresh()
    UserControl.Refresh
End Sub
Private Sub UserControl_Resize()
textbox1.Left = 40
textbox1.Top = 40
textbox1.Width = UserControl.ScaleWidth - 90
textbox1.Height = UserControl.ScaleHeight - 90
cadre.Left = 0
cadre.Top = 0
cadre.Width = UserControl.ScaleWidth
cadre.Height = UserControl.ScaleHeight
If UserControl.Height < 180 Then UserControl.Height = 290

End Sub
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_PasswordChar = m_def_PasswordChar
End Sub

Public Property Get Alignment() As Integer
    Alignment = textbox1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    textbox1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property
Public Property Get BorderStyle() As Integer
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property
Public Property Get FontName() As String
    FontName = textbox1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    textbox1.FontName() = New_FontName
    PropertyChanged "FontName"
End Property
Public Property Get FontSize() As Single
    FontSize = textbox1.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    textbox1.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = textbox1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    textbox1.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property
Public Property Get FontUnderline() As Boolean
    FontUnderline = textbox1.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    textbox1.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

Public Property Get BorderWidth() As Integer
    If BorderWidth > 5 Then BorderWidth = 5
    BorderWidth = cadre.BorderWidth
    
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    If New_BorderWidth > 5 Then New_BorderWidth = 5
    cadre.BorderWidth() = New_BorderWidth
    textbox1.Left = (New_BorderWidth * 10) + 60
    textbox1.Width = UserControl.ScaleWidth - (textbox1.Left * 3)
    PropertyChanged "BorderWidth"
    
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    cadre.BorderColor = PropBag.ReadProperty("LineColor", &H80000005)
    textbox1.text = PropBag.ReadProperty("Text", "Text1")
    textbox1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    textbox1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    cadre.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    textbox1.Locked = PropBag.ReadProperty("Locked", Faux)
    textbox1.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    textbox1.SelStart = PropBag.ReadProperty("SelStart", 0)
    textbox1.SelText = PropBag.ReadProperty("SelText", "")
    textbox1.SelLength = PropBag.ReadProperty("SelLength", 0)
    textbox1.PasswordChar = PropBag.ReadProperty("PasswordChar", m_def_PasswordChar)
    textbox1.FontBold = PropBag.ReadProperty("FontBold", 0)
    textbox1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
        textbox1.FontName = PropBag.ReadProperty("FontName", "arial")
    textbox1.FontSize = PropBag.ReadProperty("FontSize", 8)
    textbox1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    textbox1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    Set textbox1.Font = PropBag.ReadProperty("Font", "Arial")
    textbox1.Alignment = PropBag.ReadProperty("Alignment", 0)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)

    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    Call PropBag.WriteProperty("Alignment", textbox1.Alignment, 0)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("FontName", textbox1.FontName, "arial")
    Call PropBag.WriteProperty("FontSize", textbox1.FontSize, 8)
    Call PropBag.WriteProperty("FontStrikethru", textbox1.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", textbox1.FontUnderline, 0)
    Call PropBag.WriteProperty("Locked", textbox1.Locked, Faux)
    Call PropBag.WriteProperty("MaxLength", textbox1.MaxLength, 0)
    Call PropBag.WriteProperty("SelStart", textbox1.SelStart, 0)
    Call PropBag.WriteProperty("SelText", textbox1.SelText, "")
    Call PropBag.WriteProperty("SelLength", textbox1.SelLength, 0)
    Call PropBag.WriteProperty("PasswordChar", m_PasswordChar, m_def_PasswordChar)
    Call PropBag.WriteProperty("FontBold", textbox1.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", textbox1.FontItalic, 0)
    Call PropBag.WriteProperty("BorderWidth", cadre.BorderWidth, 1)
    Call PropBag.WriteProperty("LineColor", cadre.BorderColor, &H80000005)
    Call PropBag.WriteProperty("Text", textbox1.text, "Text1")
    Call PropBag.WriteProperty("BackColor", textbox1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", textbox1.ForeColor, &H80000012)
        Call PropBag.WriteProperty("Font", textbox1.Font, "Arial")
    End Sub
