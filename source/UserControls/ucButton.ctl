VERSION 5.00
Begin VB.UserControl ucButton 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   ScaleHeight     =   1860
   ScaleWidth      =   2685
   ToolboxBitmap   =   "ucButton.ctx":0000
   Begin VB.Timer tmrHover 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   480
      Top             =   0
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   0
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "ucButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Active As Boolean
Private m_Hover As Boolean
Private m_Pressed As Boolean

Public Event Click()
'Standard-Eigenschaftswerte:
Const m_def_Image As String = ""
'Eigenschaftsvariablen:
Dim m_Image As String

Public Property Get Active() As Boolean
    Active = m_Active
End Property
Public Property Let Active(ByVal NewActive As Boolean)
    
    If m_Active <> NewActive Then
       m_Active = NewActive
       Refresh
    End If
           
End Property

Public Function hWnd() As Long
    hWnd = UserControl.hWnd
End Function
Public Function hDc() As Long
    hDc = UserControl.hDc
End Function

Public Sub Refresh()
     
   Dim ParentObject As Object
   Dim l_Gradient As Gradient
   Dim strImg As String
      
   On Error Resume Next
   Set ParentObject = Parent
   On Error GoTo 0
   
   ' Draw button style if its parent is a caption (not on tabstrip)
   If TypeName(ParentObject) = "ucCaption" Then
   
      Select Case m_Scheme.CaptionStyle
           
         Case vbHorizontalGradient:
              
            ' If caption has a horizontal gradient then set backcolor to ending gradient color.
            If m_Active Then
               Me.BackColor = m_Scheme.ActiveCaptionGradient2
            Else
               Me.BackColor = m_Scheme.InactiveCaptionGradient2
            End If
           
         Case vbVerticalGradient:  ' Draw a gradient caption
           
            Set l_Gradient = New Gradient
              
            With l_Gradient
                 
               If m_Active Then
                  .Color1 = m_Scheme.ActiveCaptionGradient1
                  .Color2 = m_Scheme.ActiveCaptionGradient2
               Else
                  .Color1 = m_Scheme.InactiveCaptionGradient1
                  .Color2 = m_Scheme.InactiveCaptionGradient2
               End If
                   
               .Angle = 90
                   
               .Draw Me
           
           End With
             
   '      Case vbVisualStudioNET:
           
      End Select
      
   Else
      UserControl.Cls
      Me.BackColor = m_Scheme.BackColor
   End If
   
   UserControl.Refresh
   
   If Me.Enabled Then
   
      If Active Then
         strImg = Me.Image & "_ACTIVE"
      Else
         strImg = Me.Image & "_INACTIVE"
      End If
      
      If m_Hover Then
         strImg = strImg & "_HOVER"
         If m_Pressed Then
            strImg = strImg & "_PRESS"
         End If
      End If
   Else
      strImg = Me.Image & "_DISABLED"
   End If
   
   imgIcon.Picture = getResourceIcon(strImg)
   
End Sub
Private Sub imgIcon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   UserControl_MouseDown Button, Shift, x, y
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Not m_Pressed Then
      m_Pressed = True
      Refresh
   End If
End Sub
Private Sub imgIcon_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   UserControl_MouseUp Button, Shift, x, y
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If m_Pressed Then
      m_Pressed = False
      Refresh
   End If
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Not m_Hover Then
      tmrHover.Enabled = True
      'm_Pressed = False
      m_Hover = True
      Refresh
   End If
End Sub
Private Sub imgIcon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   UserControl_MouseMove Button, Shift, x, y
End Sub

Private Sub tmrHover_Timer()
   
   Dim r As RECT
   Dim Pt As PointAPI
   
   GetWindowRect UserControl.hWnd, r
   GetCursorPos Pt
      
   If modDeclare.PtInRect(r, Pt.x, Pt.y) = 0 Then
      m_Hover = False
      tmrHover.Enabled = False
      Refresh
   End If
   
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Gibt eine Grafik zurück, die in einem Steuerelement angezeigt werden soll, oder legt diese fest."
    Set Picture = imgIcon.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set imgIcon.Picture = New_Picture
    imgIcon.Move (ScaleWidth / imgIcon.Width) / 2, (ScaleHeight / imgIcon.Height) / 2
    PropertyChanged "Picture"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=imgIcon,imgIcon,-1,ToolTipText
Public Property Get ToolTip() As String
Attribute ToolTip.VB_Description = "Gibt den Text zurück, der angezeigt wird, wenn die Maus über dem Steuerelement verweilt, oder legt den Text fest."
    ToolTip = imgIcon.ToolTipText
End Property

Public Property Let ToolTip(ByVal New_ToolTipText As String)
    imgIcon.ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Private Sub imgIcon_Click()
   RaiseEvent Click
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    imgIcon.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    
    imgIcon.Picture = getResourceIcon(ICON_VIEW_CLOSE_ACTIVE)
    m_Image = PropBag.ReadProperty("Image", m_def_Image)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("ToolTipText", imgIcon.ToolTipText, "")
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Image", m_Image, m_def_Image)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gibt die Hintergrundfarbe zurück, die verwendet wird, um Text und Grafik in einem Objekt anzuzeigen, oder legt diese fest."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=13,0,0,
Public Property Get Image() As String
    Image = m_Image
End Property

Public Property Let Image(ByVal New_Image As String)
    m_Image = New_Image
    PropertyChanged "Image"
End Property

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_Image = m_def_Image
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gibt einen Wert zurück, der bestimmt, ob ein Objekt auf vom Benutzer erzeugte Ereignisse reagieren kann, oder legt diesen fest."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Refresh
    PropertyChanged "Enabled"
End Property

