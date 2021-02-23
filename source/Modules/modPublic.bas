Attribute VB_Name = "modPublic"
Option Explicit

Public Const vbRelFolder As Long = 99
Public Const vbRelNone As Long = 0

Public m_Scheme As New IScheme     ' The perspective color scheme
Public m_Views As New List         ' This list stores all registered views
Public m_Editors As New List       ' This list stores all registered editors
Public m_Windows As New List       ' This list stores all floating windows
Public m_Placeholders As New List  ' View placeholders of a folder
Public m_CursorPos As PointAPI

Private m_FormStyle As Long

Public Sub PrintText(ByVal Text As String, _
                     ByVal hDc As Long, _
                     ByVal x As Long, _
                     ByVal y As Long, _
            Optional ByVal Winkel As Long = 0, _
            Optional ByVal FontSize As Long = 12, _
            Optional ByVal FontWidth As VbFontWidth = fwStandard, _
            Optional ByVal Color As OLE_COLOR = vbWhite, _
            Optional ByVal Italic As Long = 0, _
            Optional ByVal FontName As String = "Arial", _
            Optional ByVal Underline As Long = 0)

   Dim hFont As Long
   Dim FontMem As Long
   Dim Result As Long

   hFont = CreateFont(-FontSize, 0, Winkel * 10, Winkel * 10, FontWidth, Italic, _
                      Underline, 0, 1, 7, 0, 0, 0, FontName)

   SetTextColor hDc, Color
   FontMem = SelectObject(hDc, hFont)
   Result = TextOut(hDc, x, y, Text, Len(Text))
   Result = SelectObject(hDc, FontMem)
   Result = DeleteObject(hFont)
   
End Sub




' Checks the syntax of an id. The id should not contain the following
' letters: . % & / \ ?.
'
' @Id The id to check
'
' @Throws Returns false if the syntax of the id is unsupported; true otherwise.
Public Sub CheckId(ByVal Id As String)
   
   If InStr(1, Id, ".", vbBinaryCompare) > 0 Or _
      InStr(1, Id, "%", vbBinaryCompare) > 0 Or _
      InStr(1, Id, "&", vbBinaryCompare) > 0 Or _
      InStr(1, Id, "/", vbBinaryCompare) > 0 Or _
      InStr(1, Id, "\", vbBinaryCompare) > 0 Or _
      InStr(1, Id, "?", vbBinaryCompare) > 0 Then
      
      ' Id is unsupported
      Err.Raise 3735
   End If
      
End Sub

Public Sub SetWindowStyle(ByVal hWnd As Long, ByVal BorderStyle As VbWindowStyle)
     
   Dim l_Style As Long
  
   If m_FormStyle = 0 Then
      m_FormStyle = GetWindowLong(hWnd, GWL_STYLE)
   End If
   
   ' set new window style
   If BorderStyle = VbNone Then
      
      l_Style = m_FormStyle And Not WS_DLGFRAME And Not WS_EX_APPWINDOW _
                            And Not WS_BORDER And Not WS_EX_WINDOWEDGE Or _
                            WS_EX_MDICHILD Or WS_CHILDWINDOW And Not WS_EX_NOPARENTNOTIFY
   
      SetWindowLong hWnd, GWL_STYLE, l_Style
   
      ' tell the nonclient area of the form to redraw itself
      Call SetWindowPos(hWnd, 0&, 0&, 0&, 0&, 0&, SWP_FRAMECHANGED)
   
   Else
   
      l_Style = GetWindowLong(hWnd, GWL_STYLE) ' Get current style
      l_Style = l_Style And Not WS_CAPTION Or WS_EX_NOPARENTNOTIFY
   
      SetWindowLong hWnd, GWL_STYLE, l_Style
    
      ' tell the nonclient area of the form to redraw itself
      Call SetWindowPos(hWnd, 0&, -100, -100, 10, 10, SWP_FRAMECHANGED)
               
   End If

End Sub

' Returns the style of the selected theme scheme.
'
' @Scheme A constant of the VbWindowScheme enumeration.
Public Function Scheme() As VbWindowsScheme
   
   On Error GoTo ERROR_HANDLE
   
   Dim SchemeName As String
   Dim RegKeyTheme As String
   
   RegKeyTheme = "Software\Microsoft\Windows\CurrentVersion\ThemeManager"
   
   SchemeName = GetKeyValue(VbHKEY_CURRENT_USER, RegKeyTheme, "ColorName")
   
   Select Case SchemeName
      Case "NormalColor":    Scheme = VbNormalColor
      Case "HomeStead":      Scheme = VbHomeStead
      Case "Metallic":       Scheme = VbMetallic
      Case Else:             Scheme = VbClassic
   End Select
  
Finally:
      
   Exit Function
   
ERROR_HANDLE:

   Scheme = VbClassic
   
   GoTo Finally

End Function

Public Sub DrawDragRect(Rc As RECT, Optional ByVal Size As Long = 2)
        
   Dim DrawRect As RECT
   Dim hDc As Long
   Dim i As Long
        
   hDc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)

   For i = 0 To Size
           
      With DrawRect
         .Top = Rc.Top + i
         .Bottom = Rc.Bottom - i
         .Left = Rc.Left + i
         .Right = Rc.Right - i
      End With
           
      DrawFocusRect hDc, DrawRect
           
   Next i
        
   DeleteDC hDc
        
End Sub
