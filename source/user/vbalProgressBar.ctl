VERSION 5.00
Begin VB.UserControl vbalProgressBar 
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   450
   ScaleWidth      =   4800
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   315
      Left            =   1140
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   315
      Left            =   1620
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "vbalProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_SOFT = &H1000    ' For softer buttons
Private Const BF_FLAT = &H4000    ' For flat rather than 3D borders
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_SINGLELINE = &H20
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Const OPAQUE = 2
Private Const TRANSPARENT = 1
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Private Declare Function OpenThemeData Lib "uxtheme.dll" _
   (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" _
   (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal lHDC As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pRect As RECT, pClipRect As RECT) As Long
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, _
    ByVal iPartId As Long, ByVal iStateId As Long, _
    pBoundingRect As RECT, pContentRect As RECT) As Long
Private Declare Function DrawThemeText Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, ByVal pszText As Long, _
    ByVal iCharCount As Long, ByVal dwTextFlag As Long, _
    ByVal dwTextFlags2 As Long, pRect As RECT) As Long
Private Declare Function DrawThemeIcon Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, pRect As RECT, _
    ByVal himl As Long, ByVal iImageIndex As Long) As Long
Private Declare Function DrawThemeEdge Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal hdc As Long, ByVal iPartId As Long, _
   ByVal iStateId As Long, pDestRect As RECT, _
   ByVal uEdge As Long, ByVal uFlags As Long, _
   pContentRect As RECT) As Long
Private Enum THEMESIZE
    TS_MIN = 0             '// minimum size
    TS_TRUE = 1            '// size without stretching
    TS_DRAW = 2             ' // size that theme mgr will use to draw part
End Enum
Private Declare Function GetThemeInt Lib "uxtheme.dll" _
   (ByVal hTheme As Long, ByVal iPartId As Long, _
    ByVal iStateId As Long, ByVal iPropId As Long, _
    piVal As Long) As Long
Private Const PROGRESSCHUNKSIZE = 2411
Private Const PROGRESSSPACESIZE = 2412

Private Const S_OK = 0

Public Enum EVPRGAppearanceConstants
   epbaFlat
   epba3DThin
   epba3D
End Enum
Public Enum EVPRGBorderStyleConstants
   epbsNone
   epbsInset
   epbsRaised
End Enum
Public Enum EVPRGPictureModeConstants
   epbpStretch
   epbpTile
End Enum
Public Enum EVPRGHorizontalTextAlignConstants
   epbthLeft
   epbthCenter
   epbthRight
End Enum
Public Enum EVPRGVerticalTextAlignConstants
   epbtvTop
   epbtvVCenter
   epbtvBottom
End Enum

Private m_cMemDC As pcMemDC
Private m_hWnd As Long
Private m_eAppearance As EVPRGAppearanceConstants
Private m_eBorderStyle As EVPRGBorderStyleConstants
Private m_oForeColor As OLE_COLOR
Private m_oBarColor As OLE_COLOR
Private m_oBarForeColor As OLE_COLOR
Private m_eBarPictureMode As EVPRGPictureModeConstants
Private m_eBackPictureMode As EVPRGPictureModeConstants
Private m_lMin As Long
Private m_lMax As Long
Private m_lValue As Long
Private m_eTextAlignX As EVPRGHorizontalTextAlignConstants
Private m_eTextAlignY As EVPRGVerticalTextAlignConstants
Private m_bShowText As Boolean
Private m_sText As String
Private m_bSegments As Boolean
Private m_bXpStyle As Boolean
Public Event Draw(ByVal hdc As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal bDoDefault As Boolean)
Public Property Get Segments() As Boolean
Attribute Segments.VB_Description = "Gets/sets whether the bar is split into segments.  Not applicable when XPStyle is set."
   Segments = m_bSegments
End Property
Public Property Let Segments(ByVal bState As Boolean)
   m_bSegments = bState
   pDraw
   PropertyChanged "Segments"
End Property
Public Property Get XpStyle() As Boolean
Attribute XpStyle.VB_Description = "Gets/sets whether the bar is displayed using XP Visual Styles.  Only valid on XP systems or above."
   XpStyle = m_bXpStyle
End Property
Public Property Let XpStyle(ByVal bState As Boolean)
   m_bXpStyle = bState
   pDraw
   PropertyChanged "XpStyle"
End Property
Public Property Get Text() As String
Attribute Text.VB_Description = "Gets/sets the text to display when ShowText is set."
   Text = m_sText
End Property

Public Property Let Text(ByVal sText As String)
   m_sText = sText
   pDraw
   PropertyChanged "Text"
End Property

Public Property Get TextAlignX() As EVPRGHorizontalTextAlignConstants
Attribute TextAlignX.VB_Description = "Gets/sets the horizontal alignment of the text."
   TextAlignX = m_eTextAlignX
End Property
Public Property Let TextAlignX(ByVal eAlign As EVPRGHorizontalTextAlignConstants)
   m_eTextAlignX = eAlign
   pDraw
   PropertyChanged "TextAlignX"
End Property
Public Property Get TextAlignY() As EVPRGVerticalTextAlignConstants
Attribute TextAlignY.VB_Description = "Gets/sets the vertical alignment of the text."
   TextAlignY = m_eTextAlignY
End Property
Public Property Let TextAlignY(ByVal eAlign As EVPRGVerticalTextAlignConstants)
   m_eTextAlignY = eAlign
   pDraw
   PropertyChanged "TextAlignY"
End Property

Public Property Get ShowText() As Boolean
Attribute ShowText.VB_Description = "Gets/sets whether text is shown over the bar."
   ShowText = m_bShowText
End Property

Public Property Let ShowText(ByVal bState As Boolean)
   m_bShowText = bState
   pDraw
   PropertyChanged "ShowText"
End Property

Public Property Get Percent() As Double
Attribute Percent.VB_Description = "Gets the current progress bar percentage complete."
Dim fPercent As Double
   fPercent = (m_lValue - m_lMin) / (m_lMax - m_lMin)
   If fPercent > 1# Then fPercent = 1#
   If fPercent < 0# Then fPercent = 0#
   Percent = fPercent * 100#
End Property

Public Property Get Min() As Long
Attribute Min.VB_Description = "Gets/sets the minimum value of the bar."
   Min = m_lMin
End Property
Public Property Let Min(ByVal lMin As Long)
   m_lMin = lMin
   pDraw
   PropertyChanged "Min"
End Property
Public Property Get Max() As Long
Attribute Max.VB_Description = "Gets/sets the maximum value of the bar."
   Max = m_lMax
End Property
Public Property Let Max(ByVal lMax As Long)
   m_lMax = lMax
   pDraw
   PropertyChanged "Max"
End Property
Public Property Get Value() As Long
Attribute Value.VB_Description = "Gets/sets the value of the bar."
   Value = m_lValue
End Property
Public Property Let Value(ByVal lValue As Long)
   m_lValue = lValue
   pDraw
   PropertyChanged "Value"
End Property

Public Property Get BorderStyle() As EVPRGBorderStyleConstants
Attribute BorderStyle.VB_Description = "Gets/sets the border style of the control. Not applicable when using XPStyle."
   BorderStyle = m_eBorderStyle
End Property
Public Property Let BorderStyle(ByVal eStyle As EVPRGBorderStyleConstants)
   m_eBorderStyle = eStyle
   pDraw
   PropertyChanged "BorderStyle"
End Property
Public Property Get Appearance() As EVPRGAppearanceConstants
Attribute Appearance.VB_Description = "Gets/sets the border appearance of the control.  Not applicable when using XPStyle."
   Appearance = m_eAppearance
End Property
Public Property Let Appearance(ByVal eAppearance As EVPRGAppearanceConstants)
   m_eAppearance = eAppearance
   pDraw
   PropertyChanged "Appearance"
End Property
Private Sub pDraw()
Dim lHDC As Long
Dim lhDCU As Long
Dim bMem As Boolean
Dim tR As RECT, tBR As RECT, tSR As RECT, tWR As RECT, tXPR As RECT
Dim lWidth As Long, lHeight As Long
Dim lColor As Long
Dim hBr As Long
Dim hRgn As Long
Dim fPercent As Double
Dim bDrawText As Boolean
Dim hFntOld As Long
Dim iFnt As IFont
Dim i As Long
Dim lSegmentWidth As Long, lSegmentSpacing As Long
Dim bDrawnXpStyle As Boolean
Dim hTheme As Long
Dim hR As Long
Dim bDrawn As Boolean
Dim bDoDefault As Boolean
   
   GetClientRect m_hWnd, tR
   lWidth = Abs(tR.Right - tR.Left)
   lHeight = Abs(tR.Bottom - tR.Top)

   lhDCU = UserControl.hdc
   lHDC = m_cMemDC.hdc(lWidth, lHeight)
   If lHDC = 0 Then
      lHDC = lhDCU
   Else
      bMem = True
   End If
   
   bDoDefault = True
   RaiseEvent Draw(lHDC, tR.Left, tR.Top, lWidth, lHeight, bDoDefault)
   If bDoDefault Then
   
      ' Draw background:
      If pbPic(picBack) Then
         If m_eBackPictureMode = epbpTile Then
            TileArea lHDC, 0, 0, lWidth, lHeight, picBack.hdc, picBack.ScaleWidth \ Screen.TwipsPerPixelX, picBack.ScaleHeight \ Screen.TwipsPerPixelY, 0, 0
         Else
            StretchBlt lHDC, 0, 0, lWidth, lHeight, picBack.hdc, 0, 0, picBack.ScaleWidth \ Screen.TwipsPerPixelX, picBack.ScaleHeight \ Screen.TwipsPerPixelY, vbSrcCopy
         End If
      Else
         If (m_bXpStyle) Then
            On Error Resume Next
            hTheme = OpenThemeData(hwnd, StrPtr("Progress"))
            On Error GoTo 0
            If (hTheme <> 0) Then
               hR = GetThemeInt(hTheme, 0, 0, PROGRESSCHUNKSIZE, lSegmentWidth)
               If (hR = S_OK) Then
                  hR = GetThemeInt(hTheme, 0, 0, PROGRESSSPACESIZE, lSegmentSpacing)
                  If (hR = S_OK) Then
                     lSegmentWidth = lSegmentWidth + lSegmentSpacing
                     If (Width > Height) Then
                        hR = DrawThemeBackground(hTheme, lHDC, 1, 0, tR, tR)
                     Else
                        hR = DrawThemeBackground(hTheme, lHDC, 2, 0, tR, tR)
                     End If
                     If (hR = S_OK) Then
                        bDrawn = True
                     End If
                  End If
               End If
            End If
         End If
         
         If Not (bDrawn) Then
            lColor = UserControl.BackColor
            If lColor And &H80000000 Then
               hBr = GetSysColorBrush(lColor And &H1F&)
            Else
               hBr = CreateSolidBrush(lColor)
            End If
            FillRect lHDC, tR, hBr
            DeleteObject hBr
         End If
      End If
         
      If (m_bSegments) And Not (bDrawn) Then
         lSegmentWidth = 8
         lSegmentSpacing = 2
      End If
   
         
      LSet tWR = tR
      If m_eBorderStyle > epbsNone Then
         If bDrawn Then
            InflateRect tR, -1, -1
         Else
            If m_eAppearance = epba3D Then
               InflateRect tR, -2, -2
            Else
               InflateRect tR, -1, -1
            End If
         End If
      End If
      
      If (m_bShowText) And Len(m_sText) > 0 Then
         bDrawText = True
      End If
      If (bDrawText) And Not (bDrawn) Then
         Set iFnt = UserControl.Font
         hFntOld = SelectObject(lHDC, iFnt.hFont)
         SetBkMode lHDC, TRANSPARENT
         SetTextColor lHDC, TranslateColor(m_oForeColor)
         DrawText lHDC, " " & m_sText & " ", -1, tR, DT_SINGLELINE Or m_eTextAlignX Or m_eTextAlignY * 4
         SelectObject lHDC, hFntOld
      End If
      
      ' Draw bar:
      ' Get the bar rectangle:
      LSet tBR = tR
      fPercent = (m_lValue - m_lMin) / (m_lMax - m_lMin)
      If fPercent > 1# Then fPercent = 1#
      If fPercent < 0# Then fPercent = 0#
      If Width > Height Then
         tBR.Right = tR.Left + (tR.Right - tR.Left) * fPercent
         If (m_bSegments Or bDrawn) Then
            ' Quantise bar:
            tBR.Right = tBR.Right - ((tBR.Right - tBR.Left) Mod (lSegmentWidth + lSegmentSpacing))
            'Debug.Assert ((tBR.Right - tBR.Left) Mod (lSegmentWidth + lSegmentSpacing) = 0)
            If tBR.Right < tR.Left Then
               tBR.Right = tR.Left
            End If
         End If
      Else
         fPercent = 1# - fPercent
         tBR.Top = tR.Top + (tR.Bottom - tR.Top) * fPercent
         If (m_bSegments Or bDrawn) Then
            ' Quantise bar:
            tBR.Top = tBR.Top - ((tBR.Top - tBR.Bottom) Mod (lSegmentWidth + lSegmentSpacing))
            If tBR.Top > tR.Bottom Then
               tBR.Top = tR.Bottom
            End If
         End If
      End If
      
      If Not bDrawn Then
         hRgn = CreateRectRgnIndirect(tBR)
         SelectClipRgn lHDC, hRgn
      End If
      
      If pbPic(picBar) Then
         If m_eBarPictureMode = epbpTile Then
            TileArea lHDC, 0, tBR.Top, tBR.Right - tBR.Left, tBR.Bottom - tBR.Top, picBar.hdc, picBar.ScaleWidth \ Screen.TwipsPerPixelX, picBar.ScaleHeight \ Screen.TwipsPerPixelY, 0, 0
         Else
            StretchBlt lHDC, 0, 0, lWidth, lHeight, picBar.hdc, 0, 0, picBar.ScaleWidth \ Screen.TwipsPerPixelX, picBar.ScaleHeight \ Screen.TwipsPerPixelY, vbSrcCopy
         End If
      Else
         If bDrawn Then
            LSet tXPR = tBR
            InflateRect tXPR, -2, -2
            tXPR.Right = tXPR.Right + 1
            tXPR.Bottom = tXPR.Bottom + 1
            If (Width > Height) Then
               hR = DrawThemeBackground(hTheme, lHDC, 3, 0, tXPR, tXPR)
            Else
               hR = DrawThemeBackground(hTheme, lHDC, 4, 0, tXPR, tXPR)
            End If
         Else
            lColor = m_oBarColor
            If lColor And &H80000000 Then
               hBr = GetSysColorBrush(lColor And &H1F&)
            Else
               hBr = CreateSolidBrush(lColor)
            End If
            FillRect lHDC, tBR, hBr
            DeleteObject hBr
         End If
      End If
      
      If m_bSegments And Not bDrawn Then
         lColor = UserControl.BackColor
         If lColor And &H80000000 Then
            hBr = GetSysColorBrush(lColor And &H1F&)
         Else
            hBr = CreateSolidBrush(lColor)
         End If
         LSet tSR = tR
         If Width > Height Then
            For i = tBR.Left + lSegmentWidth To tBR.Right Step lSegmentWidth + lSegmentSpacing
               tSR.Left = i
               tSR.Right = i + lSegmentSpacing
               FillRect lHDC, tSR, hBr
            Next i
         Else
            For i = tBR.Bottom To tBR.Top + lSegmentWidth Step -(lSegmentWidth + lSegmentSpacing)
               tSR.Top = i
               tSR.Bottom = i + lSegmentSpacing
               FillRect lHDC, tSR, hBr
            Next i
         End If
         DeleteObject hBr
      End If
         
      If bDrawText Then
         Set iFnt = UserControl.Font
         hFntOld = SelectObject(lHDC, iFnt.hFont)
         If (bDrawn) Then
            Dim rcContent As RECT
            hR = GetThemeBackgroundContentRect(hTheme, _
                   lHDC, 0, 0, tR, rcContent)
            hR = DrawThemeText(hTheme, lHDC, 0, 0, _
               StrPtr(m_sText), -1, _
               DT_SINGLELINE Or m_eTextAlignX Or m_eTextAlignY * 4, _
               0, rcContent)
         Else
            SetBkMode lHDC, TRANSPARENT
            SetTextColor lHDC, TranslateColor(m_oBarForeColor)
            DrawText lHDC, " " & m_sText & " ", -1, _
               tR, DT_SINGLELINE Or m_eTextAlignX Or m_eTextAlignY * 4
         End If
         SelectObject lHDC, hFntOld
      End If
         
      If Not bDrawn Then
         SelectClipRgn lHDC, 0
         DeleteObject hRgn
         
         ' Draw border:
         Select Case m_eBorderStyle
         Case epbsRaised
            Select Case m_eAppearance
            Case epbaFlat
               Border lHDC, epbaFlat, tWR, True
            Case epba3DThin
               Border lHDC, epba3DThin, tR, True
            Case epba3D
               Border lHDC, epba3D, tWR, True
            End Select
         Case epbsInset
            Select Case m_eAppearance
            Case epbaFlat
               Border lHDC, epbaFlat, tWR, False
            Case epba3DThin
               Border lHDC, epba3DThin, tWR, False
            Case epba3D
               Border lHDC, epba3D, tWR, False
            End Select
         End Select
      End If
   
   End If
   
   ' Swap memdc<->Screen
   If bMem Then
      m_cMemDC.Draw lhDCU, 0, 0, lWidth, lHeight
   End If
   
   If (hTheme) Then
      CloseThemeData hTheme
   End If

End Sub

Private Function pbPic(ByVal picThis As PictureBox) As Boolean
   If Not (picThis.Picture Is Nothing) Then
      If Not picThis.Picture.handle = 0 Then
         pbPic = True
      End If
   End If
End Function
Private Sub Border( _
      ByVal lHDC As Long, _
      ByVal lStyle As Long, _
      ByRef tR As RECT, _
      ByVal bRaised As Boolean _
   )
Dim hPenDark As Long, hPenLight As Long, hPenBlack As Long
Dim hPenOld As Long
Dim tJunk As POINTAPI

   Select Case lStyle
   Case 0
      hPenBlack = CreatePen(0, 1, 0)
      hPenOld = SelectObject(lHDC, hPenBlack)
      MoveToEx lHDC, tR.Left, tR.Top, tJunk
      LineTo lHDC, tR.Right - 1, tR.Top
      LineTo lHDC, tR.Right - 1, tR.Bottom - 1
      LineTo lHDC, tR.Left, tR.Bottom - 1
      LineTo lHDC, tR.Left, tR.Top
      SelectObject lHDC, hPenOld
      DeleteObject hPenBlack
   Case 1
      hPenDark = CreatePen(0, 1, GetSysColor(vbButtonShadow And &H1F&))
      hPenLight = CreatePen(0, 1, GetSysColor(vb3DHighlight And &H1F&))
      If bRaised Then
         MoveToEx lHDC, tR.Left, tR.Bottom - 2, tJunk
         hPenOld = SelectObject(lHDC, hPenLight)
         LineTo lHDC, tR.Left, tR.Top
         LineTo lHDC, tR.Right - 1, tR.Top
         SelectObject lHDC, hPenOld
         MoveToEx lHDC, tR.Right - 1, tR.Top, tJunk
         hPenOld = SelectObject(lHDC, hPenDark)
         LineTo lHDC, tR.Right - 1, tR.Bottom - 1
         LineTo lHDC, tR.Left - 1, tR.Bottom - 1
         SelectObject lHDC, hPenOld
      Else
         MoveToEx lHDC, tR.Left, tR.Bottom - 1, tJunk
         hPenOld = SelectObject(lHDC, hPenDark)
         LineTo lHDC, tR.Left, tR.Top
         LineTo lHDC, tR.Right, tR.Top
         SelectObject lHDC, hPenOld
         MoveToEx lHDC, tR.Right - 1, tR.Top + 1, tJunk
         hPenOld = SelectObject(lHDC, hPenLight)
         LineTo lHDC, tR.Right - 1, tR.Bottom - 1
         LineTo lHDC, tR.Left, tR.Bottom - 1
         SelectObject lHDC, hPenOld
      End If
      DeleteObject hPenDark
      DeleteObject hPenLight
   Case 2
      If bRaised Then
         DrawEdge lHDC, tR, EDGE_RAISED, BF_RECT Or BF_SOFT
      Else
         DrawEdge lHDC, tR, EDGE_SUNKEN, BF_RECT Or BF_SOFT
      End If
   End Select
End Sub
      
Private Sub TileArea( _
        ByVal hdc As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long, _
        ByVal lSrcDC As Long, _
        ByVal lBitmapW As Long, _
        ByVal lBitmapH As Long, _
        ByVal lSrcOffsetX As Long, _
        ByVal lSrcOffsetY As Long _
    )
Dim lSrcX As Long
Dim lSrcY As Long
Dim lSrcStartX As Long
Dim lSrcStartY As Long
Dim lSrcStartWidth As Long
Dim lSrcStartHeight As Long
Dim lDstX As Long
Dim lDstY As Long
Dim lDstWidth As Long
Dim lDstHeight As Long

    lSrcStartX = ((x + lSrcOffsetX) Mod lBitmapW)
    lSrcStartY = ((y + lSrcOffsetY) Mod lBitmapH)
    lSrcStartWidth = (lBitmapW - lSrcStartX)
    lSrcStartHeight = (lBitmapH - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (y + Height)
        If (lDstY + lDstHeight) > (y + Height) Then
            lDstHeight = y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = x
        lSrcX = lSrcStartX
        Do While lDstX < (x + Width)
            If (lDstX + lDstWidth) > (x + Width) Then
                lDstWidth = x + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt hdc, lDstX, lDstY, lDstWidth, lDstHeight, lSrcDC, lSrcX, lSrcY, vbSrcCopy
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = lBitmapW
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = lBitmapH
    Loop
End Sub


Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gets/sets the back colour of the control. Not applicable when using XPStyle."
   BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(oColor As OLE_COLOR)
   UserControl.BackColor = oColor
   pDraw
   PropertyChanged "BackColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Gets/sets the colour of text which is drawn over the background of the bar."
   ForeColor = m_oForeColor
End Property
Public Property Let ForeColor(oColor As OLE_COLOR)
   m_oForeColor = oColor
   pDraw
   PropertyChanged "ForeColor"
End Property
Public Property Get Font() As IFont
Attribute Font.VB_Description = "Gets/sets the font used to draw text on the progress bar."
   Set Font = UserControl.Font
End Property
Public Property Set Font(ByRef fnt As IFont)
   Set UserControl.Font = fnt
   pDraw
   PropertyChanged "Font"
End Property
Public Property Let Font(ByRef fnt As IFont)
   Set UserControl.Font = fnt
   pDraw
   PropertyChanged "Font"
End Property
Public Property Get BarColor() As OLE_COLOR
Attribute BarColor.VB_Description = "Gets/sets the colour of the progress bar.  Not applicable when using XPStyle."
   BarColor = m_oBarColor
End Property
Public Property Let BarColor(oColor As OLE_COLOR)
   m_oBarColor = oColor
   pDraw
   PropertyChanged "BarColor"
End Property
Public Property Get BarForeColor() As OLE_COLOR
Attribute BarForeColor.VB_Description = "Gets/sets the colour of text which is drawn over the bar."
   BarForeColor = m_oBarForeColor
End Property
Public Property Let BarForeColor(oColor As OLE_COLOR)
   m_oBarForeColor = oColor
   pDraw
   PropertyChanged "BarForeColor"
End Property

Public Function ModifyBarPicture( _
      Optional ByVal fLuminance As Double = 1, _
      Optional ByVal fSaturation As Double = 1 _
   )
Attribute ModifyBarPicture.VB_Description = "Applies image processing to the bar picture, allowing you to adjust the luminance or saturation of the image."
   If (pbPic(picBar)) Then
      Dim cDib As New pcDibSection
      cDib.CreateFromPicture picBar
      cDib.ModifyHLS 1, fLuminance, fSaturation
      cDib.PaintPicture picBar.hdc
   End If
End Function
Public Function ModifyPicture( _
      Optional ByVal fLuminance As Double = 1, _
      Optional ByVal fSaturation As Double = 1 _
   )
Attribute ModifyPicture.VB_Description = "Applies image processing to the background picture, allowing you to adjust the luminance or saturation of the image."
   If (pbPic(picBack)) Then
      Dim cDib As New pcDibSection
      cDib.CreateFromPicture picBack
      cDib.ModifyHLS 1, fLuminance, fSaturation
      cDib.PaintPicture picBack.hdc
   End If
End Function

Public Property Get BarPicture() As IPicture
Attribute BarPicture.VB_Description = "Gets/sets a picture to use as the bar in the progress bar."
   Set BarPicture = picBar.Picture
End Property
Public Property Let BarPicture(pic As IPicture)
   pPicture pic, picBar
End Property
Public Property Set BarPicture(pic As IPicture)
   pPicture pic, picBar
End Property
Public Property Get BarPictureMode() As EVPRGPictureModeConstants
Attribute BarPictureMode.VB_Description = "Gets/sets the drawing mode (stretch or tile) applied when drawing the bar picture."
   BarPictureMode = m_eBarPictureMode
End Property
Public Property Let BarPictureMode(ByVal eMode As EVPRGPictureModeConstants)
   m_eBarPictureMode = eMode
   pDraw
   PropertyChanged "BarPictureMode"
End Property
Public Property Get BackPictureMode() As EVPRGPictureModeConstants
Attribute BackPictureMode.VB_Description = "Gets/sets the drawing mode (stretch or tile) applied when drawing the background picture."
   BackPictureMode = m_eBackPictureMode
End Property
Public Property Let BackPictureMode(ByVal eMode As EVPRGPictureModeConstants)
   m_eBackPictureMode = eMode
   pDraw
   PropertyChanged "BackPictureMode"
End Property

Public Property Get Picture() As IPicture
Attribute Picture.VB_Description = "Gets/sets the picture shown in the background of the progress bar control."
   Set Picture = picBack.Picture
End Property
Public Property Let Picture(pic As IPicture)
   pPicture pic, picBack
End Property
Public Property Set Picture(pic As IPicture)
   pPicture pic, picBack
End Property
Private Sub pPicture(pic As IPicture, picStore As PictureBox)
   Set picStore.Picture = pic
   pDraw
   PropertyChanged "Picture"
   PropertyChanged "BarPicture"
End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Sub UserControl_Initialize()
   Set m_cMemDC = New pcMemDC
   m_eAppearance = epba3DThin
   m_eBorderStyle = epbsInset
   m_oBarColor = &H800000
   m_oBarForeColor = &HFFFFFF
   m_eBarPictureMode = epbpTile
   m_eBackPictureMode = epbpTile
   m_lMax = 100
   m_eTextAlignX = epbthCenter
   m_eTextAlignY = epbtvVCenter
End Sub

Private Sub UserControl_InitProperties()
   m_hWnd = UserControl.hwnd
End Sub

Private Sub UserControl_Paint()
   pDraw
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_hWnd = UserControl.hwnd
   Picture = PropBag.ReadProperty("Picture", Nothing)
   BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
   ForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
   Appearance = PropBag.ReadProperty("Appearance", epba3DThin)
   BorderStyle = PropBag.ReadProperty("BorderStyle", epbsInset)
   BarColor = PropBag.ReadProperty("BarColor", &H800000)
   BarForeColor = PropBag.ReadProperty("BarForeColor", &HFFFFFF)
   BarPicture = PropBag.ReadProperty("BarPicture", Nothing)
   BarPictureMode = PropBag.ReadProperty("BarPictureMode", epbpTile)
   BackPictureMode = PropBag.ReadProperty("BackPictureMode", epbpTile)
   Min = PropBag.ReadProperty("Min", 0)
   Max = PropBag.ReadProperty("Max", 100)
   Value = PropBag.ReadProperty("Value", 0)
   ShowText = PropBag.ReadProperty("ShowText", False)
   TextAlignX = PropBag.ReadProperty("TextAlignX", epbthCenter)
   TextAlignY = PropBag.ReadProperty("TextAlignY", epbtvVCenter)
   Text = PropBag.ReadProperty("Text", "")
   Font = PropBag.ReadProperty("Font", UserControl.Font)
   Segments = PropBag.ReadProperty("Segments", False)
   XpStyle = PropBag.ReadProperty("XpStyle", False)
End Sub

Private Sub UserControl_Resize()
   pDraw
End Sub

Private Sub UserControl_Terminate()
   Set m_cMemDC = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "Picture", Picture, Nothing
   PropBag.WriteProperty "BackColor", BackColor, vbButtonFace
   PropBag.WriteProperty "ForeColor", ForeColor, vbWindowText
   PropBag.WriteProperty "Appearance", Appearance, epba3DThin
   PropBag.WriteProperty "BorderStyle", BorderStyle, epbsInset
   PropBag.WriteProperty "BarColor", BarColor, &H800000
   PropBag.WriteProperty "BarForeColor", BarForeColor, &HFFFFFF
   PropBag.WriteProperty "BarPicture", BarPicture, Nothing
   PropBag.WriteProperty "BarPictureMode", BarPictureMode, epbpTile
   PropBag.WriteProperty "BackPictureMode", BackPictureMode, epbpTile
   PropBag.WriteProperty "Min", Min, 0
   PropBag.WriteProperty "Max", Max, 100
   PropBag.WriteProperty "Value", Value, 0
   PropBag.WriteProperty "ShowText", ShowText, False
   PropBag.WriteProperty "TextAlignX", TextAlignX, epbthCenter
   PropBag.WriteProperty "TextAlignY", TextAlignY, epbtvVCenter
   PropBag.WriteProperty "Text", Text, ""
   PropBag.WriteProperty "Font", Font
   PropBag.WriteProperty "Segments", Segments, False
   PropBag.WriteProperty "XpStyle", XpStyle, False
End Sub


