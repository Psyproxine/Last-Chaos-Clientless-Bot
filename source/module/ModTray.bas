Attribute VB_Name = "ModTray"
Option Explicit
Private Declare Function Shell_NotifyIconA Lib "shell32.dll" _
   (ByVal dwMessage As Long, lpData As NOTIFYICONDATAA) As Long
   
Private Declare Function Shell_NotifyIconW Lib "shell32.dll" _
   (ByVal dwMessage As Long, lpData As NOTIFYICONDATAW) As Long


Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

Private Const NOTIFYICON_VERSION = 3

Private Type NOTIFYICONDATAA
   cbSize As Long             ' 4
   hwnd As Long               ' 8
   uID As Long                ' 12
   uFlags As Long             ' 16
   uCallbackMessage As Long   ' 20
   hIcon As Long              ' 24
   szTip As String * 128      ' 152
   dwState As Long            ' 156
   dwStateMask As Long        ' 160
   szInfo As String * 256     ' 416
   uTimeOutOrVersion As Long  ' 420
   szInfoTitle As String * 64 ' 484
   dwInfoFlags As Long        ' 488
   guidItem As Long           ' 492
End Type
Private Type NOTIFYICONDATAW
   cbSize As Long             ' 4
   hwnd As Long               ' 8
   uID As Long                ' 12
   uFlags As Long             ' 16
   uCallbackMessage As Long   ' 20
   hIcon As Long              ' 24
   szTip(0 To 255) As Byte    ' 280
   dwState As Long            ' 284
   dwStateMask As Long        ' 288
   szInfo(0 To 511) As Byte   ' 800
   uTimeOutOrVersion As Long  ' 804
   szInfoTitle(0 To 127) As Byte ' 932
   dwInfoFlags As Long        ' 936
   guidItem As Long           ' 940
End Type


Private nfIconDataA As NOTIFYICONDATAA
Private nfIconDataW As NOTIFYICONDATAW

Private Const NOTIFYICONDATAA_V1_SIZE_A = 88
Private Const NOTIFYICONDATAA_V1_SIZE_U = 152
Private Const NOTIFYICONDATAA_V2_SIZE_A = 488
Private Const NOTIFYICONDATAA_V2_SIZE_U = 936

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const WM_MOUSEMOVE = &H200
Private Declare Function GetVersion Lib "kernel32" () As Long

'Public Event SysTrayMouseDown(ByVal eButton As MouseButtonConstants)
'Public Event SysTrayMouseUp(ByVal eButton As MouseButtonConstants)
'Public Event SysTrayMouseMove()
'Public Event SysTrayDoubleClick(ByVal eButton As MouseButtonConstants)
'Public Event MenuClick(ByVal lIndex As Long, ByVal sKey As String)
'Public Event BalloonShow()
'Public Event BalloonHide()
'Public Event BalloonTimeOut()
'Public Event BalloonClicked()

Public Enum EBalloonIconTypes
   NIIF_NONE = 0
   NIIF_INFO = 1
   NIIF_WARNING = 2
   NIIF_ERROR = 3
   NIIF_NOSOUND = &H10
End Enum

Private m_bAddedMenuItem As Boolean
Private m_iDefaultIndex As Long

Private m_bUseUnicode As Boolean
Private m_bSupportsNewVersion As Boolean

Public Sub ShowBalloonTip(ByVal sMessage As String, Optional ByVal eIcon As EBalloonIconTypes, Optional ByVal lTimeOutMs = 100)
   On Error Resume Next
   Dim lR As Long
   If (m_bSupportsNewVersion) Then
      If (m_bUseUnicode) Then
         StringToArray sMessage, nfIconDataW.szInfo, 512
         StringToArray IIf(Constate = 4, LoginChar(Char).name, User) & " - Aggressive Powered!", nfIconDataW.szInfoTitle, 128
         nfIconDataW.uTimeOutOrVersion = lTimeOutMs
         nfIconDataW.dwInfoFlags = eIcon
         nfIconDataW.uFlags = NIF_INFO
         lR = Shell_NotifyIconW(NIM_MODIFY, nfIconDataW)
      Else
         nfIconDataA.szInfo = sMessage
         nfIconDataA.szInfoTitle = IIf(Constate = 4, LoginChar(Char).name, User) & " - Aggressive Powered!"
         nfIconDataA.uTimeOutOrVersion = lTimeOutMs
         nfIconDataA.dwInfoFlags = eIcon
         nfIconDataA.uFlags = NIF_INFO
         lR = Shell_NotifyIconA(NIM_MODIFY, nfIconDataA)
      End If
   Else
      ' can't do it, fail silently.
   End If
End Sub

Public Property Get ToolTip() As String
Dim sTip As String
Dim iPos As Long
    sTip = nfIconDataA.szTip
    iPos = InStr(sTip, Chr$(0))
    If (iPos <> 0) Then
        sTip = Left$(sTip, iPos - 1)
    End If
    ToolTip = sTip
End Property
Public Property Let ToolTip(ByVal sTip As String)
   If (m_bUseUnicode) Then
      StringToArray sTip, nfIconDataW.szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
      nfIconDataW.uFlags = NIF_TIP
      Shell_NotifyIconW NIM_MODIFY, nfIconDataW
   Else
      If (sTip & Chr$(0) <> nfIconDataA.szTip) Then
         nfIconDataA.szTip = sTip & Chr$(0)
         nfIconDataA.uFlags = NIF_TIP
         Shell_NotifyIconA NIM_MODIFY, nfIconDataA
      End If
   End If
End Property

Public Property Get IconHandle() As Long
    IconHandle = nfIconDataA.hIcon
End Property
Public Property Let IconHandle(ByVal hIcon As Long)
   If (m_bUseUnicode) Then
      If (hIcon <> nfIconDataW.hIcon) Then
         nfIconDataW.hIcon = hIcon
         nfIconDataW.uFlags = NIF_ICON
         Shell_NotifyIconW NIM_MODIFY, nfIconDataW
      End If
   Else
      If (hIcon <> nfIconDataA.hIcon) Then
         nfIconDataA.hIcon = hIcon
         nfIconDataA.uFlags = NIF_ICON
         Shell_NotifyIconA NIM_MODIFY, nfIconDataA
      End If
   End If
End Property

Public Property Get DefaultMenuIndex() As Long
   DefaultMenuIndex = m_iDefaultIndex
End Property

Public Sub AddTray()
On Error Resume Next
   ' Get version:
   Dim lMajor As Long
   Dim lMinor As Long
   Dim bIsNt As Long
   Dim nm As String
        nm = User & " - Aggressive Powered!"

        
   GetWindowsVersion lMajor, lMinor, , , bIsNt

   If (bIsNt) Then
      m_bUseUnicode = True
      If (lMajor >= 5) Then
         ' 2000 or XP
         m_bSupportsNewVersion = True
      End If
   ElseIf (lMajor = 4) And (lMinor = 90) Then
      ' Windows ME
      m_bSupportsNewVersion = True
   End If
   
   
   'Add the icon to the system tray...
   Dim lR As Long
   
   If (m_bUseUnicode) Then
      With nfIconDataW
         .hwnd = frmMain.TRAY.hwnd
         .uID = frmMain.Icon
         .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = frmMain.Icon.handle
         StringToArray nm, .szTip, unicodeSize(IIf(m_bSupportsNewVersion, 128, 64))
         .uCallbackMessage = WM_MOUSEMOVE
         If (m_bSupportsNewVersion) Then
            .uTimeOutOrVersion = NOTIFYICON_VERSION
         End If
         .cbSize = nfStructureSize
      End With
      lR = Shell_NotifyIconW(NIM_ADD, nfIconDataW)
      If (m_bSupportsNewVersion) Then
         Shell_NotifyIconW NIM_SETVERSION, nfIconDataW
      End If
      
   Else
      With nfIconDataA
         .hwnd = frmMain.TRAY.hwnd
         .uID = frmMain.Icon
         .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
         .uCallbackMessage = WM_MOUSEMOVE
         .hIcon = frmMain.Icon.handle
         .szTip = nm & Chr$(0)
         If (m_bSupportsNewVersion) Then
            .uTimeOutOrVersion = NOTIFYICON_VERSION
         End If
         .cbSize = nfStructureSize
      End With
      lR = Shell_NotifyIconA(NIM_ADD, nfIconDataA)
      If (m_bSupportsNewVersion) Then
         lR = Shell_NotifyIconA(NIM_SETVERSION, nfIconDataA)
      End If
   End If
   
End Sub

Private Sub StringToArray( _
      ByVal sString As String, _
      bArray() As Byte, _
      ByVal lMaxSize As Long _
   )
Dim B() As Byte
Dim i As Long
Dim j As Long
   If Len(sString) > 0 Then
      B = sString
      For i = LBound(B) To UBound(B)
         bArray(i) = B(i)
         If (i = (lMaxSize - 2)) Then
            Exit For
         End If
      Next i
      For j = i To lMaxSize - 1
         bArray(j) = 0
      Next j
   End If
End Sub
Private Function unicodeSize(ByVal lSize As Long) As Long
   If (m_bUseUnicode) Then
      unicodeSize = lSize * 2
   Else
      unicodeSize = lSize
   End If
End Function

Private Property Get nfStructureSize() As Long
   If (m_bSupportsNewVersion) Then
      If (m_bUseUnicode) Then
         nfStructureSize = NOTIFYICONDATAA_V2_SIZE_U
      Else
         nfStructureSize = NOTIFYICONDATAA_V2_SIZE_A
      End If
   Else
      If (m_bUseUnicode) Then
         nfStructureSize = NOTIFYICONDATAA_V1_SIZE_U
      Else
         nfStructureSize = NOTIFYICONDATAA_V1_SIZE_A
      End If
   End If
End Property

Public Sub DelTray()
   If (m_bUseUnicode) Then
      Shell_NotifyIconW NIM_DELETE, nfIconDataW
   Else
      Shell_NotifyIconA NIM_DELETE, nfIconDataA
   End If
End Sub

Private Sub GetWindowsVersion( _
      Optional ByRef lMajor = 0, _
      Optional ByRef lMinor = 0, _
      Optional ByRef lRevision = 0, _
      Optional ByRef lBuildNumber = 0, _
      Optional ByRef bIsNt = False _
   )
Dim lR As Long
   lR = GetVersion()
   lBuildNumber = (lR And &H7F000000) \ &H1000000
   If (lR And &H80000000) Then lBuildNumber = lBuildNumber Or &H80
   lRevision = (lR And &HFF0000) \ &H10000
   lMinor = (lR And &HFF00&) \ &H100
   lMajor = (lR And &HFF)
   bIsNt = ((lR And &H80000000) = 0)
End Sub


