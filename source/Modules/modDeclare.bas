Attribute VB_Name = "modDeclare"
Option Explicit


Public Const SW_HIDE As Long = 0           ' Hide window constant
Public Const SW_SHOW As Long = 5           ' Show window constant

Public Const FW_DONTCARE As Long = 0       ' Standard
Public Const FW_THIN As Long = 100         ' Thin
Public Const FW_EXTRALIGHT As Long = 200   ' Extra light
Public Const FW_LIGHT As Long = 300        ' Light
Public Const FW_NORMAL As Long = 400       ' Normal
Public Const FW_MEDIUM As Long = 500       ' Medium
Public Const FW_SEMIBOLD As Long = 600     ' Semi bold
Public Const FW_BOLD As Long = 700         ' Bold
Public Const FW_EXTRABOLD As Long = 800    ' Extra bold
Public Const FW_HEAVY As Long = 900        ' Heavy

' Registry Root Keys
Public Const HKEY_CLASSES_ROOT As Long = &H80000000     ' Root
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002    ' Local Machine
Public Const HKEY_USERS As Long = &H80000003            ' Users
Public Const HKEY_CURRENT_USER As Long = &H80000001     ' Current User
Public Const HKEY_CURRENT_CONFIG As Long = &H80000005   ' Current Config

Public Const KEY_READ As Long = &H20019    ' Read Access
Public Const REG_SZ As Long = 1            ' VBNullChar terminated String

Public Const WS_EX_MDICHILD As Long = &H40&
Public Const WS_EX_WINDOWEDGE As Long = &H100&
Public Const WS_CAPTION As Long = &HC00000
Public Const WS_BORDER As Long = &H800000
Public Const WS_DLGFRAME As Long = &H400000
Public Const WS_CHILD As Long = &H40000000
Public Const WS_CHILDWINDOW As Long = (WS_CHILD)
Public Const WM_MOUSEACTIVATE As Long = &H21
'Public Const WM_CLOSE As Long = &H10
Public Const WM_COMMAND As Long = &H111

Public Const WS_EX_APPWINDOW As Long = &H40000
Public Const WS_EX_NOPARENTNOTIFY As Long = &H4&

Public Const GWL_STYLE As Long = (-16)
Public Const GWL_HWNDPARENT As Long = -8

Public Const SWP_FRAMECHANGED As Long = &H20 ' Set window position constant - sends message frame changed to the window

' Rectangle
Public Type RECT
   Left As Long     ' Left of the rectangle
   Top As Long      ' Top of the rectangle
   Right As Long    ' Right of the rectangle
   Bottom As Long   ' Bottom of the rectangle
End Type

' Point
Public Type POINTAPI
   X As Long        ' X position of the point.
   y As Long        ' Y position of the point.
End Type

' Registry keys
Public Enum VbHKey
   VbHKEY_CLASSES_ROOT = HKEY_CLASSES_ROOT
   VbHKEY_LOCAL_MACHINE = HKEY_LOCAL_MACHINE
   VbHKEY_USERS = HKEY_USERS
   VbHKEY_CURRENT_USER = HKEY_CURRENT_USER
   VbHKEY_CURRENT_CONFIG = HKEY_CURRENT_CONFIG
End Enum

' Windows scheme constants.
Public Enum VbWindowsScheme
   VbClassic = 0        ' Classic
   VbNormalColor = 1    ' Normal Color (Blue)
   VbMetallic = 2       ' Metallic (Silver)
   VbHomeStead = 3      ' HomeStead (Olive)
End Enum

' Window border style constants.
Public Enum VbWindowStyle
   VbNone = 0           ' No border
   VbToolWin = 1        ' Tool window
End Enum

' Font width constants.
Public Enum VbFontWidth
   fwStandard = FW_DONTCARE         ' Standard font
   fwThin = FW_THIN                 ' Thin font
   fwExtraLight = FW_EXTRALIGHT     ' Extra light font
   fwLight = FW_LIGHT               ' Light font
   fwNormal = FW_NORMAL             ' Normat font
   fwMedium = FW_MEDIUM             ' Medium font
   fwSemiBold = FW_SEMIBOLD         ' Semi bold font
   fwBold = FW_BOLD                 ' Bold font
   fwExtraBold = FW_EXTRABOLD       ' Extra bold font
   fwHeavy = FW_HEAVY               ' Heavy font
End Enum

Public Declare Function PtInRect Lib "user32.dll" ( _
   ByRef lpRect As RECT, _
   ByVal X As Long, _
   ByVal y As Long _
) As Long

Public Declare Function GetWindowRect Lib "user32.dll" ( _
   ByVal hwnd As Long, _
   ByRef lpRect As RECT _
) As Long

Public Declare Function GetCursorPos Lib "user32.dll" ( _
   ByRef lpPoint As POINTAPI _
) As Long

Public Declare Function CreateDCAsNull Lib "gdi32" _
   Alias "CreateDCA" ( _
   ByVal lpDriverName As String, _
   lpDeviceName As Any, _
   lpOutput As Any, _
   lpInitData As Any _
) As Long

Public Declare Function CopyRect Lib "user32.dll" ( _
   ByRef lpDestRect As RECT, _
   ByRef lpSourceRect As RECT _
) As Long

Public Declare Function DeleteDC Lib "gdi32" ( _
   ByVal hdc As Long _
) As Long

Public Declare Function DrawFocusRect Lib "user32" ( _
   ByVal hdc As Long, _
   lpRect As RECT _
) As Long

Public Declare Function RegOpenKeyEx Lib "advapi32" _
   Alias "RegOpenKeyExA" ( _
   ByVal HKey As Long, _
   ByVal lpSubKey As String, _
   ByVal ulOptions As Long, _
   ByVal samDesired As Long, _
   ByRef phkResult As Long _
) As Long

Public Declare Function RegQueryValueEx Lib "advapi32" _
   Alias "RegQueryValueExA" ( _
   ByVal HKey As Long, _
   ByVal lpValueName As String, _
   ByVal lpReserved As Long, _
   ByRef lpType As Long, _
   ByVal lpData As String, _
   ByRef lpcbData As Long _
) As Long

Public Declare Function ShowWindow Lib "user32.dll" ( _
   ByVal hwnd As Long, _
   ByVal nCmdShow As Long _
) As Long

Public Declare Function RegCloseKey Lib "advapi32" ( _
   ByVal HKey As Long _
) As Long

Public Declare Function SetWindowPos Lib "user32.dll" ( _
   ByVal hwnd As Long, _
   ByVal hWndInsertAfter As Long, _
   ByVal X As Long, _
   ByVal y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal wFlags As Long _
) As Long

Public Declare Function SetParent Lib "user32.dll" ( _
   ByVal hWndChild As Long, _
   ByVal hWndNewParent As Long _
) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long _
 ) As Long
 
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
   ByVal hwnd As Long, _
   ByVal nIndex As Long _
) As Long

Public Declare Function SetTextColor Lib "gdi32.dll" ( _
   ByVal hdc As Long, _
   ByVal crColor As Long _
) As Long

Public Declare Function CreateFont Lib "gdi32.dll" Alias "CreateFontA" ( _
   ByVal nHeight As Long, _
   ByVal nWidth As Long, _
   ByVal nEscapement As Long, _
   ByVal nOrientation As Long, _
   ByVal fnWeight As Long, _
   ByVal fdwItalic As Long, _
   ByVal fdwUnderline As Long, _
   ByVal fdwStrikeOut As Long, _
   ByVal fdwCharSet As Long, _
   ByVal fdwOutputPrecision As Long, _
   ByVal fdwClipPrecision As Long, _
   ByVal fdwQuality As Long, _
   ByVal fdwPitchAndFamily As Long, _
   ByVal lpszFace As String _
) As Long

Public Declare Function SelectObject Lib "gdi32" ( _
   ByVal hdc As Long, _
   ByVal hObject As Long _
) As Long

Public Declare Function DeleteObject Lib "gdi32" ( _
   ByVal hObject As Long _
) As Long

Public Declare Function TextOut Lib "gdi32" _
   Alias "TextOutA" ( _
   ByVal hdc As Long, _
   ByVal X As Long, _
   ByVal y As Long, _
   ByVal lpString As String, _
   ByVal nCount As Long _
) As Long

Public Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long

Public Function GetKeyValue(ByVal MainKey As VbHKey, ByVal SubKey As String, ByVal value As String) As String
   
   Dim RetVal As Long
   Dim HKey As Long
   Dim TmpSNum As String * 255
   
   RetVal = RegOpenKeyEx(MainKey, SubKey, 0&, KEY_READ, HKey)
   
   If RetVal <> 0 Then
      GetKeyValue = "Can't open the registry."
      Exit Function
   End If
   
   RetVal = RegQueryValueEx(HKey, value, 0, REG_SZ, ByVal TmpSNum, Len(TmpSNum))
    
   If RetVal <> 0 Then
      GetKeyValue = "Can't read or find the registry."
      Exit Function
   End If
   
   GetKeyValue = Left$(TmpSNum, InStr(1, TmpSNum, vbNullChar) - 1)
   
   RetVal = RegCloseKey(HKey)
   
End Function
