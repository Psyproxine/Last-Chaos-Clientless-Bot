VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MagneticWnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Dieser Quellcode stammt von http://www.ActiveVB.de
' und kann frei verwendet werden. Für eventuelle Schäden
' wird nicht gehaftet.
'
' Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
' Ansonsten viel Spaß und Erfolg mit diesem Quellcode!
'

'========================================================================================
' Class:         cMagneticWnd.cls
' Author:        Carles P.V. - ©2004 (*)
' Dependencies:
' Last revision: 2004.11.30
' Version:       1.0.8
'----------------------------------------------------------------------------------------
'
' (*) 1. Code based on original post by Benjamin Wilger:
'
'        Magnetic Forms
'        http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=27489&lngWId=1
'
'     2. Self-Subclassing UserControl/Class template (IDE safe) by Paul Caton:
'
'        Self-subclassing Controls/Forms - NO dependencies (v1.1.0010 2004.10.07)
'        http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'----------------------------------------------------------------------------------------
'
' History:
'
'     1.0.0: First release.
'
'     1.0.1: Use of DeferWindowPos() instead of MoveWindow.
'            Better in case of moving/sizing multiple windows simultaneously.
'            Thanks to jeremyxtz for suggestion.
'
'     1.0.2: Hereditary glueing.
'
'     1.0.3: - Removed 'RemoveWindow()' method.
'              Now, class process WM_DESTROY message and automatically removes window.
'            - Glueing checked in AddWindow().
'
'     1.0.4: Fixed: incorrect checking of 'hereditary glueing'.
'            I hope it's working fine now! Sorry.
'
'     1.0.5: Final update, I hope.
'            Added: hereditary magnetism (magnetism is also working for child windows).
'            I think that WinAmp's *behaviour* is now fully emulated :-)
'
'     1.0.6: Never say final update:
'            Added CheckGlueing() method. Call in case repositioning manually a window
'            and want to enable/check (glue) again, if any. This was only checked for
'            first time when new window added to collection.
'            Thanks to Gandolf_The_GUI for info.
'
'     1.0.7: Returning to manual destroying window (W9x problems)
'
'     1.0.8: - Added processing of WM_SYSCOMMAND and WM_COMMAND:
'              1. When window *state* is changed from 'system menu' or caption buttons.
'              2. When window *state* is changed *externaly*.
'              Thanks to LaVolpe for suggesting solution.
'            - Added checking for maximized windows: At time to extract rectangles,
'              maximized windows will take work area rectangle. This avoids edge
'              offset that causes real window rectangle to go out of screen (work) area.
'----------------------------------------------------------------------------------------

Option Explicit

'========================================================================================
' Subclasser declarations
'========================================================================================

Private Enum eMsgWhen
    [MSG_AFTER] = 1                                  'Message calls back after the original (previous) WndProc
    [MSG_BEFORE] = 2                                 'Message calls back before the original (previous) WndProc
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE 'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES     As Long = -1          'All messages added or deleted
Private Const CODE_LEN         As Long = 197         'Length of the machine code in bytes
Private Const GWL_WNDPROC      As Long = -4          'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04         As Long = 88          'Table B (before) address patch offset
Private Const PATCH_05         As Long = 93          'Table B (before) entry count patch offset
Private Const PATCH_08         As Long = 132         'Table A (after) address patch offset
Private Const PATCH_09         As Long = 137         'Table A (after) entry count patch offset

Private Type tSubData                                'Subclass data type
    hWnd                       As Long               'Handle of the window being subclassed
    nAddrSub                   As Long               'The address of our new WndProc (allocated memory).
    nAddrOrig                  As Long               'The address of the pre-existing WndProc
    nMsgCntA                   As Long               'Msg after table entry count
    nMsgCntB                   As Long               'Msg before table entry count
    aMsgTblA()                 As Long               'Msg after table array
    aMsgTblB()                 As Long               'Msg Before table array
End Type

Private sc_aSubData()          As tSubData           'Subclass data array
Private sc_aBuf(1 To CODE_LEN) As Byte               'Code buffer byte array
Private sc_pCWP                As Long               'Address of the CallWindowsProc
Private sc_pEbMode             As Long               'Address of the EbMode IDE break/stop/running function
Private sc_pSWL                As Long               'Address of the SetWindowsLong function
  
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

'========================================================================================
' cMagneticWnd
'========================================================================================

'-- API

Private Type PointAPI
    X1 As Long
    Y1 As Long
End Type

Private Type RECT2
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
End Type

Private Const SPI_GETWORKAREA  As Long = 48

Private Const WM_SIZING        As Long = &H214
Private Const WM_MOVING        As Long = &H216
Private Const WM_ENTERSIZEMOVE As Long = &H231
Private Const WM_EXITSIZEMOVE  As Long = &H232
Private Const WM_SYSCOMMAND    As Long = &H112
Private Const WM_COMMAND       As Long = &H111

Private Const WMSZ_LEFT        As Long = 1
Private Const WMSZ_RIGHT       As Long = 2
Private Const WMSZ_TOP         As Long = 3
Private Const WMSZ_TOPLEFT     As Long = 4
Private Const WMSZ_TOPRIGHT    As Long = 5
Private Const WMSZ_BOTTOM      As Long = 6
Private Const WMSZ_BOTTOMLEFT  As Long = 7
Private Const WMSZ_BOTTOMRIGHT As Long = 8

Private Const SC_MINIMIZE      As Long = &HF020&
Private Const SC_RESTORE       As Long = &HF120&

Private Const SWP_NOSIZE       As Long = &H1
Private Const SWP_NOZORDER     As Long = &H4
Private Const SWP_NOACTIVATE   As Long = &H10

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BeginDeferWindowPos Lib "user32" (ByVal nNumWindows As Long) As Long
Private Declare Function DeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long, ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT2) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT2, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function UnionRect Lib "user32" (lpDestRect As RECT2, lpSrc1Rect As RECT2, lpSrc2Rect As RECT2) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

'-- Private types:

Private Type WND_INFO
    hWnd       As Long
    hWndParent As Long
    Glue       As Boolean
End Type

'-- Private constants:

Private Const LB_RECT As Long = 16

'-- Private variables:

Private m_uWndInfo()  As WND_INFO
Private m_lWndCount   As Long
Private m_rcWnd()     As RECT2
Private m_ptAnchor    As PointAPI
Private m_ptOffset    As PointAPI
Private m_ptCurr      As PointAPI
Private m_ptLast      As PointAPI

'-- Property variables:

Private m_lSnapWidth As Long

'//

Private Sub Class_Initialize()
    
    '-- Default snap width
    m_lSnapWidth = 10
    
    '-- Initialize array (handled windows info)
    ReDim m_uWndInfo(0) As WND_INFO
    m_lWndCount = 0
End Sub

Private Sub Class_Terminate()
    
    '-- Stop subclassing
    If (m_lWndCount) Then
        Call Subclass_StopAll
    End If
End Sub



'========================================================================================
' Subclass handler: MUST be the first Public routine in this file.
'                   That includes public properties also.
'========================================================================================

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
'
'Parameters:
'   bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
'   bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
'   lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
'   lng_hWnd - The window handle
'   uMsg     - The message number
'   wParam   - Message related data
'   lParam   - Message related data
'
'Notes:
'   If you really know what you're doing, it's possible to change the values of the
'   hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
'   values get passed to the default handler.. and optionaly, the 'after' callback
  
  Dim rcWnd As RECT2
  Dim lc    As Long
  
    Select Case uMsg
        
        '-- Size/Move starting
        Case WM_ENTERSIZEMOVE
            
            '-- Get Desktop area (as first rectangle)
            Call SystemParametersInfo(SPI_GETWORKAREA, 0, m_rcWnd(0), 0)
            
            '-- Get rectangles of all handled windows
            For lc = 1 To m_lWndCount
                
                '-- Window maximized ?
                If (IsZoomed(m_uWndInfo(lc).hWnd)) Then
                    '-- Take work are rectangle
                    Call CopyMemory(m_rcWnd(lc), m_rcWnd(0), LB_RECT)
                  Else
                    '-- Get window rectangle
                    Call GetWindowRect(m_uWndInfo(lc).hWnd, m_rcWnd(lc))
                End If
                
                '-- Is it our current window ?
                If (m_uWndInfo(lc).hWnd = lng_hWnd) Then
                    '-- Get anchor-offset
                    Call GetCursorPos(m_ptAnchor)
                    Call GetCursorPos(m_ptLast)
                    m_ptOffset.X1 = m_rcWnd(lc).X1 - m_ptLast.X1
                    m_ptOffset.Y1 = m_rcWnd(lc).Y1 - m_ptLast.Y1
                End If
            Next lc
        
        '-- Sizing
        Case WM_SIZING
            
            Call CopyMemory(rcWnd, ByVal lParam, LB_RECT)
            Call pvSizeRect(lng_hWnd, rcWnd, wParam)
            Call CopyMemory(ByVal lParam, rcWnd, LB_RECT)
            
            bHandled = True
            lReturn = 1
        
        '-- Moving
        Case WM_MOVING
            
            Call CopyMemory(rcWnd, ByVal lParam, LB_RECT)
            Call pvMoveRect(lng_hWnd, rcWnd)
            Call CopyMemory(ByVal lParam, rcWnd, LB_RECT)
            
            bHandled = True
            lReturn = 1
        
        '-- Size/Move finishing
        Case WM_EXITSIZEMOVE
            
            Call pvCheckGlueing
            
        '-- Special case: *menu* call
        Case WM_SYSCOMMAND
            
            If (wParam = SC_MINIMIZE Or wParam = SC_RESTORE) Then
                Call pvCheckGlueing
            End If
        
        '-- Special case: *control* call
        Case WM_COMMAND
            
            Call pvCheckGlueing
    End Select
End Sub



'========================================================================================
' Methods
'========================================================================================

Public Function AddWindow(ByVal hWnd As Long, Optional ByVal hWndParent As Long = 0) As Boolean

  Dim lc As Long
    
    '-- Already in collection ?
    For lc = 1 To m_lWndCount
        If (hWnd = m_uWndInfo(lc).hWnd) Then Exit Function
    Next lc
    
    '-- Validate windows
    If (IsWindow(hWnd) And (IsWindow(hWndParent) Or hWndParent = 0)) Then
        
        '-- Increase count
        m_lWndCount = m_lWndCount + 1
        '-- Resize arrays
        ReDim Preserve m_uWndInfo(0 To m_lWndCount)
        ReDim Preserve m_rcWnd(0 To m_lWndCount)
        
        '-- Add info
        With m_uWndInfo(m_lWndCount)
            .hWnd = hWnd
            .hWndParent = hWndParent
        End With
        
        '-- Check glueing for first time
        Call pvCheckGlueing
        
        '-- Start subclassing
        Call Subclass_Start(hWnd)
        Call Subclass_AddMsg(hWnd, WM_ENTERSIZEMOVE)
        Call Subclass_AddMsg(hWnd, WM_SIZING, [MSG_BEFORE])
        Call Subclass_AddMsg(hWnd, WM_MOVING, [MSG_BEFORE])
        Call Subclass_AddMsg(hWnd, WM_EXITSIZEMOVE)
        Call Subclass_AddMsg(hWnd, WM_SYSCOMMAND)
        Call Subclass_AddMsg(hWnd, WM_COMMAND)
        
        '-- Success
        AddWindow = True
    End If
End Function

Public Function RemoveWindow(ByVal hWnd As Long) As Boolean

  Dim lc1 As Long
  Dim lc2 As Long

    For lc1 = 1 To m_lWndCount
        
        If (hWnd = m_uWndInfo(lc1).hWnd) Then
            
            '-- Move down
            For lc2 = lc1 To m_lWndCount - 1
                m_uWndInfo(lc2) = m_uWndInfo(lc2 + 1)
            Next lc2
            
            '-- Resize arrays
            m_lWndCount = m_lWndCount - 1
            ReDim Preserve m_uWndInfo(m_lWndCount)
            ReDim Preserve m_rcWnd(m_lWndCount)
            
            '-- Remove parent relationships
            For lc2 = 1 To m_lWndCount
                If (m_uWndInfo(lc2).hWndParent = hWnd) Then
                    m_uWndInfo(lc2).hWndParent = 0
                End If
            Next lc2
            
            '-- Stop subclassing / verify connections
            Call Subclass_Stop(hWnd)
            Call pvCheckGlueing
            
            '-- Success
            RemoveWindow = True
            Exit For
        End If
    Next lc1
End Function

Public Sub CheckGlueing()
        
    '-- Check ALL windows for possible new *connections*.
    Call pvCheckGlueing
End Sub



'========================================================================================
' Properties
'========================================================================================

Public Property Get SnapWidth() As Long
    SnapWidth = m_lSnapWidth
End Property

Public Property Let SnapWidth(ByVal New_SnapWidth As Long)
    m_lSnapWidth = New_SnapWidth
End Property



'========================================================================================
' Private
'========================================================================================

Private Sub pvSizeRect(ByVal hWnd As Long, rcWnd As RECT2, ByVal lfEdge As Long)
    
  Dim rcTmp As RECT2
  Dim lc    As Long
    
    '-- Get a copy
    Call CopyMemory(rcTmp, rcWnd, LB_RECT)
    
    '-- Check all windows
    For lc = 0 To m_lWndCount
        
        With m_rcWnd(lc)
            
            '-- Avoid current window
            If (m_uWndInfo(lc).hWnd <> hWnd) Then
                
                '-- X magnetism
                If (rcWnd.Y1 < .Y2 + m_lSnapWidth And rcWnd.Y2 > .Y1 - m_lSnapWidth) Then
                    
                    Select Case lfEdge
                        
                      Case WMSZ_LEFT, WMSZ_TOPLEFT, WMSZ_BOTTOMLEFT
                    
                        Select Case True
                          Case Abs(rcTmp.X1 - .X1) < m_lSnapWidth: rcWnd.X1 = .X1
                          Case Abs(rcTmp.X1 - .X2) < m_lSnapWidth: rcWnd.X1 = .X2
                        End Select
                
                      Case WMSZ_RIGHT, WMSZ_TOPRIGHT, WMSZ_BOTTOMRIGHT
                        
                        Select Case True
                          Case Abs(rcTmp.X2 - .X1) < m_lSnapWidth: rcWnd.X2 = .X1
                          Case Abs(rcTmp.X2 - .X2) < m_lSnapWidth: rcWnd.X2 = .X2
                        End Select
                    End Select
                End If
                
                '-- Y magnetism
                If (rcWnd.X1 < .X2 + m_lSnapWidth And rcWnd.X2 > .X1 - m_lSnapWidth) Then
                    
                    Select Case lfEdge
                        
                      Case WMSZ_TOP, WMSZ_TOPLEFT, WMSZ_TOPRIGHT
                        
                        Select Case True
                          Case Abs(rcTmp.Y1 - .Y1) < m_lSnapWidth: rcWnd.Y1 = .Y1
                          Case Abs(rcTmp.Y1 - .Y2) < m_lSnapWidth: rcWnd.Y1 = .Y2
                        End Select
                    
                      Case WMSZ_BOTTOM, WMSZ_BOTTOMLEFT, WMSZ_BOTTOMRIGHT
                        
                        Select Case True
                          Case Abs(rcTmp.Y2 - .Y1) < m_lSnapWidth: rcWnd.Y2 = .Y1
                          Case Abs(rcTmp.Y2 - .Y2) < m_lSnapWidth: rcWnd.Y2 = .Y2
                        End Select
                    End Select
                End If
            End If
        End With
    Next lc
End Sub

Private Sub pvMoveRect(ByVal hWnd As Long, rcWnd As RECT2)
    
  Dim lc1   As Long
  Dim lc2   As Long
  Dim lWId  As Long
  Dim rcTmp As RECT2
  Dim lOffx As Long
  Dim lOffy As Long
  Dim hDWP  As Long
    
    '== Get current cursor position
    
    Call GetCursorPos(m_ptCurr)
    
    '== Check magnetism for current window
    
    '-- 'Move' current window
    Call OffsetRect(rcWnd, (m_ptCurr.X1 - rcWnd.X1) + m_ptOffset.X1, 0)
    Call OffsetRect(rcWnd, 0, (m_ptCurr.Y1 - rcWnd.Y1) + m_ptOffset.Y1)
    
    '-- Check all windows
    For lc1 = 0 To m_lWndCount
        
        '-- Avoid current window
        If (m_uWndInfo(lc1).hWnd <> hWnd) Then
                
            '-- Avoid child windows
            If (m_uWndInfo(lc1).Glue = False Or m_uWndInfo(lc1).hWndParent <> hWnd) Then
                    
                With m_rcWnd(lc1)
                
                    '-- X magnetism
                    If (rcWnd.Y1 < .Y2 + m_lSnapWidth And rcWnd.Y2 > .Y1 - m_lSnapWidth) Then
                    
                        Select Case True
                          Case Abs(rcWnd.X1 - .X1) < m_lSnapWidth: lOffx = .X1 - rcWnd.X1
                          Case Abs(rcWnd.X1 - .X2) < m_lSnapWidth: lOffx = .X2 - rcWnd.X1
                          Case Abs(rcWnd.X2 - .X1) < m_lSnapWidth: lOffx = .X1 - rcWnd.X2
                          Case Abs(rcWnd.X2 - .X2) < m_lSnapWidth: lOffx = .X2 - rcWnd.X2
                        End Select
                    End If
                    
                    '-- Y magnetism
                    If (rcWnd.X1 < .X2 + m_lSnapWidth And rcWnd.X2 > .X1 - m_lSnapWidth) Then
                    
                        Select Case True
                          Case Abs(rcWnd.Y1 - .Y1) < m_lSnapWidth: lOffy = .Y1 - rcWnd.Y1
                          Case Abs(rcWnd.Y1 - .Y2) < m_lSnapWidth: lOffy = .Y2 - rcWnd.Y1
                          Case Abs(rcWnd.Y2 - .Y1) < m_lSnapWidth: lOffy = .Y1 - rcWnd.Y2
                          Case Abs(rcWnd.Y2 - .Y2) < m_lSnapWidth: lOffy = .Y2 - rcWnd.Y2
                        End Select
                    End If
                End With
            End If
        End If
    Next lc1
    
    '== Check magnetism for child windows
    
    For lc1 = 1 To m_lWndCount
        
        '-- Child and connected window ?
        If (m_uWndInfo(lc1).Glue And m_uWndInfo(lc1).hWndParent = hWnd) Then
            
            '-- 'Move' child window
            Call CopyMemory(rcTmp, m_rcWnd(lc1), LB_RECT)
            Call OffsetRect(rcTmp, m_ptCurr.X1 - m_ptAnchor.X1, 0)
            Call OffsetRect(rcTmp, 0, m_ptCurr.Y1 - m_ptAnchor.Y1)
            
            For lc2 = 0 To m_lWndCount
                                        
                If (lc1 <> lc2) Then
                    
                    '-- Avoid child windows
                    If (m_uWndInfo(lc2).Glue = False And m_uWndInfo(lc2).hWnd <> hWnd) Then
                    
                        With m_rcWnd(lc2)
                    
                            '-- X magnetism
                            If (rcTmp.Y1 < .Y2 + m_lSnapWidth And rcTmp.Y2 > .Y1 - m_lSnapWidth) Then
                                
                                Select Case True
                                  Case Abs(rcTmp.X1 - .X1) < m_lSnapWidth: lOffx = .X1 - rcTmp.X1
                                  Case Abs(rcTmp.X1 - .X2) < m_lSnapWidth: lOffx = .X2 - rcTmp.X1
                                  Case Abs(rcTmp.X2 - .X1) < m_lSnapWidth: lOffx = .X1 - rcTmp.X2
                                  Case Abs(rcTmp.X2 - .X2) < m_lSnapWidth: lOffx = .X2 - rcTmp.X2
                                End Select
                            End If
                            
                            '-- Y magnetism
                            If (rcTmp.X1 < .X2 + m_lSnapWidth And rcTmp.X2 > .X1 - m_lSnapWidth) Then
                            
                                Select Case True
                                  Case Abs(rcTmp.Y1 - .Y1) < m_lSnapWidth: lOffy = .Y1 - rcTmp.Y1
                                  Case Abs(rcTmp.Y1 - .Y2) < m_lSnapWidth: lOffy = .Y2 - rcTmp.Y1
                                  Case Abs(rcTmp.Y2 - .Y1) < m_lSnapWidth: lOffy = .Y1 - rcTmp.Y2
                                  Case Abs(rcTmp.Y2 - .Y2) < m_lSnapWidth: lOffy = .Y2 - rcTmp.Y2
                                End Select
                            End If
                        End With
                    End If
                End If
            Next lc2
        End If
    Next lc1
    
    '== Apply offsets
    
    Call OffsetRect(rcWnd, lOffx, lOffy)
    
    '== Glueing (move child windows, if any)
    
    hDWP = BeginDeferWindowPos(1)
    
    For lc1 = 1 To m_lWndCount
        With m_uWndInfo(lc1)
            '-- Is parent our current window ?
            If (.hWndParent = hWnd And .Glue) Then
                '-- Move 'child' window
                lWId = pvWndGetInfoIndex(hWnd)
                With m_rcWnd(lc1)
                    Call DeferWindowPos(hDWP, m_uWndInfo(lc1).hWnd, 0, .X1 - (m_rcWnd(lWId).X1 - rcWnd.X1), .Y1 - (m_rcWnd(lWId).Y1 - rcWnd.Y1), 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOZORDER)
                End With
            End If
        End With
    Next lc1
    
    Call EndDeferWindowPos(hDWP)
    
    '== Store last cursor position
    
    m_ptLast = m_ptCurr
End Sub

Private Sub pvCheckGlueing()
    
  Dim lcMain As Long
  Dim lc1    As Long
  Dim lc2    As Long
  Dim lWId   As Long
    
    '-- Get all windows rectangles / Reset glueing
    For lc1 = 1 To m_lWndCount
        
        Call GetWindowRect(m_uWndInfo(lc1).hWnd, m_rcWnd(lc1))
        m_uWndInfo(lc1).Glue = False
    Next lc1
    
    '-- Check direct connection
    For lc1 = 1 To m_lWndCount
        
        If (m_uWndInfo(lc1).hWndParent) Then
        
            '-- Get parent window info index
            lWId = pvWndParentGetInfoIndex(m_uWndInfo(lc1).hWndParent)
            '-- Connected ?
            m_uWndInfo(lc1).Glue = pvWndsConnected(m_rcWnd(lWId), m_rcWnd(lc1))
        End If
    Next lc1
    
    '-- Check indirect connection
    For lcMain = 1 To m_lWndCount
        
        For lc1 = 1 To m_lWndCount
            
            If (m_uWndInfo(lc1).Glue) Then
                
                For lc2 = 1 To m_lWndCount
                
                    If (lc1 <> lc2) Then
                    
                        If (m_uWndInfo(lc1).hWndParent = m_uWndInfo(lc2).hWndParent) Then
                            '-- Connected ?
                            If (m_uWndInfo(lc2).Glue = False) Then
                                m_uWndInfo(lc2).Glue = pvWndsConnected(m_rcWnd(lc1), m_rcWnd(lc2))
                            End If
                        End If
                    End If
                Next lc2
            End If
        Next lc1
    Next lcMain
End Sub

Private Function pvWndsConnected(rcWnd1 As RECT2, rcWnd2 As RECT2) As Boolean
    
  Dim rcUnion As RECT2
  
    '-- Calc. union rectangle of windows
    Call UnionRect(rcUnion, rcWnd1, rcWnd2)
    
    '-- Bounding glue-rectangle
    If ((rcUnion.X2 - rcUnion.X1) <= (rcWnd1.X2 - rcWnd1.X1) + (rcWnd2.X2 - rcWnd2.X1) And _
        (rcUnion.Y2 - rcUnion.Y1) <= (rcWnd1.Y2 - rcWnd1.Y1) + (rcWnd2.Y2 - rcWnd2.Y1) _
         ) Then
        
        '-- Edge coincidences ?
        If (rcWnd1.X1 = rcWnd2.X1 Or rcWnd1.X1 = rcWnd2.X2 Or _
            rcWnd1.X2 = rcWnd2.X1 Or rcWnd1.X2 = rcWnd2.X2 Or _
            rcWnd1.Y1 = rcWnd2.Y1 Or rcWnd1.Y1 = rcWnd2.Y2 Or _
            rcWnd1.Y2 = rcWnd2.Y1 Or rcWnd1.Y2 = rcWnd2.Y2 _
            ) Then
            
            pvWndsConnected = True
        End If
    End If
End Function

Private Function pvWndGetInfoIndex(ByVal hWnd As Long) As Long
    
  Dim lc As Long
    
    For lc = 1 To m_lWndCount
        If (m_uWndInfo(lc).hWnd = hWnd) Then
            pvWndGetInfoIndex = lc
            Exit For
        End If
    Next lc
End Function

Private Function pvWndParentGetInfoIndex(ByVal hWndParent As Long) As Long
    
  Dim lc As Long
    
    For lc = 1 To m_lWndCount
        If (m_uWndInfo(lc).hWnd = hWndParent) Then
            pvWndParentGetInfoIndex = lc
            Exit For
        End If
    Next lc
End Function



'========================================================================================
' Subclass code - The programmer may call any of the following Subclass_??? routines
'========================================================================================

Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
'Parameters:
'   lng_hWnd - The handle of the window for which the uMsg is to be added to the callback table
'   uMsg     - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
'   When     - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  
    With sc_aSubData(zIdx(lng_hWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Delete a message from the table of those that will invoke a callback.
'Parameters:
'   lng_hWnd - The handle of the window for which the uMsg is to be removed from the callback table
'   uMsg     - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
'   When     - Whether the msg is to be removed from the before, after or both callback tables
  
    With sc_aSubData(zIdx(lng_hWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Function Subclass_InIDE() As Boolean
'Return whether we're running in the IDE.
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Start subclassing the passed window handle
'Parameters:
'   lng_hWnd - The handle of the window to be subclassed
'Returns;
'   The sc_aSubData() index

  Dim i                        As Long                       'Loop index
  Dim J                        As Long                       'Loop index
  Dim nSubIdx                  As Long                       'Subclass data index
  Dim sSubCode                 As String                     'Subclass code string
  
  Const GMEM_FIXED             As Long = 0                   'Fixed memory GlobalAlloc flag
  Const PAGE_EXECUTE_READWRITE As Long = &H40&               'Allow memory to execute without violating XP SP2 Data Execution Prevention
  Const PATCH_01               As Long = 18                  'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02               As Long = 68                  'Address of the previous WndProc
  Const PATCH_03               As Long = 78                  'Relative address of SetWindowsLong
  Const PATCH_06               As Long = 116                 'Address of the previous WndProc
  Const PATCH_07               As Long = 121                 'Relative address of CallWindowProc
  Const PATCH_0A               As Long = 186                 'Address of the owner object
  Const FUNC_CWP               As String = "CallWindowProcA" 'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM               As String = "EbMode"          'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL               As String = "SetWindowLongA"  'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER               As String = "user32"          'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5               As String = "vba5"            'Location of the EbMode function if running VB5
  Const MOD_VBA6               As String = "vba6"            'Location of the EbMode function if running VB6

    'If it's the first time through here..
    If (sc_aBuf(1) = 0) Then

        'Build the hex pair subclass string
        sSubCode = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E32F09C978078B450CF2AF75248D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF501CC3"
    
        'Convert the string from hex pairs to bytes and store in the machine code buffer
        i = 1
        Do While J < CODE_LEN
            J = J + 1
            sc_aBuf(J) = CByte("&H" & Mid$(sSubCode, i, 2))                       'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                      'Next pair of hex characters
    
        'Get API function addresses
        If (Subclass_InIDE) Then                                                  'If we're running in the VB IDE
            sc_aBuf(16) = &H90                                                    'Patch the code buffer to enable the IDE state code
            sc_aBuf(17) = &H90                                                    'Patch the code buffer to enable the IDE state code
            sc_pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                            'Get the address of EbMode in vba6.dll
            If (sc_pEbMode = 0) Then                                              'Found?
                sc_pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                        'VB5 perhaps
            End If
        End If
    
        Call zPatchVal(VarPtr(sc_aBuf(1)), PATCH_0A, ObjPtr(Me))                  'Patch the address of this object instance into the static machine code buffer
    
        sc_pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                   'Get the address of the CallWindowsProc function
        sc_pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                   'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                     'Create the first sc_aSubData element
    
      Else
        nSubIdx = zIdx(lng_hWnd, True)
        If (nSubIdx = -1) Then                                                    'If an sc_aSubData element isn't being re-cycled
            nSubIdx = UBound(sc_aSubData()) + 1                                   'Calculate the next element
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                  'Create a new sc_aSubData element
        End If
    
        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                             'Allocate memory for the machine code WndProc
        Call VirtualProtect(ByVal .nAddrSub, CODE_LEN, PAGE_EXECUTE_READWRITE, i) 'Mark memory as executable
        Call RtlMoveMemory(ByVal .nAddrSub, sc_aBuf(1), CODE_LEN)                 'Copy the machine code from the static byte array to the code array in sc_aSubData
    
        .hWnd = lng_hWnd                                                          'Store the hWnd
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)                'Set our WndProc in place
    
        Call zPatchRel(.nAddrSub, PATCH_01, sc_pEbMode)                           'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                           'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, sc_pSWL)                              'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                           'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, sc_pCWP)                              'Patch the relative address of the CallWindowProc api function
    End With
End Function

Private Sub Subclass_StopAll()
'Stop all subclassing
  
  Dim i As Long
  
    i = UBound(sc_aSubData())                                                     'Get the upper bound of the subclass data array
    Do While i >= 0                                                               'Iterate through each element
        With sc_aSubData(i)
            If (.hWnd <> 0) Then                                                  'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hWnd)                                         'Subclass_Stop
            End If
        End With
    
        i = i - 1                                                                 'Next element
    Loop
End Sub

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Stop subclassing the passed window handle
'Parameters:
'   lng_hWnd - The handle of the window to stop being subclassed
  
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)                       'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                                    'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                                    'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                                'Release the machine code memory
        .hWnd = 0                                                                 'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                             'Clear the before table
        .nMsgCntA = 0                                                             'Clear the after table
        Erase .aMsgTblB                                                           'Erase the before table
        Erase .aMsgTblA                                                           'Erase the after table
    End With
End Sub

'----------------------------------------------------------------------------------------
'These z??? routines are exclusively called by the Subclass_??? routines.
'----------------------------------------------------------------------------------------

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
'Worker sub for Subclass_AddMsg
  
  Dim nEntry  As Long                                                             'Message table entry index
  Dim nOff1   As Long                                                             'Machine code buffer offset 1
  Dim nOff2   As Long                                                             'Machine code buffer offset 2
  
    If (uMsg = ALL_MESSAGES) Then                                                 'If all messages
        nMsgCnt = ALL_MESSAGES                                                    'Indicates that all messages will callback
      Else                                                                        'Else a specific message number
        Do While nEntry < nMsgCnt                                                 'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1
        
            If (aMsgTbl(nEntry) = 0) Then                                         'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                            'Re-use this entry
                Exit Sub                                                          'Bail
            ElseIf (aMsgTbl(nEntry) = uMsg) Then                                  'The msg is already in the table!
                Exit Sub                                                          'Bail
            End If
        Loop                                                                      'Next entry

        nMsgCnt = nMsgCnt + 1                                                     'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                              'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                   'Store the message number in the table
    End If

    If (When = eMsgWhen.MSG_BEFORE) Then                                          'If before
        nOff1 = PATCH_04                                                          'Offset to the Before table
        nOff2 = PATCH_05                                                          'Offset to the Before table entry count
      Else                                                                        'Else after
        nOff1 = PATCH_08                                                          'Offset to the After table
        nOff2 = PATCH_09                                                          'Offset to the After table entry count
    End If

    If (uMsg <> ALL_MESSAGES) Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                          'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                         'Patch the appropriate table entry count
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
'Return the memory address of the passed function in the passed dll
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                        'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
'Worker sub for Subclass_DelMsg
  
  Dim nEntry As Long
  
    If (uMsg = ALL_MESSAGES) Then                                                 'If deleting all messages
        nMsgCnt = 0                                                               'Message count is now zero
        If When = eMsgWhen.MSG_BEFORE Then                                        'If before
            nEntry = PATCH_05                                                     'Patch the before table message count location
          Else                                                                    'Else after
            nEntry = PATCH_09                                                     'Patch the after table message count location
        End If
        Call zPatchVal(nAddr, nEntry, 0)                                          'Patch the table message count to zero
      Else                                                                        'Else deleteting a specific message
        Do While nEntry < nMsgCnt                                                 'For each table entry
            nEntry = nEntry + 1
            If (aMsgTbl(nEntry) = uMsg) Then                                      'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                                               'Mark the table slot as available
                Exit Do                                                           'Bail
            End If
        Loop                                                                      'Next entry
    End If
End Sub

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the sc_aSubData() array index of the passed hWnd
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                            'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If (.hWnd = lng_hWnd) Then                                            'If the hWnd of this element is the one we're looking for
                If (Not bAdd) Then                                                'If we're searching not adding
                    Exit Function                                                 'Found
                End If
            ElseIf (.hWnd = 0) Then                                               'If this an element marked for reuse.
                If (bAdd) Then                                                    'If we're adding
                    Exit Function                                                 'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                                                           'Decrement the index
    Loop
  
    If (Not bAdd) Then
        Debug.Assert False                                                        'hWnd not found, programmer error
    End If

'If we exit here, we're returning -1, no freed elements were found
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
'Patch the machine code buffer at the indicated offset with the relative address to the target address.
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
'Patch the machine code buffer at the indicated offset with the passed value
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
'Worker function for Subclass_InIDE
    zSetTrue = True
    bValue = True
End Function

