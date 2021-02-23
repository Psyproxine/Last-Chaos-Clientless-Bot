Attribute VB_Name = "modSubClass"
Option Explicit

Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" ( _
   ByVal hWnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long _
) As Long

Declare Function CallWindowProc Lib "user32" _
   Alias "CallWindowProcA" ( _
   ByVal lpPrevWndFunc As Long, _
   ByVal hWnd As Long, _
   ByVal msg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long _
) As Long

Public Const GWL_WNDPROC As Long = (-4)
Public Const WM_CAPTION_CHANGED As Long = &H7C

Public m_SubClassList As List
Public m_Perspective As ucPerspective
Public m_ListenCaptionChanged As Boolean

Private Function isIDE() As Boolean

   On Error GoTo ERROR_HANDLE
   
   Debug.Print 1 / 0
   
   Exit Function
   
ERROR_HANDLE:

   isIDE = True

End Function

' Hook a window by its handle.
'
' @lngHwnd The window handle
Public Sub Hook(ByVal lngHwnd As Long)
   
   Dim isHooked As Boolean
   Dim PrevProc As Long
   
   ' If isIDE Then Exit Sub
   
   isHooked = Not m_SubClassList.IsEmpty
   
   If isHooked Then
      isHooked = m_SubClassList.Contains("View_" & lngHwnd)
   End If
   
   If Not isHooked Then
      
      PrevProc = SetWindowLong(lngHwnd, GWL_WNDPROC, AddressOf WindowProc)
   
      m_SubClassList.Add "View_" & lngHwnd, PrevProc
   
   End If
   
End Sub

' Unhook a window by its handle.
'
' @lngHwnd The window handle
Public Sub UnHook(ByVal lngHwnd As Long)
    
    Dim PrevProc As Long
    
    If m_SubClassList Is Nothing Then Exit Sub
    
    If Not m_SubClassList.IsEmpty Then
       
       PrevProc = m_SubClassList.Item("View_" & lngHwnd)
    
       If PrevProc = 0 Then
          PrevProc = m_SubClassList.Item(1)
       End If
    
       If PrevProc > 0 Then
          SetWindowLong lngHwnd, GWL_WNDPROC, PrevProc
    
          m_SubClassList.Remove "View_" & lngHwnd

       End If
    
    End If
    
End Sub

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    On Error Resume Next
    
    Dim i As Long
    Dim l_View As View
    Dim l_Folder As Folder
    Dim PrevProc As Long
    Dim FolderRect As RECT
    
    If m_SubClassList Is Nothing Then Exit Function
    PrevProc = m_SubClassList.Item("View_" & hWnd)
    
    
    ' TODO workaround if window caption changed
    '      -> window style was reset
    If uMsg <> WM_COMMAND Then
       WindowProc = CallWindowProc(PrevProc, hWnd, uMsg, wParam, lParam)
    End If
    
    Select Case uMsg
   
       Case WM_COMMAND
            
            ' TODO workaround if window caption changed
            '      -> window style was reset
            ' Debug.Print "WM_COMMAND"
        
       Case WM_MOUSEACTIVATE
            
            If Not m_Perspective Is Nothing Then
               
               ' Show the activated view with the window handle
               With m_Views
                  For i = 0 To .Count
                     Set l_View = .Item(i)
          
                     If l_View.View.hWnd = hWnd Then
                        m_Perspective.ShowView l_View.ViewId
                        GoTo Finally
                     End If
                  Next i
               End With
                  
            End If
        
        ' View or Editor caption has changed
        Case WM_CAPTION_CHANGED

            If Not m_Perspective Is Nothing And _
               m_ListenCaptionChanged And _
               wParam = -20 Then

               ' Get view id of the editor by the window handle
               With m_Editors
                  For i = 0 To .Count
                     Set l_View = .Item(i)

                     If l_View.View.hWnd = hWnd Then
                        Exit For
                     Else
                        Set l_View = Nothing
                     End If
                  Next i
               End With

               ' Get view id of the view by the window handle
               With m_Views
                  For i = 0 To .Count
                     Set l_View = .Item(i)

                     If l_View.View.hWnd = hWnd Then
                        Exit For
                     Else
                        Set l_View = Nothing
                     End If
                  Next i
               End With

               ' Change capiton and reset style
               If Not l_View Is Nothing Then
                  m_ListenCaptionChanged = False

                  'Dim l_ucFolder As Form
                  Dim l_ucFolder As ucFolder
                  Dim l_ViewForm As Form
                  Set l_ViewForm = l_View.View
                  Set l_ucFolder = l_View.View.Parent
                  
                  SetWindowStyle hWnd, VbNone
                  SetWindowPos hWnd, 0&, 1, 1, 1, 1, 0&
                  
                  m_Perspective.RefreshView l_View.ViewId
                  
                 'm_Perspective.Refresh

                  GoTo Finally
               End If

            End If
    End Select
    
Finally:
   
   m_ListenCaptionChanged = True
   Set l_View = Nothing
   Set l_Folder = Nothing
    
End Function

