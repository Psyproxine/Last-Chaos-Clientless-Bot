VERSION 5.00
Begin VB.UserControl ucPerspective 
   Alignable       =   -1  'True
   ClientHeight    =   8145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10695
   PropertyPages   =   "ucPerspective.ctx":0000
   ScaleHeight     =   8145
   ScaleWidth      =   10695
   ToolboxBitmap   =   "ucPerspective.ctx":0022
   Begin VB.Timer tmrDrag 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   240
      Top             =   240
   End
End
Attribute VB_Name = "ucPerspective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

' Enumeration of the different color schemes constants.
Public Enum vbColorScheme
   VbWindowsXPScheme = 0        ' Windows XP color scheme
   VbWindowsVistaScheme = 1     ' Windows Longhorn color scheme
   VbOffice2003Scheme = 2       ' Office 2003 color scheme
   VbEclipse3Scheme = 3         ' Eclipse 3.x scheme
   VbVS2005Scheme = 4           ' Visual Studio 2005 color scheme
   VbCustomScheme = 99          ' Custom color scheme
End Enum

' Enumeration of the different caption style constants.
Public Enum vbCaptionStyle
   vbHorizontalGradient = 0     ' Draws a horizontal gradient caption.
   vbVerticalGradient = 1       ' Draws a vertical gradient caption.
End Enum

Private m_ActivePerspectiveId As String ' id of the active perspective
Private m_ActiveFolderId As String ' id of the active folder
Private m_ActiveViewId As String ' id of the active view
Private m_Perspectives As List ' A list of all added perspectives
Private m_MaximizedFolderId As String
Private m_MainHwnd As Long

Private m_OldRelation As vbRelationship
Private m_OldRc As RECT
Private m_DropFolderId As String
Private m_MagneticWnd As New MagneticWnd

' Properties
Dim m_ColorScheme As vbColorScheme
Dim m_CaptionStyle As vbCaptionStyle

Dim m_ActiveCaptionForeColor As OLE_COLOR
Dim m_ActiveCaptionGradient1 As OLE_COLOR
Dim m_ActiveCaptionGradient2 As OLE_COLOR
Dim m_InactiveCaptionForeColor As OLE_COLOR
Dim m_InactiveCaptionGradient1 As OLE_COLOR
Dim m_InactiveCaptionGradient2 As OLE_COLOR

Dim m_FocusTabForeColor As OLE_COLOR
Dim m_FocusTabGradient1 As OLE_COLOR
Dim m_FocusTabGradient2 As OLE_COLOR
Dim m_FocusTabGradientAngle As Long
Dim m_ActiveTabForeColor As OLE_COLOR
Dim m_ActiveTabGradient1 As OLE_COLOR
Dim m_ActiveTabGradient2 As OLE_COLOR
Dim m_ActiveTabGradientAngle As Long

Dim m_BackColor As OLE_COLOR
Dim m_FrameColor As OLE_COLOR
Dim m_FrameWidth As OLE_COLOR
Dim m_EditorAreaBackColor As OLE_COLOR
Dim m_ViewCaptions As Boolean
Dim m_ViewCaptionIcons As Boolean
Dim m_Magnetic As Boolean

Private Const m_def_ColorScheme As Long = 0
Private Const m_def_CaptionStyle As Long = 0
Private Const m_def_Magnetic As Boolean = False

Private Const m_def_ActiveCaptionForeColor As Long = vbActiveTitleBarText
Private Const m_def_ActiveCaptionGradient1 As Long = vbActiveTitleBar
Private Const m_def_ActiveCaptionGradient2 As Long = vbActiveTitleBar
Private Const m_def_ActiveCaptionGradientAngle As Long = 180
Private Const m_def_InactiveCaptionForeColor As Long = vbInactiveTitleBarText
Private Const m_def_InactiveCaptionGradient1 As Long = vbInactiveTitleBar
Private Const m_def_InactiveCaptionGradient2 As Long = vbInactiveTitleBar
Private Const m_def_InactiveCaptionGradientAngle As Long = 180
Private Const m_def_FocusTabForeColor As Long = vbBlack
Private Const m_def_FocusTabGradient1 As Long = vbButtonFace
Private Const m_def_FocusTabGradient2 As Long = vbWhite
Private Const m_def_FocusTabGradientAngle As Long = 180
Private Const m_def_ActiveTabForeColor As Long = vbBlack
Private Const m_def_ActiveTabGradient1 As Long = vbButtonFace
Private Const m_def_ActiveTabGradient2 As Long = vbWhite
Private Const m_def_ActiveTabGradientAngle As Long = 180
Private Const m_def_BackColor As Long = vbButtonFace
Private Const m_def_FrameColor As Long = vbHighlight
Private Const m_def_FrameWidth As Long = 0
Private Const m_def_EditorAreaBackColor As Long = vbApplicationWorkspace
Private Const m_def_ViewCaptions As Boolean = True
Private Const m_def_ViewCaptionIcons As Boolean = True

Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

' This event occures if a view was activated.
' @ViewId The view id of the activated view.
Public Event ActivateView(ByVal ViewId As String)

' This event occures if a editor was activated.
' @Editor The activated editor.
Public Event ActivateEditor(ByRef Editor As Object)

' This event occures if a view was closed.
' @ViewId The view id of the closed view.
Public Event CloseView(ByVal ViewId As String, Cancel As Boolean)

' This event occures if a editor was closed.
' @Editor The closed editor.
Public Event CloseEditor(ByRef Editor As Object, Cancel As Boolean)

' This event occures if a editor was opend.
' @Editor The opend editor.
Public Event OpenEditor(ByRef Editor As Object)

' This event occures if a folder was restored.
' @FolderId The id of the restored folder.
Public Event RestoreView(ByVal ViewId As String)

' This event occures if a folder was maximized.
' @FolderId The id of the maximized folder.
Public Event MaximizeView(ByVal ViewId As String)

' This event occures if a perspective was opend.
' @PerspectiveId The id of the opened perspective.
Public Event OpenPerspective(ByVal PerspectiveId As String)

' This event occures if a perspective was closed.
' @PerspectiveId The id of the closed perspective.
Public Event ClosePerspective(ByVal PerspectiveId As String, Cancel As Boolean)

' Initialize the control and module variables.
Private Sub UserControl_Initialize()
   
   Set m_Perspectives = New List
   Set m_Scheme = New SchemeWinXP
   
   ' SubClass views / editors
   Set modSubClass.m_Perspective = Me
   Set modSubClass.m_SubClassList = New List
   
End Sub

' Terminate the control and all editors, views and module variables.
Private Sub UserControl_Terminate()
   
   On Error Resume Next
   
   Dim i As Long
   Dim l_Perspective As Perspective
   Dim l_ucFolder As ucFolder
   Dim l_View As View
         
   '-----------------------------------------
   ' Unload all editor views
   '-----------------------------------------
   Set l_Perspective = ActivePerspective
   
   If Not l_Perspective Is Nothing Then
   
      Set l_ucFolder = Controls.Item(l_Perspective.ID_EDITOR_AREA)
        
      If Not l_ucFolder Is Nothing Then
      
         With l_ucFolder.Views
            For i = 0 To .Count
               Set l_View = .Item(i)
               
               modSubClass.UnHook l_View.View.hwnd
                      
               Unload l_View.View
               Set l_View.View = Nothing
               
            Next i
         End With
         
      End If
      
   End If
   
   '-----------------------------------------
   ' Unhook all views
   '-----------------------------------------
   With m_Views
      For i = 0 To .Count
         Set l_View = .Item(i)
         modSubClass.UnHook l_View.View.hwnd
      Next i
   End With
   
   Set modSubClass.m_Perspective = Nothing
   Set modSubClass.m_SubClassList = Nothing
  
   Set l_Perspective = Nothing
   Set l_ucFolder = Nothing
   Set l_View = Nothing
   
   '-------------------------------------------
   ' Free module variables
   '-------------------------------------------
   Set m_Perspectives = Nothing
   Set m_Views = Nothing
   Set m_Scheme = Nothing
   
End Sub

' Refresh the control if the control was resized.
Private Sub UserControl_Resize()
   Refresh
End Sub

' Set the window handle of the main window (the window contains the perspective control).
' The window handle will be used to set as the owner of the floating windows.
'
' @New_MainHwnd The window handle of the main window.
Public Property Let MainHwnd(ByVal New_MainHwnd As Long)
   
   m_MainHwnd = New_MainHwnd
   
   ' Set the main window
   Call m_MagneticWnd.AddWindow(New_MainHwnd)
   
End Property

' Returns the window handle of the main window.
'
' @MainHwnd The window handle of the main window.
Private Property Get MainHwnd() As Long
   MainHwnd = m_MainHwnd
End Property

' Use this function to set a new color scheme. Just create a new class which implements
' the IScheme interface and describes the controls colors.
'
' @New_Scheme A class implementing the IScheme interface.
Public Sub SetScheme(New_Scheme As IScheme)
   Set m_Scheme = New_Scheme
   Refresh
End Sub

' Returns an unique folder id.
'
' @GetUniqueFolderId An unique folder id.
Private Function GetUniqueFolderId() As String
   
   Dim c As Control
   Static Idx As Long
   
   Do
      GetUniqueFolderId = "vbPerspective_Folder_" & Idx
   
      Idx = Idx + 1
      
      On Error Resume Next
      Set c = Nothing
      Set c = Controls.Item(GetUniqueFolderId)
      On Error GoTo 0
      
   Loop While (Not c Is Nothing)
         
End Function

' Drop the current active view to a new folder. The functions creates new folders depending on
' the drop relationship.
'
' @DropFolderId The folder id to drop the current active view.
' @DropRelationship The relationship to the drop folder id.
Private Sub DropActiveViewOn(ByVal DropFolderId As String, ByVal DropRelation As vbRelationship)
   
   On Error Resume Next
   
   Dim w As Long
   Dim h As Long
   
   Dim l_Perspective As Perspective
   Dim l_DragFolder As Folder
   Dim l_DropFolder As Folder
   Dim l_ucDragFolder As ucFolder
   Dim l_ucDropFolder As ucFolder
   
   Dim l_FolderRect As RECT
   Dim l_DragViewId As String
   Dim l_CursorPos As POINTAPI
   Dim l_ViewAdded As Boolean
   
   Dim frm As frmToolWin
   
   Set l_Perspective = ActivePerspective
   
   If Len(ActiveViewId) = 0 Then
      Exit Sub
   Else
      l_DragViewId = ActiveViewId
   End If
      
   If Len(ActiveFolderId) > 0 Then
      Set l_DragFolder = l_Perspective.Folders.Item(ActiveFolderId)
   End If
   If Len(DropFolderId) > 0 Then
      Set l_DropFolder = l_Perspective.Folders.Item(DropFolderId)
   End If
   
   '----------------------------------------------------------------------
   ' Add view to drop folder
   '----------------------------------------------------------------------
   Select Case DropRelation
   
      Case vbRelFolder:     ' The drop folder already exists -> Get it from the controls
                            Set l_ucDropFolder = Controls.Item(DropFolderId)
      
      Case vbRelFloating:   ' Create a floating window
      
                            ' -----------------------------------------------------------------------
                            ' Check if the draged folder is already a floating window
                            ' -----------------------------------------------------------------------
                            If l_DragFolder.Relationship <> vbRelFloating Then
                               
                               ' If the drag folder is not floating then create a new folder
                               Set l_DropFolder = l_Perspective.AddFolder(GetUniqueFolderId, DropRelation, 0.5, DropFolderId)
                               Set l_ucDropFolder = CreateFolder(l_DropFolder, False)
                               
                            Else
                               
                               ' If the drag folder is already floating then get folder from Controls
                               Set l_ucDropFolder = Controls.Item(l_DragFolder.FolderId)
                               
                            End If
                               
                            ' -----------------------------------------------------------------------
                            ' Calculate and set the new position and size, depends on the
                            ' position of the mouse cursor.
                            ' -----------------------------------------------------------------------
                            
                            ' Get the current size of the draged folder and the cursor position
                            Set l_ucDragFolder = Controls.Item(l_DragFolder.FolderId)
                            GetCursorPos l_CursorPos
                            GetWindowRect l_ucDragFolder.hwnd, l_FolderRect
                            
                            ' Set the size of the drop folder
                            With l_ucDropFolder
                               
                               w = (l_FolderRect.Right - l_FolderRect.Left)
                               h = (l_FolderRect.Bottom - l_FolderRect.Top)
                                                            
                               .LeftPos = l_CursorPos.x - m_CursorPos.x
                               .RightPos = .LeftPos + w
                               .TopPos = l_CursorPos.y - m_CursorPos.y
                               .BottomPos = .TopPos + h - m_CursorPos.y
      
                               .Refresh
                            End With
                            
                            ' -----------------------------------------------------------------------
                            ' If the folder is floating (parent window handle is not
                            ' equal to the perspective control window handle)
                            ' -----------------------------------------------------------------------
                            If GetParent(l_ucDropFolder.hwnd) <> hwnd Then
                            
                               ' Move the window
                               Set frm = m_Windows.Item(l_ucDropFolder.FolderId)
                               
                               With frm
                                  .Top = (l_CursorPos.y - m_CursorPos.y) * Screen.TwipsPerPixelY
                                  .Left = (l_CursorPos.x - m_CursorPos.x) * Screen.TwipsPerPixelX
                                  .Width = w * Screen.TwipsPerPixelX
                                  .Height = (h - m_CursorPos.y) * Screen.TwipsPerPixelY
                                  
                                  .RefreshFolder
                               End With
                                                             
                               ' Make the new toolwindow magnetic
                               If Magnetic Then
                                  Call m_MagneticWnd.AddWindow(frm.hwnd, MainHwnd)
                               End If
                               
                            End If
      Case vbRelNone:
      
      Case Else:
                            ' -----------------------------------------------------------------------
                            ' Create a new folder on the perspective control
                            ' -----------------------------------------------------------------------
                            Set l_DropFolder = l_Perspective.AddFolder(GetUniqueFolderId, DropRelation, 0.5, DropFolderId)
                            Set l_ucDropFolder = CreateFolder(l_DropFolder)
                            
   End Select
      
   Refresh
      
   '----------------------------------------------------------------------
   ' Add view to folder
   '----------------------------------------------------------------------
   If Not l_ucDropFolder Is Nothing Then
      With l_ucDropFolder
          If Not .ContainsView(l_DragViewId) Then
            .AddView m_Views.Item(l_DragViewId), True
            l_ViewAdded = True
         End If
         
         If Not l_DropFolder Is Nothing Then
            If Not l_DropFolder.Views.Contains(l_DragViewId) Then
               l_DropFolder.Views.Add l_DragViewId, l_DragViewId
            End If
         End If
         
         .ShowView l_DragViewId
         .Active = True
         .Refresh
      End With
   End If
   
   '----------------------------------------------------------------------
   ' Remove view from drag folder
   '----------------------------------------------------------------------
   If l_ViewAdded Then
      Set l_ucDragFolder = Controls.Item(ActiveFolderId)
      With l_ucDragFolder
         .RemoveView l_DragViewId
         
         ' Activate another view on the drag folder
         If Not .Views.IsEmpty Then
            .ShowView .Views.Item(.Views.Count).ViewId
         End If
         
         If Not l_DragFolder Is Nothing Then
            If l_DragFolder.Views.Contains(l_DragViewId) Then
               l_DragFolder.Views.Remove l_DragViewId
            End If
         End If
         
         .Active = False
         .Refresh
      End With
   End If
   
   'Refresh
   
   Set l_Perspective = Nothing
   Set l_DragFolder = Nothing
   Set l_DropFolder = Nothing
   Set l_ucDragFolder = Nothing
   Set l_ucDropFolder = Nothing
   
End Sub

' The drag timer draws a focus rectangle while the user drags a view.
Private Sub tmrDrag_Timer()
   
   Dim i As Long

   Dim l_Perspective As Perspective
   Dim l_Folder As Folder
   Dim l_NewRelation As vbRelationship
   Dim l_NewRc As RECT
   Dim l_Floating As Boolean
   
   Dim l_CursorPos As POINTAPI
   
   Set l_Perspective = ActivePerspective
     
   Screen.MousePointer = 99
   tmrDrag.Interval = 1
            
   ' If view is maximized -> restore view before start drag
   '
   ' TODO: If start dragging the restore button should set as maximize button
   If Len(m_MaximizedFolderId) > 0 Then
      EventRaise "MaximizeView", m_MaximizedFolderId, ActiveFolderId
   End If
      
   ' Get the current cursor position
   GetCursorPos l_CursorPos
   
   ' Iterate all folders and check if the mouse cursor is above
   For i = l_Perspective.Folders.Count To 0 Step -1
            
      Set l_Folder = l_Perspective.Folders.Item(i)

      l_NewRc = GetFolderRect(l_Folder.FolderId)
         
      If l_Folder.Relationship = vbRelFloating And _
         StrComp(l_Folder.FolderId, ActiveFolderId) = 0 Then
         
         ' Keep a floating window floating
         l_NewRelation = vbRelFloating
         
      Else
         
         ' Check if mouse cursor is over folder
         If PtInRect(l_NewRc, l_CursorPos.x, l_CursorPos.y) = 1 Then
              
            ' Mouse cursor is over the folder -> Set relation to folder relation
            l_NewRelation = vbRelFolder
           
            ' Check if the folder is floating
            If l_Folder.Relationship <> vbRelFloating Then
               ' Calculate the relationship (left, right, top, bottom)
               ' by the position of the cursor
               l_NewRelation = GetRelationship(l_CursorPos, l_NewRc)
            Else
               ' If mouse cursor is over a floating window -> set relationship = folder
               l_NewRelation = vbRelFolder
            End If
           
            ' Check if dropfolder is the editor area
            If l_NewRelation = vbRelFolder And _
               StrComp(l_Folder.FolderId, l_Perspective.ID_EDITOR_AREA) = 0 Then
              
               ' Don't drop the view to the editor area -> create a floating window
               l_NewRelation = vbRelFloating
               
            Else

               ' If mouse cursor is over the drag view then don't allow left, right,
               ' top and bottom relationships. Only folder relation is allowed.
               If StrComp(l_Folder.FolderId, ActiveFolderId) = 0 Then
                  l_NewRelation = vbRelFolder
               End If

               ' Save the id of the folder to drop the view
               m_DropFolderId = l_Folder.FolderId
               
               GoTo Finally
               
            End If
         
         
         End If
      End If
   Next
      
   l_NewRelation = vbRelNone
   
   '---------------------------------------------------------------------
   ' Draw floating view
   '---------------------------------------------------------------------
   If Len(ActiveFolderId) > 0 Then
      
      Dim l_ucFolder As ucFolder
      Dim w As Long
      Dim h As Long
           
      l_NewRelation = vbRelFloating
      
      Set l_ucFolder = Controls.Item(ActiveFolderId)
      GetWindowRect l_ucFolder.hwnd, l_NewRc
      
      With l_NewRc
         w = (.Right - .Left)
         h = (.Bottom - .Top)
         .Left = l_CursorPos.x - m_CursorPos.x
         .Right = .Left + w
         .Top = l_CursorPos.y - m_CursorPos.y
         .Bottom = .Top + h - m_CursorPos.y
      End With
      
      Set l_ucFolder = Nothing
      
   End If
   
Finally:
  
   '---------------------------------------------------------------------
   ' Set the Mouse Cursor
   '---------------------------------------------------------------------
   Select Case l_NewRelation
      Case vbRelBottom:    Screen.MouseIcon = getResourceCursor(CURSOR_ARROW_BOTTOM)
      Case vbRelTop:       Screen.MouseIcon = getResourceCursor(CURSOR_ARROW_TOP)
      Case vbRelLeft:      Screen.MouseIcon = getResourceCursor(CURSOR_ARROW_LEFT)
      Case vbRelRight:     Screen.MouseIcon = getResourceCursor(CURSOR_ARROW_RIGHT)
      Case vbRelFolder:
                          If l_Floating Then
                             Screen.MouseIcon = LoadPicture()
                          Else
                             Screen.MouseIcon = getResourceCursor(CURSOR_ARROW_FOLDER)
                          End If
      Case Else:          Screen.MouseIcon = LoadPicture()
   End Select

   '---------------------------------------------------------------------
   ' Draw drag rectangle
   '---------------------------------------------------------------------
   If Not CompareRect(l_NewRc, m_OldRc) Or _
      l_NewRelation <> m_OldRelation Then
               
      If m_OldRelation <> vbRelNone Then
         ' Remove old rectangle
         DrawDragRect GetDragRelRect(m_OldRelation, m_OldRc)
      End If
           
      ' Drag new rectangle
      If l_NewRelation <> vbRelNone Then
         DrawDragRect GetDragRelRect(l_NewRelation, l_NewRc)
      End If
   End If

   '---------------------------------------------------------------------
   ' Save the current rectangle and relation
   '---------------------------------------------------------------------
   CopyRect m_OldRc, l_NewRc
   m_OldRelation = l_NewRelation

   '---------------------------------------------------------------------
   ' Save the drop rectangle and relation
   '---------------------------------------------------------------------

   Set l_Perspective = Nothing
   Set l_Folder = Nothing
   
End Sub

' Compares two rectangles and returns true if the rectangles are equal.
'
' @param Rect1
'        The first rectangle
' @param Rect2
'        The second rectangle
'
' @return
'        CompareRect Returns true if the rectangles are the same; false otherwise.
Private Function CompareRect(ByRef Rect1 As RECT, ByRef RECT2 As RECT) As Boolean
   
   With Rect1
      If .Left <> RECT2.Left Then
         CompareRect = False
         Exit Function
      ElseIf .Right <> RECT2.Right Then
         CompareRect = False
         Exit Function
      ElseIf .Top <> RECT2.Top Then
         CompareRect = False
         Exit Function
      ElseIf .Bottom <> RECT2.Bottom Then
         CompareRect = False
         Exit Function
      End If
   End With

   CompareRect = True

End Function

' Get the rectangle of the folder.
'
' @FolderId The id of the folder.
'
' @GetFolderRect The rectangle of the folder.
Private Function GetFolderRect(ByVal FolderId As String) As RECT

   Dim l_Folder As Folder
   Dim l_RcParent As RECT ' Parent Rect
   Dim l_ucFolder As ucFolder
   Dim twipsX As Long
   Dim twipsY As Long
   
   twipsX = Screen.TwipsPerPixelX
   twipsY = Screen.TwipsPerPixelY

   If Len(FolderId) > 0 Then

      Set l_ucFolder = Controls.Item(FolderId)
      If l_ucFolder Is Nothing Then Exit Function

      If GetParent(l_ucFolder.hwnd) = UserControl.hwnd Then

         ' Docked folder
         GetWindowRect UserControl.hwnd, l_RcParent

         With GetFolderRect
            .Left = l_RcParent.Left + (l_ucFolder.LeftPos / twipsX)
            .Right = l_RcParent.Left + (l_ucFolder.RightPos / twipsX)
            .Top = l_RcParent.Top + (l_ucFolder.TopPos / twipsY)
            .Bottom = l_RcParent.Top + (l_ucFolder.BottomPos / twipsY)
         End With
      Else

         ' Floating window
         GetWindowRect GetParent(l_ucFolder.hwnd), l_RcParent

         With GetFolderRect
            .Left = l_RcParent.Left
            .Right = l_RcParent.Right
            .Top = l_RcParent.Top
            .Bottom = l_RcParent.Bottom

         End With
      End If
   End If
  
   Set l_Folder = Nothing
   Set l_ucFolder = Nothing
   
End Function

' Returns the relation to drop the folder.
'
' @CursorPos The position of the mouse cursor.
'
' @GetRelationship The relation to drop the folder.
Private Function GetRelationship(CursorPos As POINTAPI, Rc As RECT) As vbRelationship
   
   Dim RcRel As RECT
  
   ' Left
   RcRel = GetDragRelRect(vbRelLeft, Rc)
   
   If PtInRect(RcRel, CursorPos.x, CursorPos.y) = 1 Then
      GetRelationship = vbRelLeft
      Exit Function
   End If
   
   ' Right
   RcRel = GetDragRelRect(vbRelRight, Rc)

   If PtInRect(RcRel, CursorPos.x, CursorPos.y) = 1 Then
      GetRelationship = vbRelRight
      Exit Function
   End If

   ' Top
   RcRel = GetDragRelRect(vbRelTop, Rc)

   If PtInRect(RcRel, CursorPos.x, CursorPos.y) = 1 Then
      GetRelationship = vbRelTop
      Exit Function
   End If

   ' Bottom
   RcRel = GetDragRelRect(vbRelBottom, Rc)

   If PtInRect(RcRel, CursorPos.x, CursorPos.y) = 1 Then
      GetRelationship = vbRelBottom
      Exit Function
   End If
   
   GetRelationship = vbRelFolder
   
End Function

' Returns the rectangle of the folder relation.
'
' @Relation The relation.
' @FolderRect The rectangle of the folder.
'
' @GetDragRelRect The rectangle of the folder relation.
Private Function GetDragRelRect(Relation As vbRelationship, FolderRect As RECT) As RECT
   
   Const REL_SIZE As Long = 100
   Const DOCK_MARGIN As Double = 0.3
   
   With GetDragRelRect
      .Left = FolderRect.Left
      .Right = FolderRect.Right
      .Top = FolderRect.Top
      .Bottom = FolderRect.Bottom
   End With
 
   With GetDragRelRect
      
      Select Case Relation
         
         Case vbRelLeft:
         
            If .Left + REL_SIZE < (.Left + (.Right - .Left) * DOCK_MARGIN) Then
               .Right = .Left + REL_SIZE
            Else
               .Right = .Left + ((.Right - .Left) * DOCK_MARGIN)
            End If
         
         Case vbRelRight:
         
            If .Right - REL_SIZE > (.Right - (.Right - .Left) * DOCK_MARGIN) Then
               .Left = .Right - REL_SIZE
            Else
               .Left = (.Right - (.Right - .Left) * DOCK_MARGIN)
            End If
         
         Case vbRelTop:
         
            If .Top + REL_SIZE < (.Top + (.Bottom - .Top) * DOCK_MARGIN) Then
               .Bottom = .Top + REL_SIZE
            Else
               .Bottom = (.Top + (.Bottom - .Top) * DOCK_MARGIN)
            End If
         
         Case vbRelBottom:
         
            If .Bottom - REL_SIZE > (.Bottom - (.Bottom - .Top) * DOCK_MARGIN) Then
               .Top = .Bottom - REL_SIZE
            Else
               .Top = (.Bottom - (.Bottom - .Top) * DOCK_MARGIN)
            End If
         
      End Select
      
   End With
   
End Function

' Draws a focus rectangle.
'
' @Rc The rectangle of the focus rectangle.
' @Width The width of the rectangle border.
Private Sub DrawRect(Rc As RECT, Optional ByVal Width As Long = 2)
        
   Dim DrawRect As RECT
   Dim hdc As Long
   Dim i As Long
        
   hdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)

   For i = 0 To Width
           
      With DrawRect
         .Top = Rc.Top + i
         .Bottom = Rc.Bottom - i
         .Left = Rc.Left + i
         .Right = Rc.Right - i
      End With
           
      DrawFocusRect hdc, DrawRect
           
   Next i
        
   DeleteDC hdc
        
End Sub

' Close a view by its id. The view control won't be unloaded, it will just get hidden.
'
' @ViewId The id of the view to close.
Public Sub CloseView(ByVal ViewId As String)

   Dim i As Long
   Dim l_Cancel As Boolean
   Dim l_Perspective As Perspective
   Dim l_View As View
   Dim l_Folder As Folder
   Dim l_ucFolder As ucFolder
   
   Set l_Perspective = ActivePerspective
   
   If l_Perspective Is Nothing Then
      Exit Sub
   End If
   
   With l_Perspective.Folders
      
      ' Iterate all perspective folders
      For i = 0 To .Count
      
         Set l_Folder = .Item(i)
         Set l_ucFolder = Controls.Item(l_Folder.FolderId)
         
         With l_ucFolder
            
            ' Find the folder which contains the view to close
            If .ContainsView(ViewId) Then
               
               ' Get view by id
               Set l_View = m_Views.Item(ViewId)
               
               ' Raise close events
               If StrComp(l_ucFolder.FolderId, l_Perspective.ID_EDITOR_AREA) = 0 Then
                  RaiseEvent CloseEditor(l_View.View, l_Cancel)
               Else
                  RaiseEvent CloseView(l_View.ViewId, l_Cancel)
               End If
               
               If Not l_Cancel Then
               
                  ' Remove view from folder
                  With l_Folder.Views
                     
                     ' Set placeholder for the closed view (to show view in same folder)
                     If StrComp(l_ucFolder.FolderId, l_Perspective.ID_EDITOR_AREA) = 0 Then
                        ' Unload editor form
                         m_Editors.Remove l_View.ViewId
                         Unload l_View.View
                     Else
                        ' Set a placeholder for the closed view
                        l_Perspective.AddPlaceholder ViewId, l_Folder.FolderId
                     End If
                     
                     .Remove ViewId
                     
                     If .Count > 0 Then
                        l_Folder.ActiveViewId = .Item(.Count - 1)
                     Else
                        If StrComp(l_ucFolder.FolderId, l_Perspective.ID_EDITOR_AREA) = 0 Then
                           'l_ucFolder.ActiveViewId = ""
                           l_Folder.ActiveViewId = ""
                        End If
                     End If
                     
                  End With
                  
                  .RemoveView ViewId
                  
                  ' The view control won't be unloaded, just hide it
                  ShowWindow l_View.View.hwnd, SW_HIDE
                  SetParent l_View.View.hwnd, 0
                  
                  .Refresh
               End If
               
               Exit Sub
               
            End If
         End With
      Next i
   End With
      
   Set l_Perspective = Nothing
   Set l_View = Nothing
      
End Sub

' This function removes a folder by its id.
'
' @FolderId The id of the folder.
Private Sub RemoveFolder(ByVal FolderId As String)
 
   Dim i As Long
   Dim l_ParentHwnd As Long
   Dim l_Perspective As Perspective
   Dim l_Folder As Folder
   Dim l_ucFolder As ucFolder
   Dim l_ucSplitBar As ucSplitBar
   Dim l_ucRemoveFolder As ucFolder
   Dim l_RemoveFolder As Folder
   Dim l_LastFolderId As String
   Dim l_NewLastFolderId As String
   
   Set l_Perspective = ActivePerspective
   
   Set l_RemoveFolder = l_Perspective.Folders.Item(FolderId)
   Set l_ucRemoveFolder = Controls.Item(FolderId)
   
   l_LastFolderId = l_ucRemoveFolder.LastRefFolderId
   
   With l_Perspective.Folders
      
      If Len(l_LastFolderId) > 0 Then
         
         For i = 0 To .Count
            Set l_Folder = .Item(i)
           
            On Error Resume Next
            Set l_ucFolder = Controls.Item(l_Folder.FolderId)
            On Error GoTo 0
            
            If Not l_ucFolder Is Nothing Then
               If StrComp(l_ucFolder.LastRefFolderId, l_RemoveFolder.FolderId) = 0 Then
                  l_ucFolder.LastRefFolderId = l_LastFolderId
                  Exit For
               End If
            End If
         Next i

         For i = 0 To .Count
            Set l_Folder = .Item(i)
           
            If StrComp(l_Folder.RefId, l_RemoveFolder.FolderId) = 0 Then
               l_Folder.RefId = l_LastFolderId
               
               If StrComp(l_LastFolderId, l_Folder.FolderId) <> 0 Then
                  l_NewLastFolderId = l_Folder.FolderId
               End If
            End If
         Next i
         
      Else

         For i = 0 To .Count
            Set l_Folder = .Item(i)

            If StrComp(l_Folder.RefId, l_RemoveFolder.FolderId) = 0 Then
               l_Folder.RefId = l_RemoveFolder.RefId
               l_Folder.Relationship = l_RemoveFolder.Relationship
               l_Folder.Ratio = l_RemoveFolder.Ratio
               
               l_NewLastFolderId = l_Folder.FolderId

            End If
         Next i
      
         For i = 0 To .Count
            Set l_Folder = .Item(i)
            
            On Error Resume Next
            Set l_ucFolder = Controls.Item(l_Folder.FolderId)
            On Error GoTo 0
            If StrComp(l_ucFolder.LastRefFolderId, l_RemoveFolder.FolderId) = 0 Then
               l_ucFolder.LastRefFolderId = l_NewLastFolderId
            End If
         Next i
      
      End If
      
      If Len(l_LastFolderId) > 0 Then
         .Remove l_LastFolderId
        
         Dim Idx As Long
         Idx = .IndexOf(FolderId)
        
         .Remove FolderId
        
         l_RemoveFolder.FolderId = l_LastFolderId
        
         .Add l_RemoveFolder.FolderId, l_RemoveFolder, Idx
         
         On Error Resume Next
         Set l_ucFolder = Controls.Item(l_RemoveFolder.FolderId)
         On Error GoTo 0
         
         If Not l_ucFolder Is Nothing Then
            l_ucFolder.LastRefFolderId = l_NewLastFolderId
         End If
      Else
         .Remove FolderId
      End If
   
      For i = 0 To .Count
         Set l_Folder = .Item(i)

            On Error Resume Next
            Set l_ucSplitBar = Controls.Item("SplitBar_" & l_Folder.FolderId)
            On Error GoTo 0

            If Not l_ucSplitBar Is Nothing Then
               Set l_ucSplitBar.Folder = l_Folder
               Select Case l_Folder.Relationship
                  Case vbRelLeft, vbRelRight
                     l_ucSplitBar.Orientation = espVertical
                  Case Else
                     l_ucSplitBar.Orientation = espHorizontal
               End Select
            End If

      Next i
      
   End With
   
   

   
   l_ParentHwnd = GetParent(l_ucRemoveFolder.hwnd)
   
   On Error Resume Next
   Controls.Remove FolderId
   Controls.Remove "SplitBar_" & FolderId
   On Error GoTo 0
   
   If l_ParentHwnd <> hwnd Then
      'CloseWindow GetParent(l_ucRemoveFolder.hWnd)
      
      m_Windows.Remove FolderId
      DestroyWindow l_ParentHwnd ', WM_QUIT, 0, 0
      
   End If
   Refresh
   
   Set l_Perspective = Nothing
   Set l_ucRemoveFolder = Nothing
   Set l_RemoveFolder = Nothing
   Set l_Folder = Nothing
   
End Sub

' Raise an event by its name.
' ATTENTION: This public method is only for internal use. Child controls will call this method.
'
' @EventName The name of the event to raise.
' @FolderId The current folder id.
' @ViewId The current view id.
Public Sub EventRaise(ByVal EventName As String, ByVal FolderId As String, ByVal ViewId As String)
           
   Dim l_Perspective As Perspective
   Dim l_Folder As Folder
           
   Set l_Perspective = ActivePerspective
   
   ' ----------------------------------------------------------------
   ' Set active folder and active view
   ' ----------------------------------------------------------------
   If Len(FolderId) > 0 Then
      m_ActiveFolderId = FolderId
   End If
   If Len(ViewId) > 0 Then
      m_ActiveViewId = ViewId
   End If
   
   ' ----------------------------------------------------------------
   ' Save the folders active view
   ' ----------------------------------------------------------------
   Set l_Folder = l_Perspective.Folders.Item(FolderId)
       
   If Not l_Folder Is Nothing Then
       l_Folder.ActiveViewId = ViewId
   End If
   
   ' ----------------------------------------------------------------
   ' Raise event by name
   ' ----------------------------------------------------------------
   Select Case EventName
   
      Case "ActivateView":
                            ShowView ViewId
                            
      Case "CloseView":

                            CloseView ViewId
                            
      Case "MaximizeView":
                            tmrDrag.Enabled = False
                            
                            If Len(m_MaximizedFolderId) = 0 Then
                               m_MaximizedFolderId = FolderId
                               RaiseEvent MaximizeView(ViewId)
                            Else
                               m_MaximizedFolderId = vbNullString
                               RaiseEvent RestoreView(ViewId)
                            End If
                            
                            Refresh
      
                            RefreshFolderControls
                            
      Case "RemoveFolder":
      
                            m_MaximizedFolderId = vbNullString
                            
                            If StrComp(l_Perspective.ID_EDITOR_AREA, FolderId) <> 0 Then
                               RemoveFolder FolderId
                            End If
                            
                            Refresh
      
                            RefreshFolderControls
      
      Case "PrevView":
                            If StrComp(FolderId, l_Perspective.ID_EDITOR_AREA) = 0 Then
                               PreviousEditor
                            Else
                               PreviousView
                            End If
      
      Case "NextView":
                            If StrComp(FolderId, l_Perspective.ID_EDITOR_AREA) = 0 Then
                               NextEditor
                            Else
                               NextView
                            End If
      
      Case "StartDrag":
                            ShowView ViewId
                            
                            tmrDrag.Interval = 400
                            tmrDrag.Enabled = True
                            
                            RaiseEvent ActivateView(ViewId)
                            
      Case "EndDrag", "FloatingWindow":
      
                            tmrDrag.Enabled = False
                            
                            If (tmrDrag.Interval < 400) Or _
                                StrComp(EventName, "FloatingWindow") = 0 Then
                                
                                tmrDrag.Interval = 400
                                
                                If StrComp(EventName, "FloatingWindow") = 0 Then
                                   ' Create a floating window without drag and drop
                                   m_DropFolderId = vbNullString
                                   m_OldRelation = vbRelFloating
                                Else
                                   ' Remove drag rectangle
                                   If m_OldRelation <> vbRelNone Then
                                      DrawDragRect GetDragRelRect(m_OldRelation, m_OldRc)
                                   End If
                                End If
                                
                                If StrComp(m_DropFolderId, ActiveFolderId) <> 0 Or _
                                   m_OldRelation = vbRelFloating Then
                                   DropActiveViewOn m_DropFolderId, m_OldRelation
                                End If
                            
                            End If
                            
                            m_OldRelation = vbRelNone
                            m_DropFolderId = vbNullString
                            
                            Screen.MousePointer = vbDefault

   End Select
   
End Sub

' Returns the active perspective id.
'
' @ActivePerspectiveId The active perspective id.
Public Property Get ActivePerspectiveId() As String
   ActivePerspectiveId = m_ActivePerspectiveId
End Property

' Returns the active folder id.
'
' @ActiveFolderId The active folder id.
Public Property Get ActiveFolderId() As String
   ActiveFolderId = m_ActiveFolderId
End Property

' Returns the active view id.
'
' @ActiveViewId The active view id.
Public Property Get ActiveViewId() As String
   ActiveViewId = m_ActiveViewId
End Property

' Get the active editor.
'
' @ActiveEditor Returns the active editor or nothing if no editor is open.
Public Property Get ActiveEditor() As Object
   
   On Error GoTo Finally
   
   Dim l_Perspective As Perspective
   Dim l_ucFolder As ucFolder
   Dim l_View As View
   
   Set l_Perspective = ActivePerspective
   
   If l_Perspective Is Nothing Then Exit Property ' No active perspective found!
   
   On Error Resume Next
   Set l_ucFolder = Controls.Item(l_Perspective.ID_EDITOR_AREA)
   On Error GoTo Finally
   
   If l_ucFolder Is Nothing Then Exit Property ' No editor area found!
      
   If l_ucFolder.Views.IsEmpty Then Exit Property ' No open editor !
   
   Set l_View = l_ucFolder.Views.Item(l_ucFolder.ActiveViewId)
   
   Set ActiveEditor = l_View.View
   
Finally:

   Set l_Perspective = Nothing
   Set l_ucFolder = Nothing
   
End Property

' Get the active editor.
'
' @ActiveEditor Returns the active editor or nothing if no editor is open.
Public Sub CloseEditor(ByRef Editor As Object)
   
   On Error GoTo Finally
   
   Dim i As Long
   Dim l_Perspective As Perspective
   Dim l_ucFolder As ucFolder
   Dim l_View As View
   
   Set l_Perspective = ActivePerspective
   
   If l_Perspective Is Nothing Then Exit Sub      ' No active perspective found!
   
   On Error Resume Next
   Set l_ucFolder = Controls.Item(l_Perspective.ID_EDITOR_AREA)
   On Error GoTo Finally
   
   If l_ucFolder Is Nothing Then Exit Sub      ' No editor area found!
      
   If l_ucFolder.Views.IsEmpty Then Exit Sub      ' No open editor !
   
   With l_ucFolder.Views
      For i = 0 To .Count
         Set l_View = .Item(i)
         
         If l_View.View.hwnd = Editor.hwnd Then
            CloseView l_View.ViewId
           
            GoTo Finally
         End If
      Next i
   End With
   
Finally:

   Set l_Perspective = Nothing
   Set l_ucFolder = Nothing
   
End Sub

' This method returns a array of all open editors or nothing if no editor is open.
'
' The returned array contains a list of forms.
'
' @Editors Returns a array of all editors or nothing if no editor is open.
Public Property Get Editors() As Variant
   
   On Error GoTo Finally
   
   Dim l_Perspective As Perspective
   Dim l_ucFolder As ucFolder
   Dim l_View As View
   Dim EditorList() As Object
   Dim i As Long
   
   Set l_Perspective = ActivePerspective
   
   If l_Perspective Is Nothing Then Exit Property ' No active perspective found!
   
   On Error Resume Next
   Set l_ucFolder = Controls.Item(l_Perspective.ID_EDITOR_AREA)
   On Error GoTo Finally
   
   If l_ucFolder Is Nothing Then Exit Property ' No editor area found!
      
   With l_ucFolder.Views

      If .IsEmpty Then Exit Property ' No open editor !
 
      For i = 0 To .Count
       
         Set l_View = .Item(i)
           
         ReDim Preserve EditorList(i) As Object
         Set EditorList(i) = l_View.View
        
      Next
      
   End With
   
   Editors = EditorList
   
Finally:

   Set l_Perspective = Nothing
   Set l_ucFolder = Nothing
   
End Property

' Returns the active perspective.
'
' @ActivePerspective The active perspective.
Private Property Get ActivePerspective() As Perspective
      
   If Len(ActivePerspectiveId) > 0 Then
      Set ActivePerspective = Perspectives.Item(ActivePerspectiveId)
   Else
      Err.Raise 1, , "No active perspective!"
   End If
   
   If ActivePerspective Is Nothing Then
      Err.Raise 1, , "No active perspective!"
   End If
   
End Property

' Returns a list of all added perspectives.
'
' @Perspectives List of all perspectives.
Public Property Get Perspectives() As List
   Set Perspectives = m_Perspectives
End Property

' Adds a view to the public view list.
'
' @ViewId The id of the added view.
' @View The view control is a normal Visual Basic form (which chould implement the IViewPart interface).
Public Sub AddView(ByVal ViewId As String, ByRef View As Object)
    
    Dim l_View As View
    Set l_View = New View
    
    With l_View
       .ViewId = ViewId
       Set .View = View
    End With
    
    m_Views.Add ViewId, l_View
    
    Set l_View = Nothing
    
End Sub


' Creates and adds a new perspective.
'
' @PerspectiveId The id of the new perspective.
'
' @AddPerspective Returns the created perspective instance.
Public Function AddPerspective(ByVal PerspectiveId As String) As Perspective
   
   CheckId PerspectiveId
   
   Set AddPerspective = New Perspective
       AddPerspective.PerspectiveId = PerspectiveId
      
   m_Perspectives.Add PerspectiveId, AddPerspective
   
End Function

' This method indicates the user control to create the visual components of
' a perspective. You can add a new perspective if you call the AddPerspective()
' method. The AddPerspective() method expects a string parameter which represents
' the perspectives id.
'
' @PerspectiveId The id of the perspective to show.
Public Sub ShowPerspective(ByVal PerspectiveId As String)
   
   Dim i As Long
   Dim l_Perspective As Perspective
   Dim l_Folder As Folder
   
   ' Check if the perspective is still active
   If StrComp(m_ActivePerspectiveId, PerspectiveId) = 0 Then
      Exit Sub
   End If
   
   On Error Resume Next
   Set l_Perspective = ActivePerspective
   On Error GoTo 0
   
   If Not l_Perspective Is Nothing Then
      ' Close active perspective before showing next
      ClosePerspective
      
      If Len(ActivePerspectiveId) > 0 Then
         ' Closing the perspective was canceled by user
         Exit Sub
      End If
   End If
   
   If Not m_Perspectives.Contains(PerspectiveId) Then Err.Raise 1, , "No perspective id found!"
   
   Set l_Perspective = m_Perspectives.Item(PerspectiveId)
   
   If l_Perspective Is Nothing Then Err.Raise 1, , "No perspective found!"
   
   If StrComp(m_ActivePerspectiveId, PerspectiveId) = 0 Then
      Exit Sub
   ElseIf Len(m_ActivePerspectiveId) > 0 Then
      ClosePerspective
   End If
   
   m_ActivePerspectiveId = PerspectiveId
   
   RaiseEvent OpenPerspective(PerspectiveId)
   
   ' ---------------------------------------------------------------------------------
   ' Add folder for the editor area
   ' ---------------------------------------------------------------------------------
   If l_Perspective.EditorAreaVisible Then
      
      If Not l_Perspective.Folders.Contains(l_Perspective.ID_EDITOR_AREA) Then
         
         Set l_Folder = New Folder
             l_Folder.FolderId = l_Perspective.ID_EDITOR_AREA
             
         l_Perspective.Folders.Add l_Folder.FolderId, l_Folder, 0
         
         CreateFolder l_Folder
      End If
      
      Set l_Folder = Nothing
      
   End If
   
   ' ---------------------------------------------------------------------------------
   ' Create all folders (+ splitbars)
   ' ---------------------------------------------------------------------------------
   With l_Perspective.Folders
      For i = 0 To .Count
         CreateFolder .Item(i)
      Next i
   End With
   
   If Len(l_Perspective.ActiveViewId) > 0 Then
      ShowView l_Perspective.ActiveViewId
   End If
   
   Refresh
   
   RefreshFolderControls
      
   Set l_Perspective = Nothing
   
End Sub

' Iterates and refreshs all folder controls.
Private Sub RefreshFolderControls()
   
   Dim i As Long
   Dim ctrl As Control
   
   For i = 0 To Controls.Count - 1
   
      Set ctrl = Controls.Item(i)
      
      If StrComp(TypeName(ctrl), "ucFolder") = 0 Then
         ctrl.Refresh
      End If
      
   Next i

End Sub

' Close the active perspective.
Public Sub ClosePerspective()
   
   On Error Resume Next
   
   Dim f As Long
   Dim v As Long
   Dim frm As frmToolWin
   Dim l_Cancel As Boolean
   Dim l_Perspective As Perspective
   Dim l_View As View
   Dim l_Folder As Folder
   Dim l_ucFolder As ucFolder
   Dim l_ucSplitBar As ucSplitBar
   
   Set l_Perspective = ActivePerspective
   
   If l_Perspective Is Nothing Then Err.Raise 1, , "No active perspective found!"
      
   RaiseEvent ClosePerspective(ActivePerspectiveId, l_Cancel)
   
   If l_Cancel Then
      ' Closing the perspective was canceled by user
      Exit Sub
   End If
      
   l_Perspective.ActiveViewId = ActiveViewId
   
   ' ---------------------------------------------------------------------------------
   ' Remove floating windows
   ' ---------------------------------------------------------------------------------
   With m_Windows
      For f = 0 To .Count
          
          Set frm = .Item(f)
          
          If Not frm Is Nothing Then
             
             Set l_ucFolder = frm.Folder
          
             If Not l_ucFolder Is Nothing Then
             
                With l_ucFolder.Views
                   For v = 0 To .Count
                       Set l_View = .Item(v)
                       
                       If Not l_View Is Nothing Then
                          ShowWindow l_View.View.hwnd, SW_HIDE
                          SetParent l_View.View.hwnd, 0
                          
                          modSubClass.UnHook l_View.View.hwnd
                       End If
                       
                   Next v
                End With
                
                Set l_Folder = l_Perspective.Folders.Item(l_ucFolder.FolderId)
          
                With l_Folder.Position
                   .Left = frm.Left / Screen.TwipsPerPixelX
                   .Right = (frm.Width / Screen.TwipsPerPixelX)
                   .Top = frm.Top / Screen.TwipsPerPixelY
                   .Bottom = (frm.Height / Screen.TwipsPerPixelY)
                End With
             End If
          
             Unload frm
             
          End If
          
      Next f
      
      .Clear
      
   End With
   
   ' ---------------------------------------------------------------------------------
   ' Remove all folders (+ splitbars)
   ' ---------------------------------------------------------------------------------
   With l_Perspective.Folders
      For f = 0 To .Count
          
          Set l_Folder = .Item(f)
          
          If Not l_Folder Is Nothing Then
             
             ' ----------------------------------------------------------------------
             ' Remove folder
             ' ----------------------------------------------------------------------
             Set l_ucFolder = Controls.Item(l_Folder.FolderId)
             
             If Not l_ucFolder Is Nothing Then
                
                If StrComp(l_Folder.FolderId, l_Perspective.ID_EDITOR_AREA) = 0 Then
                   
                   ShowWindow l_ucFolder.hwnd, SW_HIDE
                   
                Else
                   
                   With l_ucFolder.Views
                      For v = 0 To .Count
                          Set l_View = .Item(v)
                          
                          If Not l_View Is Nothing Then
                             ShowWindow l_View.View.hwnd, SW_HIDE
                             SetParent l_View.View.hwnd, 0
                             
                             modSubClass.UnHook l_View.View.hwnd
                          End If
                          
                      Next v
                   End With
                   
                   Controls.Remove l_ucFolder
                   
                End If
                
             End If
             
             ' ----------------------------------------------------------------------
             ' Remove splitbar
             ' ----------------------------------------------------------------------
             Set l_ucSplitBar = Controls.Item("SplitBar_" & l_Folder.FolderId)
             
             If Not l_ucSplitBar Is Nothing Then
                Controls.Remove l_ucSplitBar
             End If
          End If
          
      Next f
   End With
      
   m_ActivePerspectiveId = vbNullString
   m_MaximizedFolderId = vbNullString
   
   Set l_Perspective = Nothing
   Set l_Folder = Nothing
   Set l_ucFolder = Nothing
   Set l_ucSplitBar = Nothing
   
End Sub

' Creates a new folder control from a folder description and returns the control.
' Optionally the function creates a splitter bar to resize the folder (default is true).
'
' @Folder The folder description as a folder instance.
' @SplitBar Create an optionally split bar (true is default).
'
' @CreateFolder Returns the created folder control.
Private Function CreateFolder(ByRef Folder As Folder, _
                     Optional ByVal SplitBar As Boolean = True) As ucFolder

   On Error GoTo errorHandler
   
   Dim i As Long
   
   Dim ViewId As String
   Dim l_Perspective As Perspective
   Dim l_ucFolder As Object
   Dim l_ucFolderRef As ucFolder
   Dim l_ucSplitBar As Object
   
   Set l_Perspective = ActivePerspective
   
   If StrComp(Folder.FolderId, l_Perspective.ID_EDITOR_AREA) = 0 Then
      
      On Error Resume Next
      Set l_ucFolder = Controls.Item(Folder.FolderId)
      On Error GoTo errorHandler
      
      If Not l_ucFolder Is Nothing Then
         
         l_ucFolder.Refresh
         
         ShowWindow l_ucFolder.hwnd, SW_SHOW
         Exit Function
      End If
      
   End If
   
   Set l_ucFolder = Controls.Add("Aggressive.ucFolder", Folder.FolderId)
   
   ' ------------------------------------------------------------------------------------------
   ' Create Folder
   ' ------------------------------------------------------------------------------------------
   With l_ucFolder
      
      .FolderId = Folder.FolderId
      
      With Folder.Views
         For i = 0 To .Count
            ViewId = .Item(i)
            l_ucFolder.AddView m_Views.Item(ViewId), False
         Next i
         
         If Not .IsEmpty Then
            l_ucFolder.ShowView .Item(0)
         End If
      End With
    
      If Folder.Relationship = vbRelFloating Then

         ' Remove the maximize button
        .MaximizeAble = False

        ' Create a new floating window and hide it
        Dim frm As frmToolWin
        Set frm = New frmToolWin
        Set frm.Folder = l_ucFolder
        Set frm.FolderModel = Folder
       

        frm.Visible = False

        ' Set a new window style and move the window
        SetWindowStyle frm.hwnd, VbToolWin

        ' Set the window position
        If Not (Folder.Position.Left = 0 And Folder.Position.Right = 0 And Folder.Position.Top = 0 And Folder.Position.Bottom = 0) Then
           .LeftPos = Folder.Position.Left
           .RightPos = Folder.Position.Right
           .TopPos = Folder.Position.Top
           .BottomPos = Folder.Position.Bottom

           SetWindowPos frm.hwnd, 0&, .LeftPos, .TopPos, .RightPos, .BottomPos, 0& ' SWP_FRAMECHANGED 'SWP_NOSIZE Or SWP_NOMOVE
        End If

        ' Show the floating window
        frm.Visible = True

        ' Make the new toolwindow magnetic
        If Magnetic Then
           Call m_MagneticWnd.AddWindow(frm.hwnd, MainHwnd)
        End If

         ' Set new window owner
        SetParent .hwnd, frm.hwnd
        SetWindowLong frm.hwnd, GWL_HWNDPARENT, MainHwnd

        m_Windows.Add l_ucFolder.FolderId, frm
      End If
      
      ' Set the active view of the folder
      .ShowView Folder.ActiveViewId
      
      .Refresh
      .TabStop = False
      .Active = False
      .Visible = False
   End With

   ' ------------------------------------------------------------------------------------------
   ' Set the last added folder id
   ' ------------------------------------------------------------------------------------------
   If Len(Folder.RefId) > 0 Then
      
      On Error Resume Next
      Set l_ucFolderRef = Controls.Item(Folder.RefId)
      On Error GoTo 0
      
      If Not l_ucFolderRef Is Nothing Then
         l_ucFolderRef.LastRefFolderId = Folder.FolderId
      End If
      
   End If
   
   If Not SplitBar Then GoTo Finally
   
   ' ------------------------------------------------------------------------------------------
   ' Create split bar
   ' ------------------------------------------------------------------------------------------
   If Len(Folder.RefId) > 0 And _
      StrComp(Folder.FolderId, l_Perspective.ID_EDITOR_AREA, vbBinaryCompare) <> 0 Then

      On Error Resume Next
      Set l_ucSplitBar = Controls.Item("SplitBar_" & Folder.FolderId)
      On Error GoTo 0

      If l_ucSplitBar Is Nothing Then
         Set l_ucSplitBar = Controls.Add("Aggressive.ucSplitBar", "SplitBar_" & Folder.FolderId)
      End If

      With l_ucSplitBar
         If Folder.Relationship = vbRelLeft Or _
            Folder.Relationship = vbRelRight Then
            .Orientation = espVertical
         Else
            .Orientation = espHorizontal
         End If

         Set .Folder = Folder

         .TabStop = False
         .Visible = False

      End With
   
   End If
   
   If Not l_ucFolder Is Nothing And _
      Folder.Relationship <> vbRelFloating Then
      l_ucFolder.ZOrder
      l_ucFolder.Refresh
   End If
   
Finally:
   
   Set CreateFolder = l_ucFolder
   
   Set l_Perspective = Nothing
   Set l_ucFolder = Nothing
   Set l_ucFolderRef = Nothing
   
   Exit Function
   
errorHandler:
   
   MsgBox Err.Description
   
   GoTo Finally
   
End Function

' Refreshs the perspective layout.
Public Sub Refresh()

   On Error GoTo errorHandler
   
   Dim i As Long
   Dim l_ucFolder As Variant
   Dim l_ucSplitBar As Variant
   Dim l_Folder As Folder
   Dim l_Perspective As Perspective
   
   Set l_Perspective = ActivePerspective
   
   With l_Perspective.Folders
     
   ' Folder is maximized
   If Len(m_MaximizedFolderId) > 0 Then
   
      Set l_Folder = .Item(m_MaximizedFolderId)
      
      If l_Folder.Relationship = vbRelFloating Then
         Exit Sub
      End If
      
      For i = 0 To .Count
         
         Set l_Folder = .Item(i)
         
         If Not l_Folder Is Nothing Then
            If l_Folder.Relationship <> vbRelFloating Then
         
               ' -----------------------------------------------------------------------------------
               ' Move Folder
               ' -----------------------------------------------------------------------------------
               On Error Resume Next
               Set l_ucSplitBar = Controls.Item("SplitBar_" & l_Folder.FolderId)
               Set l_ucFolder = Controls.Item(l_Folder.FolderId)
               On Error GoTo errorHandler
               
               If Not l_ucFolder Is Nothing Then
                  With l_ucFolder
                     If StrComp(l_Folder.FolderId, m_MaximizedFolderId, vbBinaryCompare) = 0 Then
                        .Visible = True
                        l_ucFolder.Move 1, 1, ScaleWidth, ScaleHeight
                        l_ucFolder.Refresh
                     Else
                        .Visible = False
                     End If
                  End With
               End If
           
            End If
         End If
      Next i
   
   Else ' No maximized folder
   
      For i = 0 To .Count
         Set l_Folder = .Item(i)

         'If l_Folder.Relationship <> vbRelFloating Then
            CalculateFolderPosition .Item(i)
         'End If
      Next i
   
      For i = 0 To .Count
         Set l_Folder = .Item(i)
         
         If Not l_Folder Is Nothing And l_Folder.Relationship <> vbRelFloating Then
            
            ' -----------------------------------------------------------------------------------
            ' Move Folder
            ' -----------------------------------------------------------------------------------
            On Error Resume Next
            Set l_ucFolder = Controls.Item(l_Folder.FolderId)
            On Error GoTo errorHandler
            
            If Not l_ucFolder Is Nothing Then
               With l_ucFolder
                  l_ucFolder.Move .LeftPos, .TopPos, .RightPos - .LeftPos, .BottomPos - .TopPos
                  .Refresh
                  .Visible = True
               End With
            End If
         
            ' -----------------------------------------------------------------------------------
            ' Move Splitbar
            ' -----------------------------------------------------------------------------------
            If StrComp(l_Folder.FolderId, l_Perspective.ID_EDITOR_AREA, vbBinaryCompare) <> 0 Then
               
               On Error Resume Next
               Set l_ucSplitBar = Controls.Item("SplitBar_" & l_Folder.FolderId)
               On Error GoTo errorHandler
         
               If Not l_ucSplitBar Is Nothing Then
                  With l_ucSplitBar
                     l_ucSplitBar.Move .LeftPos, .TopPos, .RightPos - .LeftPos, .BottomPos - .TopPos
                     .Refresh
                     .Visible = True
                  End With
               End If
               
            End If
         End If
      Next i
      
      ' ---------------------------------------------------------------------
      ' Refresh floating windows
      ' ---------------------------------------------------------------------
      For i = 0 To .Count
         Set l_Folder = .Item(i)

         If Not l_Folder Is Nothing And l_Folder.Relationship = vbRelFloating Then

            On Error Resume Next
            Set l_ucFolder = Controls.Item(l_Folder.FolderId)
            On Error GoTo errorHandler
            With l_ucFolder
               .Visible = True
               l_ucFolder.Refresh
            End With
         End If
      Next i
      
   End If
      
   End With
      
Finally:
   
   Exit Sub
   
errorHandler:
   
   GoTo Finally
   
End Sub

' Calculates the current position of a folder.
'
' @Folder The folder instance.
Private Sub CalculateFolderPosition(ByRef Folder As Folder)

   Dim l_Perspective As Perspective
   Dim l_ucFolder As ucFolder

   Set l_Perspective = ActivePerspective
   
   If Folder Is Nothing Then Exit Sub

   Set l_ucFolder = Controls.Item(Folder.FolderId)
   
   If l_ucFolder Is Nothing Then Exit Sub
   
   If (l_Perspective.EditorAreaVisible And _
       StrComp(Folder.FolderId, l_Perspective.ID_EDITOR_AREA, vbBinaryCompare) = 0) Or _
      (Not l_Perspective.EditorAreaVisible And _
       StrComp(Folder.RefId, l_Perspective.ID_EDITOR_AREA, vbBinaryCompare) = 0) Then
      
      With l_ucFolder
         .TopPos = 1
         .LeftPos = 1
         .RightPos = UserControl.ScaleWidth
         .BottomPos = UserControl.ScaleHeight
      End With
      
   Else
      CalculateFolderPositionByRef Folder
   End If
      
Finally:

   Set l_Perspective = Nothing
   
   Exit Sub
   
errorHandler:
   
   GoTo Finally
End Sub

'
'
'
Private Sub CalculateFolderPositionByRef(ByRef Folder As Folder)
   
   On Error GoTo errorHandler
   
   Const SPLITBAR_WIDTH As Long = 60
      
   Dim l_Rect As RECT
   Dim RatioSize As Long
   Dim l_ucFolder As ucFolder
   Dim l_ucFolderRef As ucFolder
   Dim l_ucSplitBar As ucSplitBar
   Dim p As Perspective
   
   On Error Resume Next
   
   Set p = ActivePerspective
   Set l_ucFolder = Controls.Item(Folder.FolderId)
   Set l_ucFolderRef = Controls.Item(Folder.RefId)
   Set l_ucSplitBar = Controls.Item("SplitBar_" & Folder.FolderId)
   On Error GoTo errorHandler
   
   Select Case Folder.Relationship
   
      Case vbRelLeft, vbRelRight:     RatioSize = (l_ucFolderRef.RightPos - l_ucFolderRef.LeftPos) * CDbl(Folder.Ratio)
      Case vbRelTop, vbRelBottom:     RatioSize = (l_ucFolderRef.BottomPos - l_ucFolderRef.TopPos) * CDbl(Folder.Ratio)
      
   End Select
   
   ' ---------------------------------------------------------------------------------
   ' Set the max rectangle to move the splitbar
   ' ---------------------------------------------------------------------------------
   If Not l_ucFolderRef Is Nothing Then
      If Not l_ucSplitBar Is Nothing Then
         With l_ucSplitBar
      
            GetWindowRect UserControl.hwnd, l_Rect
            .RectLeft = (l_ucFolderRef.LeftPos / Screen.TwipsPerPixelX) + l_Rect.Left
            .RectRight = (l_ucFolderRef.RightPos / Screen.TwipsPerPixelX) + l_Rect.Left
            .RectTop = (l_ucFolderRef.TopPos / Screen.TwipsPerPixelY) + l_Rect.Top
            .RectBottom = (l_ucFolderRef.BottomPos / Screen.TwipsPerPixelY) + l_Rect.Top
         End With
      End If
   End If
   
   ' ---------------------------------------------------------------------------------
   ' Set the positions of the folders + splitbars
   ' ---------------------------------------------------------------------------------
   Select Case Folder.Relationship
   
      Case vbRelLeft:
      
           If Not l_ucFolder Is Nothing Then
              With l_ucFolder
                 .LeftPos = l_ucFolderRef.LeftPos
                 .RightPos = .LeftPos + RatioSize - SPLITBAR_WIDTH
                 .TopPos = l_ucFolderRef.TopPos
                 .BottomPos = l_ucFolderRef.BottomPos
              End With
           End If
           
           If Not l_ucFolderRef Is Nothing Then
              With l_ucFolderRef
                 .LeftPos = l_ucFolder.RightPos + SPLITBAR_WIDTH
              End With
           End If
           
           If Not l_ucSplitBar Is Nothing Then
              With l_ucSplitBar
                 .TopPos = l_ucFolder.TopPos
                 .BottomPos = l_ucFolder.BottomPos
                 .LeftPos = l_ucFolder.RightPos
                 .RightPos = .LeftPos + SPLITBAR_WIDTH
              End With
           End If
           
      Case vbRelRight:
      
           If Not l_ucFolder Is Nothing Then
              With l_ucFolder
                 .LeftPos = l_ucFolderRef.RightPos - RatioSize + SPLITBAR_WIDTH
                 .RightPos = l_ucFolderRef.RightPos
                 .TopPos = l_ucFolderRef.TopPos
                 .BottomPos = l_ucFolderRef.BottomPos
              End With
           End If
           
           If Not l_ucFolderRef Is Nothing Then
              With l_ucFolderRef
                 .RightPos = l_ucFolder.LeftPos - SPLITBAR_WIDTH
              End With
           End If
           
           If Not l_ucSplitBar Is Nothing Then
              With l_ucSplitBar
                 .TopPos = l_ucFolder.TopPos
                 .BottomPos = l_ucFolder.BottomPos
                 .LeftPos = l_ucFolderRef.RightPos
                 .RightPos = .LeftPos + SPLITBAR_WIDTH * 2
              End With
           End If
   
    Case vbRelTop:
      
           If Not l_ucFolder Is Nothing Then
              With l_ucFolder
                 .LeftPos = l_ucFolderRef.LeftPos
                 .RightPos = l_ucFolderRef.RightPos
                 .TopPos = l_ucFolderRef.TopPos
                 .BottomPos = l_ucFolderRef.TopPos + RatioSize + SPLITBAR_WIDTH
              End With
           End If
           
           If Not l_ucFolderRef Is Nothing Then
              With l_ucFolderRef
                 .TopPos = l_ucFolder.BottomPos + SPLITBAR_WIDTH
              End With
           End If

           If Not l_ucSplitBar Is Nothing Then
              With l_ucSplitBar
                 .TopPos = l_ucFolder.BottomPos
                 .BottomPos = .TopPos + SPLITBAR_WIDTH * 2
                 .LeftPos = l_ucFolder.LeftPos
                 .RightPos = l_ucFolder.RightPos
              End With
           End If
           
      Case vbRelBottom:
      
           If Not l_ucFolder Is Nothing Then
              With l_ucFolder
                 .LeftPos = l_ucFolderRef.LeftPos
                 .RightPos = l_ucFolderRef.RightPos
                 .TopPos = l_ucFolderRef.TopPos + RatioSize + SPLITBAR_WIDTH
                 .BottomPos = l_ucFolderRef.BottomPos
              End With
           End If
           
           If Not l_ucFolderRef Is Nothing Then
              With l_ucFolderRef
                 .BottomPos = l_ucFolder.TopPos - SPLITBAR_WIDTH
              End With
           End If

           If Not l_ucSplitBar Is Nothing Then
              With l_ucSplitBar
                 .TopPos = l_ucFolderRef.BottomPos
                 .BottomPos = .TopPos + SPLITBAR_WIDTH * 2
                 .LeftPos = l_ucFolder.LeftPos
                 .RightPos = l_ucFolder.RightPos
              End With
           End If
           
   End Select
      
   Exit Sub
   
errorHandler:
      
'   Stop
   
   MsgBox "This is a known bug!" & _
          vbNewLine & _
          "I will fix it soon!", vbCritical + vbOKOnly, "Known bug"
   
'   Resume Next
      
End Sub

' Show a view by its id. If the view is visible on the active perspective then
' bring it to top and set it active. If the view is not visible the view will be
' opened in the last created folder. And if no folder exists, a new folder would
' be created.
'
' @ViewId The id of the view to show.
Public Sub ShowView(ByVal ViewId As String)
   
   Dim f As Long
   
   Dim l_ucFolder As ucFolder
   Dim l_Perspective As Perspective
   Dim l_FolderId As String
   Dim l_Folder As Folder
   Dim l_View As View
   Dim l_ViewIsVisible As Boolean
   
   Set l_Perspective = ActivePerspective
   
   If l_Perspective Is Nothing Then
      Exit Sub
   End If
   
   With l_Perspective.Folders
      If Not .IsEmpty Then
         
         ' ---------------------------------------------------------------------
         ' Inactivate all folders
         ' ---------------------------------------------------------------------
         For f = 0 To .Count
         
            Set l_Folder = .Item(f)
            Set l_ucFolder = Controls.Item(l_Folder.FolderId)
                
            ' Activate the folder which contains the view id
            If l_ucFolder.ContainsView(ViewId) Then
               
               l_Folder.ActiveViewId = ViewId
               
               With l_ucFolder
                  .ShowView ViewId
                  .Active = True
                  .Refresh
               End With
               
               If StrComp(l_Folder.FolderId, l_Perspective.ID_EDITOR_AREA) = 0 Then
                  RaiseEvent ActivateEditor(ActiveEditor)
               Else
                  RaiseEvent ActivateView(ViewId)
               End If
               
               l_ViewIsVisible = True
            Else
               l_ucFolder.Active = False
            End If
                
         Next f
         
         ' ---------------------------------------------------------------------
         ' View is not visible
         ' ---------------------------------------------------------------------
         If Not l_ViewIsVisible Then
            
            l_FolderId = m_Placeholders.Item(l_Perspective.PerspectiveId & "." & ViewId)
            
            If Len(l_FolderId) > 0 Then
               On Error Resume Next
               Set l_Folder = l_Perspective.Folders.Item(l_FolderId)
               On Error GoTo 0
            End If
            
            If l_Folder Is Nothing Then
               Set l_Folder = l_Perspective.Folders.Item(l_Perspective.Folders.Count)
            End If
            
            If Not l_Folder Is Nothing Then
            
               On Error Resume Next
               Set l_ucFolder = Controls.Item(l_Folder.FolderId)
               Set l_View = m_Views.Item(ViewId)
               On Error GoTo 0
               
               If Not l_View Is Nothing Then
               
                  If StrComp(l_Folder.FolderId, l_Perspective.ID_EDITOR_AREA) = 0 Then
                     
                     ' Create a new folder
                     Set l_Folder = l_Perspective.AddFolder(GetUniqueFolderId, vbRelRight, 0.5, l_Perspective.ID_EDITOR_AREA)
                         l_Folder.AddView ViewId
                     
                     Set l_ucFolder = CreateFolder(l_Folder, True)
                     
                     Refresh
                     
                     l_ucFolder.Refresh
                     
                  Else
                    
                     ' Add view to last added view folder
                     l_Folder.AddView l_View.ViewId
                     l_ucFolder.AddView l_View, True
                     l_ucFolder.ShowView l_View.ViewId
                     l_ucFolder.Active = True
                     l_ucFolder.Refresh
                     
                  End If
                     
                  RaiseEvent ActivateView(l_View.ViewId)
                     
               End If
               
            End If
            
         End If
         
      End If
   End With
   
   Set l_Perspective = Nothing
   
End Sub

' Refreshs a view by its id.
'
' @ViewId The id of the view to refresh.
Public Sub RefreshView(ByVal ViewId As String)
   
   Dim f As Long
   
   Dim l_ucFolder As ucFolder
   Dim l_Perspective As Perspective
   Dim l_FolderId As String
   Dim l_Folder As Folder
   Dim l_View As View
   
   Set l_Perspective = ActivePerspective
   
   If l_Perspective Is Nothing Then
      Exit Sub
   End If
   
   With l_Perspective.Folders
      If Not .IsEmpty Then
         
         ' ---------------------------------------------------------------------
         ' Inactivate all folders
         ' ---------------------------------------------------------------------
         For f = 0 To .Count
         
            Set l_Folder = .Item(f)
            Set l_ucFolder = Controls.Item(l_Folder.FolderId)
                
            ' Activate the folder which contains the view id
            If l_ucFolder.ContainsView(ViewId) Then
               
               If StrComp(l_Folder.ActiveViewId, ViewId) Then
               
                  With l_ucFolder
                     .RefreshView ViewId
                     .Active = True
                     .Refresh
                  End With
                    
'                  If StrComp(l_Folder.FolderId, l_Perspective.ID_EDITOR_AREA) = 0 Then
'                     RaiseEvent ActivateEditor(ActiveEditor)
'                  Else
'                     RaiseEvent ActivateView(ViewId)
'                  End If
                    
               Else
               
                  With l_ucFolder
                     .ShowView ViewId
                     .Active = True
                     .Refresh
                  End With
               
               End If
               
            End If
                
         Next f
         
      End If
   End With
   
   Set l_Perspective = Nothing
   
End Sub


' Show the previous view of the active folder.
Public Sub PreviousView()
   
   Dim i As Long
   Dim l_Perspective As Perspective
   Dim l_View As View
   Dim l_ucFolder As ucFolder
   
   Set l_Perspective = ActivePerspective
   
   Set l_ucFolder = Controls.Item(ActiveFolderId)
   
   With l_ucFolder.Views
      For i = 0 To .Count
         Set l_View = .Item(i)
                  
         If StrComp(l_View.ViewId, ActiveViewId) = 0 Then
            
            If i > 0 Then
               Set l_View = .Item(i - 1)
               Me.ShowView l_View.ViewId
               Exit For
               
            End If
         End If
            
      Next i
   End With
           
End Sub

' Show the next view of the active folder.
Public Sub NextView()
   
   Dim i As Long
   Dim l_Perspective As Perspective
   Dim l_View As View
   Dim l_ucFolder As ucFolder
   
   Set l_Perspective = ActivePerspective
   
   Set l_ucFolder = Controls.Item(ActiveFolderId)
   
   With l_ucFolder.Views
      For i = 0 To .Count
         Set l_View = .Item(i)
                  
         If StrComp(l_View.ViewId, ActiveViewId) = 0 Then
            
            If i < .Count Then
               Set l_View = .Item(i + 1)
               Me.ShowView l_View.ViewId
               Exit For
               
            End If
         End If
            
      Next i
   End With
           
End Sub

' Brings an editor to the top and activates it.
'
' @Editor The editor to show.
Public Sub ShowEditor(ByRef Editor As Object)
   
   Dim v As Long
   
   Dim l_ucFolder As ucFolder
   Dim l_Perspective As Perspective
   Dim l_Folder As Folder
   Dim l_View As View
   
   Set l_Perspective = ActivePerspective
   
   With l_Perspective.Folders
      
      If Not .IsEmpty Then
         
         ' ---------------------------------------------------------------------
         ' View is not visible
         ' ---------------------------------------------------------------------
         Set l_ucFolder = Controls.Item(l_Perspective.ID_EDITOR_AREA)
         
         For v = 0 To l_ucFolder.Views.Count
             Set l_View = l_ucFolder.Views.Item(v)
             
             If l_View.View.hwnd = Editor.hwnd Then
                RaiseEvent ActivateEditor(Editor)
                ShowView l_View.ViewId
             End If
         Next v
         
      End If
   End With
   
   Set l_Perspective = Nothing
   Set l_Folder = Nothing
   Set l_View = Nothing
   Set l_ucFolder = Nothing
   
End Sub

' Show the previous editor.
Public Sub PreviousEditor()
   
   Dim i As Long
   Dim Editor As Object
   Dim l_Perspective As Perspective
   Dim l_View As View
   Dim l_ucFolder As ucFolder
   
   Set l_Perspective = ActivePerspective
   
   Set l_ucFolder = Controls.Item(l_Perspective.ID_EDITOR_AREA)
   
   With l_ucFolder.Views
      For i = 0 To .Count
         Set l_View = .Item(i)
                  
         If StrComp(l_View.ViewId, ActiveViewId) = 0 Then
            
            If i > 0 Then
               Set l_View = .Item(i - 1)
               Set Editor = l_View.View
               Me.ShowEditor Editor
               Exit For
               
            End If
         End If
            
      Next i
   End With
           
End Sub

' Show the next editor.
Public Sub NextEditor()
   
   Dim i As Long
   Dim Editor As Object
   Dim l_Perspective As Perspective
   Dim l_View As View
   Dim l_ucFolder As ucFolder
   
   Set l_Perspective = ActivePerspective
   
   Set l_ucFolder = Controls.Item(l_Perspective.ID_EDITOR_AREA)
   
   With l_ucFolder.Views
      For i = 0 To .Count
         Set l_View = .Item(i)
                  
         If StrComp(l_View.ViewId, ActiveViewId) = 0 Then
            
            If i < .Count Then
               Set l_View = .Item(i + 1)
               Set Editor = l_View.View
               Me.ShowEditor Editor
               Exit For
               
            End If
         End If
            
      Next i
   End With
End Sub

' Open a new editor in the editor area.
'
' @EditorInput An editor input implementing the IEditorInput. You can use the implemented EditoFileInput class.
' @Editor The editor instance (normally a simple Visual Basic form).
Public Sub OpenEditor(ByRef Editor As Object)
   
   Static Idx As Long
   Dim l_EditorId As String
   Dim l_Perspective As Perspective
   Dim l_View As View
   Dim l_Editor As Form
   Dim l_ucFolder As ucFolder
   Dim l_Folder As Folder
   
   Idx = Idx + 1
   
   l_EditorId = "FolderId_" & Idx
   
   Set l_Perspective = ActivePerspective
   Set l_ucFolder = Controls.Item(l_Perspective.ID_EDITOR_AREA)
   Set l_Folder = l_Perspective.Folders.Item(l_Perspective.ID_EDITOR_AREA)
   
   Set l_View = New View
       
   With l_Folder
   
      With l_View
         .ViewId = l_EditorId
         Set .View = Editor
      End With
         
      m_Views.Add l_View.ViewId, l_View
      .AddView l_EditorId
               
      l_ucFolder.AddEditor l_View, True
      l_ucFolder.Refresh
      
      RaiseEvent OpenEditor(Editor)
      
   End With
   
   Set l_Perspective = Nothing
   Set l_Editor = Nothing
   Set l_Folder = Nothing
   
End Sub

' Call this function if you unload the main window. Then all views and editor
' will be unloaded automatically.
Public Sub Terminate()
   
   On Error Resume Next
   
   Dim i As Long
   Dim l_View As View
   
   With m_Views
      For i = 0 To .Count
         Set l_View = .Item(i)
         
         If Not l_View.View Is Nothing Then
            Unload l_View.View
            Set l_View.View = Nothing
         End If
      Next i
   End With
   
   
   With m_Editors
      For i = 0 To .Count
         Set l_View = .Item(i)
         
         If Not l_View.View Is Nothing Then
            Unload l_View.View
            Set l_View.View = Nothing
         End If
         
      Next i
   End With
   
   Set m_MagneticWnd = Nothing
   
End Sub

' Returns the current color scheme.
'
' @ColorScheme The current color scheme.
Public Property Get ColorScheme() As vbColorScheme
    ColorScheme = m_ColorScheme
End Property

' Set a new color scheme by a vbColorScheme constant. If vbCustomScheme is set
' the color properties of the perspective control will be used.
'
' @New_ColorScheme The new color scheme by vbColorScheme constant.
Public Property Let ColorScheme(ByVal New_ColorScheme As vbColorScheme)
    
   If m_ColorScheme <> New_ColorScheme Then
      m_ColorScheme = New_ColorScheme
    
      Select Case Me.ColorScheme
         Case VbWindowsXPScheme:      SetScheme New SchemeWinXP
         Case VbWindowsVistaScheme:       SetScheme New SchemeLonghorn
         Case VbOffice2003Scheme:     SetScheme New SchemeOffice2003
         Case VbEclipse3Scheme:       SetScheme New SchemeEclipse3
         Case VbVS2005Scheme:         SetScheme New SchemeVS2005
         Case Else:
                                        Dim l_Scheme As SchemeCustom
                                        Set l_Scheme = New SchemeCustom
                                        
                                        With l_Scheme
                                           .ActiveCaptionForeColor = Me.ActiveCaptionForeColor
                                           .ActiveCaptionGradient1 = Me.ActiveCaptionGradient1
                                           .ActiveCaptionGradient2 = Me.ActiveCaptionGradient2
                                           .InactiveCaptionForeColor = Me.InactiveCaptionForeColor
                                           .InactiveCaptionGradient1 = Me.InactiveCaptionGradient1
                                           .InactiveCaptionGradient2 = Me.InactiveCaptionGradient2
                                           .FocusTabForeColor = Me.FocusTabForeColor
                                           .FocusTabGradient1 = Me.FocusTabGradient1
                                           .FocusTabGradient2 = Me.FocusTabGradient2
                                           .FocusTabGradientAngle = Me.FocusTabGradientAngle
                                           .ActiveTabForeColor = Me.ActiveTabForeColor
                                           .ActiveTabGradient1 = Me.ActiveTabGradient1
                                           .ActiveTabGradient2 = Me.ActiveTabGradient2
                                           .ActiveTabGradientAngle = Me.ActiveTabGradientAngle
                                           .BackColor = Me.BackColor
                                           .FrameColor = Me.FrameColor
                                           .FrameWidth = Me.FrameWidth
                                           .EditorAreaBackColor = Me.EditorAreaBackColor
                                           .ViewCaptions = Me.ViewCaptions
                                           .CaptionStyle = Me.CaptionStyle
                                        End With
                                        
                                        SetScheme l_Scheme
      End Select
      
      PropertyChanged "ColorScheme"
       
   End If
    
End Property

Public Sub Load(ByVal FilePath As String)

   Dim Doc As MSXML2.DOMDocument
   Dim Node As MSXML2.IXMLDOMNode
   Dim pList As MSXML2.IXMLDOMNodeList
   Dim fList As MSXML2.IXMLDOMNodeList
   Dim vList As MSXML2.IXMLDOMNodeList
   
   Dim l_Perspective As Perspective
   Dim l_Folder As Folder
   Dim l_View As View
   
   Dim strId As String
   Dim lRel As vbRelationship
   Dim lRatio As Double
   Dim strRefId As String
   Dim strActiveId As String
   Dim blnEditorArea As Boolean
   Dim Left As Long
   Dim Right As Long
   Dim Top As Long
   Dim Bottom As Long
   
   Dim p As Long
   Dim f As Long
   Dim v As Long

   Set Doc = New MSXML2.DOMDocument
   Doc.Load FilePath
   
   ' ------------------------------------------------------------------------
   ' Perspectives
   ' ------------------------------------------------------------------------
   Set pList = Doc.selectNodes("Perspectives/Perspective")
   
   For p = 0 To pList.Length - 1
      
      Set Node = pList.Item(p)
      
      ' ------------------------------------------------------------------------
      ' Add Perspective
      ' ------------------------------------------------------------------------
      strId = Node.Attributes.getNamedItem("Id").nodeValue
      strActiveId = Node.Attributes.getNamedItem("ActiveViewId").Text
      blnEditorArea = StrComp(Node.Attributes.getNamedItem("EditorArea").Text, "show") = 0
         
      Set l_Perspective = Me.AddPerspective(strId)
          l_Perspective.ActiveViewId = strActiveId
          l_Perspective.EditorAreaVisible = blnEditorArea
          
      ' ------------------------------------------------------------------------
      ' Folders
      ' ------------------------------------------------------------------------
      Set fList = Node.selectNodes("Folder")
         
      For f = 0 To fList.Length - 1
         Set Node = fList.Item(f)
         
         ' ---------------------------------------------------------------------
         ' Add Folder
         ' ---------------------------------------------------------------------
         On Error Resume Next
         strId = Node.Attributes.getNamedItem("Id").nodeValue
         strActiveId = Node.Attributes.getNamedItem("ActiveViewId").Text
         strRefId = Node.selectSingleNode("RefId").Text
         lRatio = Node.selectSingleNode("Ratio").Text
         lRel = Node.selectSingleNode("Relationship").Text
         Left = Node.selectSingleNode("Left").Text
         Right = Node.selectSingleNode("Right").Text
         Top = Node.selectSingleNode("Top").Text
         Bottom = Node.selectSingleNode("Bottom").Text
         
         If Left < 0 Then
            Right = Right + (Left * -1)
            Left = 0
         End If
         
         Set l_Folder = l_Perspective.AddFolder(strId, lRel, lRatio, strRefId)
         With l_Folder
             .ActiveViewId = strActiveId
             With .Position
                .Left = Left
                .Right = Right
                .Top = Top
                .Bottom = Bottom
             End With
         End With
         ' ---------------------------------------------------------------------
         ' Views
         ' ---------------------------------------------------------------------
         Set vList = Node.selectNodes("Views/View")
         
         For v = 0 To vList.Length - 1
            Set Node = vList.Item(v)
            
            strId = Node.Attributes.getNamedItem("Id").nodeValue
            
            ' ------------------------------------------------------------------
            ' Add View
            ' ------------------------------------------------------------------
            l_Folder.AddView strId
            
         Next v
         
      Next f
         
   Next p
   
   Set Node = Doc.selectSingleNode("Perspectives")
   
   If Not Node Is Nothing Then
      Me.ShowPerspective Node.Attributes.getNamedItem("ActivePerspectiveId").nodeValue
   End If
   
   Set l_Perspective = Nothing
   Set l_Folder = Nothing
   Set l_View = Nothing

End Sub

Public Sub Save(ByVal FilePath As String)

   Dim sb As String
   Dim Rc As RECT
   Dim l_Perspective As Perspective
   Dim l_Folder As Folder
   Dim l_ucFolder As ucFolder
   Dim l_View As View
   Dim p As Long
   Dim f As Long
   Dim v As Long

   Set l_Perspective = New Perspective

   sb = vbNullString
   sb = sb & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbNewLine

   With Perspectives
      
      sb = sb & "<Perspectives ActivePerspectiveId=""" & Me.ActivePerspectiveId & """>" & vbNewLine
      
      For p = 0 To .Count
         With .Item(p)
            sb = sb & "   <Perspective Id=""" & .PerspectiveId & """ ActiveViewId=""" & .ActiveViewId & """ EditorArea=""" & IIf(.EditorAreaVisible, "show", "hide") & """>" & vbNewLine

            With .Folders
               For f = 0 To .Count
                  With .Item(f)
                  
                     If StrComp(.FolderId, l_Perspective.ID_EDITOR_AREA) <> 0 Then
                     
                     sb = sb & "      <Folder Id=""" & .FolderId & """ ActiveViewId=""" & .ActiveViewId & """>" & vbNewLine
                     sb = sb & "         <Ratio>" & .Ratio & "</Ratio>" & vbNewLine
                     sb = sb & "         <RefId>" & .RefId & "</RefId>" & vbNewLine
                     sb = sb & "         <Relationship>" & .Relationship & "</Relationship>" & vbNewLine
                     
                     If .Relationship = vbRelFloating Then
                        
                        On Error Resume Next
                        Set l_ucFolder = Nothing
                        Set l_ucFolder = Controls.Item(.FolderId)
                        On Error GoTo 0
                        
                        If Not l_ucFolder Is Nothing Then
                           GetWindowRect l_ucFolder.hwnd, Rc
                           With Rc
                           'With .Position
                              sb = sb & "         <Left>" & .Left - 4 & "</Left>" & vbNewLine
                              sb = sb & "         <Right>" & .Right - .Left + 7 & "</Right>" & vbNewLine
                              sb = sb & "         <Top>" & .Top - 4 & "</Top>" & vbNewLine
                              sb = sb & "         <Bottom>" & .Bottom - .Top + 7 & "</Bottom>" & vbNewLine
                           End With
                        End If
                     End If
                     
                     sb = sb & "         <Views>" & vbNewLine
                     With .Views
                        For v = 0 To .Count
                           sb = sb & "            <View Id=""" & .Item(v) & """ />" & vbNewLine
                        Next v
                     End With
                     sb = sb & "         </Views>" & vbNewLine
                     sb = sb & "      </Folder>" & vbNewLine
                     End If
                  End With
               Next f

            End With

            sb = sb & "   </Perspective>" & vbNewLine
         End With
      Next p
   End With
   
   sb = sb & "</Perspectives>" & vbNewLine
   sb = sb & "<!--"
   sb = sb & "     Version " & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine
   sb = sb & vbNewLine
   sb = sb & "     Copyright (c) 2004 - 2005, AB-Software" & vbNewLine
   sb = sb & "     All rights reserved. -->" & vbNewLine
   
   f = FreeFile
   
   Open FilePath For Output As #f
      Print #f, sb
   Close #f
   
End Sub

Public Property Get CaptionStyle() As vbCaptionStyle
    CaptionStyle = m_CaptionStyle
End Property
Public Property Let CaptionStyle(ByVal New_CaptionStyle As vbCaptionStyle)
    m_CaptionStyle = New_CaptionStyle
    PropertyChanged "CaptionStyle"
End Property

Public Property Get ActiveCaptionForeColor() As OLE_COLOR
    ActiveCaptionForeColor = m_ActiveCaptionForeColor
End Property
Public Property Let ActiveCaptionForeColor(ByVal New_ActiveCaptionForeColor As OLE_COLOR)
    m_ActiveCaptionForeColor = New_ActiveCaptionForeColor
    PropertyChanged "ActiveCaptionForeColor"
End Property

Public Property Get ActiveCaptionGradient1() As OLE_COLOR
    ActiveCaptionGradient1 = m_ActiveCaptionGradient1
End Property
Public Property Let ActiveCaptionGradient1(ByVal New_ActiveCaptionGradient1 As OLE_COLOR)
    m_ActiveCaptionGradient1 = New_ActiveCaptionGradient1
    PropertyChanged "ActiveCaptionGradient1"
End Property

Public Property Get ActiveCaptionGradient2() As OLE_COLOR
    ActiveCaptionGradient2 = m_ActiveCaptionGradient2
End Property
Public Property Let ActiveCaptionGradient2(ByVal New_ActiveCaptionGradient2 As OLE_COLOR)
    m_ActiveCaptionGradient2 = New_ActiveCaptionGradient2
    PropertyChanged "ActiveCaptionGradient2"
End Property

Public Property Get InactiveCaptionForeColor() As OLE_COLOR
    InactiveCaptionForeColor = m_InactiveCaptionForeColor
End Property
Public Property Let InactiveCaptionForeColor(ByVal New_InactiveCaptionForeColor As OLE_COLOR)
    m_InactiveCaptionForeColor = New_InactiveCaptionForeColor
    PropertyChanged "InactiveCaptionForeColor"
End Property

Public Property Get InactiveCaptionGradient1() As OLE_COLOR
    InactiveCaptionGradient1 = m_InactiveCaptionGradient1
End Property
Public Property Let InactiveCaptionGradient1(ByVal New_InactiveCaptionGradient1 As OLE_COLOR)
    m_InactiveCaptionGradient1 = New_InactiveCaptionGradient1
    PropertyChanged "InactiveCaptionGradient1"
End Property

Public Property Get InactiveCaptionGradient2() As OLE_COLOR
    InactiveCaptionGradient2 = m_InactiveCaptionGradient2
End Property
Public Property Let InactiveCaptionGradient2(ByVal New_InactiveCaptionGradient2 As OLE_COLOR)
    m_InactiveCaptionGradient2 = New_InactiveCaptionGradient2
    PropertyChanged "InactiveCaptionGradient2"
End Property

Public Property Get FocusTabForeColor() As OLE_COLOR
    FocusTabForeColor = m_FocusTabForeColor
End Property
Public Property Let FocusTabForeColor(ByVal New_FocusTabForeColor As OLE_COLOR)
    m_FocusTabForeColor = New_FocusTabForeColor
    PropertyChanged "FocusTabForeColor"
End Property

Public Property Get FocusTabGradient1() As OLE_COLOR
    FocusTabGradient1 = m_FocusTabGradient1
End Property
Public Property Let FocusTabGradient1(ByVal New_FocusTabGradient1 As OLE_COLOR)
    m_FocusTabGradient1 = New_FocusTabGradient1
    PropertyChanged "FocusTabGradient1"
End Property

Public Property Get FocusTabGradient2() As OLE_COLOR
    FocusTabGradient2 = m_FocusTabGradient2
End Property
Public Property Let FocusTabGradient2(ByVal New_FocusTabGradient2 As OLE_COLOR)
    m_FocusTabGradient2 = New_FocusTabGradient2
    PropertyChanged "FocusTabGradient2"
End Property

Public Property Get FocusTabGradientAngle() As Long
Attribute FocusTabGradientAngle.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    FocusTabGradientAngle = m_FocusTabGradientAngle
End Property
Public Property Let FocusTabGradientAngle(ByVal New_FocusTabGradientAngle As Long)
   If New_FocusTabGradientAngle >= 0 And New_FocusTabGradientAngle <= 360 Then
      m_FocusTabGradientAngle = New_FocusTabGradientAngle
      PropertyChanged "FocusTabGradientAngle"
   Else
      MsgBox "The angle of the active tabs can only be between 0 and 360!", vbCritical + vbOKOnly, "Active Tab Gradient Angle"
   End If
End Property

Public Property Get ActiveTabForeColor() As OLE_COLOR
   ActiveTabForeColor = m_ActiveTabForeColor
End Property
Public Property Let ActiveTabForeColor(ByVal New_ActiveTabForeColor As OLE_COLOR)
   m_ActiveTabForeColor = New_ActiveTabForeColor
   PropertyChanged "ActiveTabForeColor"
End Property

Public Property Get ActiveTabGradient1() As OLE_COLOR
   ActiveTabGradient1 = m_ActiveTabGradient1
End Property
Public Property Let ActiveTabGradient1(ByVal New_ActiveTabGradient1 As OLE_COLOR)
   m_ActiveTabGradient1 = New_ActiveTabGradient1
   PropertyChanged "ActiveTabGradient1"
End Property

Public Property Get ActiveTabGradient2() As OLE_COLOR
   ActiveTabGradient2 = m_ActiveTabGradient2
End Property
Public Property Let ActiveTabGradient2(ByVal New_ActiveTabGradient2 As OLE_COLOR)
   m_ActiveTabGradient2 = New_ActiveTabGradient2
   PropertyChanged "ActiveTabGradient2"
End Property

Public Property Get ActiveTabGradientAngle() As Long
Attribute ActiveTabGradientAngle.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
   ActiveTabGradientAngle = m_ActiveTabGradientAngle
End Property
Public Property Let ActiveTabGradientAngle(ByVal New_ActiveTabGradientAngle As Long)
   If New_ActiveTabGradientAngle >= 0 And New_ActiveTabGradientAngle <= 360 Then
      m_ActiveTabGradientAngle = New_ActiveTabGradientAngle
      PropertyChanged "ActiveTabGradientAngle"
   Else
      MsgBox "The angle of the inactive tabs can only be between 0 and 360!", vbCritical + vbOKOnly, "Active Tab Gradient Angle"
   End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get FrameColor() As OLE_COLOR
    FrameColor = m_FrameColor
End Property
Public Property Let FrameColor(ByVal New_FrameColor As OLE_COLOR)
    m_FrameColor = New_FrameColor
    PropertyChanged "FrameColor"
End Property

Public Property Get FrameWidth() As Long
Attribute FrameWidth.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    FrameWidth = m_FrameWidth
End Property
Public Property Let FrameWidth(ByVal New_FrameWidth As Long)
    m_FrameWidth = New_FrameWidth
    PropertyChanged "FrameWidth"
End Property

Public Property Get EditorAreaBackColor() As OLE_COLOR
    EditorAreaBackColor = m_EditorAreaBackColor
End Property
Public Property Let EditorAreaBackColor(ByVal New_EditorAreaBackColor As OLE_COLOR)
    m_EditorAreaBackColor = New_EditorAreaBackColor
    PropertyChanged "EditorAreaBackColor"
End Property

Public Property Get IsEditorAreaVisible() As Boolean
   If Len(m_ActivePerspectiveId) > 1 Then
      IsEditorAreaVisible = ActivePerspective.EditorAreaVisible
   Else
      IsEditorAreaVisible = False
   End If
End Property

Public Property Get ViewCaptions() As Boolean
Attribute ViewCaptions.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    ViewCaptions = m_ViewCaptions
End Property
Public Property Let ViewCaptions(ByVal New_ViewCaptions As Boolean)
    m_ViewCaptions = New_ViewCaptions
    PropertyChanged "ViewCaptions"
End Property

Public Property Get ViewCaptionIcons() As Boolean
Attribute ViewCaptionIcons.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    ViewCaptionIcons = m_ViewCaptionIcons
End Property
Public Property Let ViewCaptionIcons(ByVal New_ViewCaptionIcons As Boolean)
    m_ViewCaptionIcons = New_ViewCaptionIcons
    PropertyChanged "ViewCaptionIcons"
End Property

Public Property Get Magnetic() As Boolean
    Magnetic = m_Magnetic
End Property
Public Property Let Magnetic(ByVal New_Magnetic As Boolean)
    m_Magnetic = New_Magnetic
    PropertyChanged "Magnetic"
End Property

' Initialize properties of the user control
Private Sub UserControl_InitProperties()
    m_ActiveCaptionForeColor = m_def_ActiveCaptionForeColor
    m_ActiveCaptionGradient1 = m_def_ActiveCaptionGradient1
    m_ActiveCaptionGradient2 = m_def_ActiveCaptionGradient2
    
    m_InactiveCaptionForeColor = m_def_InactiveCaptionForeColor
    m_InactiveCaptionGradient1 = m_def_InactiveCaptionGradient1
    m_InactiveCaptionGradient2 = m_def_InactiveCaptionGradient2
    
    m_FocusTabForeColor = m_def_FocusTabForeColor
    m_FocusTabGradient1 = m_def_FocusTabGradient1
    m_FocusTabGradient2 = m_def_FocusTabGradient2
    m_FocusTabGradientAngle = m_def_FocusTabGradientAngle
        
    m_ActiveTabForeColor = m_def_ActiveTabForeColor
    m_ActiveTabGradient1 = m_def_ActiveTabGradient1
    m_ActiveTabGradient2 = m_def_ActiveTabGradient2
    m_ActiveTabGradientAngle = m_def_ActiveTabGradientAngle
        
    m_BackColor = m_def_BackColor
    m_FrameColor = m_def_FrameColor
    m_FrameWidth = m_def_FrameWidth
    m_EditorAreaBackColor = m_def_EditorAreaBackColor
    m_ViewCaptions = m_def_ViewCaptions
    
    ColorScheme = m_def_ColorScheme
    CaptionStyle = m_def_CaptionStyle
    m_Magnetic = m_def_Magnetic
End Sub

' Load property values from memory
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_ActiveCaptionForeColor = PropBag.ReadProperty("ActiveCaptionForeColor", m_def_ActiveCaptionForeColor)
    m_ActiveCaptionGradient1 = PropBag.ReadProperty("ActiveCaptionGradient1", m_def_ActiveCaptionGradient1)
    m_ActiveCaptionGradient2 = PropBag.ReadProperty("ActiveCaptionGradient2", m_def_ActiveCaptionGradient2)
    
    m_InactiveCaptionForeColor = PropBag.ReadProperty("InactiveCaptionForeColor", m_def_InactiveCaptionForeColor)
    m_InactiveCaptionGradient1 = PropBag.ReadProperty("InactiveCaptionGradient1", m_def_InactiveCaptionGradient1)
    m_InactiveCaptionGradient2 = PropBag.ReadProperty("InactiveCaptionGradient2", m_def_InactiveCaptionGradient2)
    
    m_FocusTabForeColor = PropBag.ReadProperty("FocusTabForeColor", m_def_FocusTabForeColor)
    m_FocusTabGradient1 = PropBag.ReadProperty("FocusTabGradient1", m_def_FocusTabGradient1)
    m_FocusTabGradient2 = PropBag.ReadProperty("FocusTabGradient2", m_def_FocusTabGradient2)
    m_FocusTabGradientAngle = PropBag.ReadProperty("FocusTabGradientAngle", m_def_FocusTabGradientAngle)
    
    m_ActiveTabForeColor = PropBag.ReadProperty("ActiveTabForeColor", m_def_ActiveTabForeColor)
    m_ActiveTabGradient1 = PropBag.ReadProperty("ActiveTabGradient1", m_def_ActiveTabGradient1)
    m_ActiveTabGradient2 = PropBag.ReadProperty("ActiveTabGradient2", m_def_ActiveTabGradient2)
    m_ActiveTabGradientAngle = PropBag.ReadProperty("ActiveTabGradientAngle", m_def_ActiveTabGradientAngle)
    
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_FrameColor = PropBag.ReadProperty("FrameColor", m_def_FrameColor)
    m_FrameWidth = PropBag.ReadProperty("FrameWidth", m_def_FrameWidth)
    m_EditorAreaBackColor = PropBag.ReadProperty("EditorAreaBackColor", m_def_EditorAreaBackColor)
    
    m_ViewCaptions = PropBag.ReadProperty("ViewCaptions", m_def_ViewCaptions)
    m_ViewCaptionIcons = PropBag.ReadProperty("ViewCaptionIcons", m_def_ViewCaptionIcons)
    
    ColorScheme = PropBag.ReadProperty("ColorScheme", m_def_ColorScheme)
    CaptionStyle = PropBag.ReadProperty("CaptionStyle", m_def_CaptionStyle)
    m_Magnetic = PropBag.ReadProperty("Magnetic", m_def_Magnetic)
End Sub

' Save property values to memory
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ActiveCaptionForeColor", m_ActiveCaptionForeColor, m_def_ActiveCaptionForeColor)
    Call PropBag.WriteProperty("ActiveCaptionGradient1", m_ActiveCaptionGradient1, m_def_ActiveCaptionGradient1)
    Call PropBag.WriteProperty("ActiveCaptionGradient2", m_ActiveCaptionGradient2, m_def_ActiveCaptionGradient2)
    
    Call PropBag.WriteProperty("InactiveCaptionForeColor", m_InactiveCaptionForeColor, m_def_InactiveCaptionForeColor)
    Call PropBag.WriteProperty("InactiveCaptionGradient1", m_InactiveCaptionGradient1, m_def_InactiveCaptionGradient1)
    Call PropBag.WriteProperty("InactiveCaptionGradient2", m_InactiveCaptionGradient2, m_def_InactiveCaptionGradient2)
    
    Call PropBag.WriteProperty("FocusTabForeColor", m_FocusTabForeColor, m_def_FocusTabForeColor)
    Call PropBag.WriteProperty("FocusTabGradient1", m_FocusTabGradient1, m_def_FocusTabGradient1)
    Call PropBag.WriteProperty("FocusTabGradient2", m_FocusTabGradient2, m_def_FocusTabGradient2)
    Call PropBag.WriteProperty("FocusTabGradientAngle", m_FocusTabGradientAngle, m_def_FocusTabGradientAngle)
    
    Call PropBag.WriteProperty("ActiveTabForeColor", m_ActiveTabForeColor, m_def_ActiveTabForeColor)
    Call PropBag.WriteProperty("ActiveTabGradient1", m_ActiveTabGradient1, m_def_ActiveTabGradient1)
    Call PropBag.WriteProperty("ActiveTabGradient2", m_ActiveTabGradient2, m_def_ActiveTabGradient2)
    Call PropBag.WriteProperty("ActiveTabGradientAngle", m_ActiveTabGradientAngle, m_def_ActiveTabGradientAngle)
    
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("FrameColor", m_FrameColor, m_def_FrameColor)
    Call PropBag.WriteProperty("FrameWidth", m_FrameWidth, m_def_FrameWidth)
    Call PropBag.WriteProperty("EditorAreaBackColor", m_EditorAreaBackColor, m_def_EditorAreaBackColor)
    
    Call PropBag.WriteProperty("ViewCaptions", m_ViewCaptions, m_def_ViewCaptions)
    Call PropBag.WriteProperty("ViewCaptionIcons", m_ViewCaptionIcons, m_def_ViewCaptionIcons)
    
    Call PropBag.WriteProperty("ColorScheme", m_ColorScheme, m_def_ColorScheme)
    Call PropBag.WriteProperty("CaptionStyle", m_CaptionStyle, m_def_CaptionStyle)
    Call PropBag.WriteProperty("Magnetic", m_Magnetic, m_def_Magnetic)
End Sub
