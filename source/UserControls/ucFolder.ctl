VERSION 5.00
Begin VB.UserControl ucFolder 
   BackColor       =   &H80000003&
   ClientHeight    =   6675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8865
   ScaleHeight     =   6675
   ScaleWidth      =   8865
   ToolboxBitmap   =   "ucFolder.ctx":0000
   Begin Aggressive.ucTabStrip ViewTabs 
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   8655
      _extentx        =   15266
      _extenty        =   582
   End
   Begin VB.PictureBox ViewArea 
      BorderStyle     =   0  'Kein
      Height          =   5895
      Left            =   120
      ScaleHeight     =   5895
      ScaleWidth      =   8655
      TabIndex        =   0
      Top             =   360
      Width           =   8655
   End
   Begin Aggressive.ucCaption ViewCaption 
      Height          =   250
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   8655
      _extentx        =   15266
      _extenty        =   503
   End
End
Attribute VB_Name = "ucFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_FolderId As String
Private m_LastRefFolderId As String
Private m_FolderViews As List ' The views of this folder
Private m_Active As Boolean ' State of this folder (true = active / false = inactive)
Private m_MaximizeAble As Boolean

Private m_TopPos As Long ' The top position
Private m_BottomPos As Long ' The bottom position
Private m_LeftPos As Long ' The left position
Private m_RightPos As Long ' The right position

Private Sub UserControl_Initialize()
   m_Active = True
   
   Set m_FolderViews = New List
End Sub

Private Sub UserControl_Terminate()
   Set m_FolderViews = Nothing
End Sub

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get Active() As Boolean
   Active = m_Active
End Property

Public Property Let Active(ByVal NewActive As Boolean)
   
   If m_Active <> NewActive Then
      
      m_Active = NewActive
      
      With ViewCaption
         .Active = NewActive
         .Refresh
      End With
      
      ViewArea_Resize
      
      ViewTabs.Active = NewActive
      Refresh
   End If
   
End Property

Public Property Get ActiveViewId() As String
   ActiveViewId = ViewTabs.ActiveViewId
End Property

Public Property Get Views() As List
   Set Views = m_FolderViews
End Property

Public Property Let FolderId(ByVal NewFolderId As String)
   m_FolderId = NewFolderId
End Property

Public Property Get FolderId() As String
   FolderId = m_FolderId
End Property

Public Property Let LastRefFolderId(ByVal NewLastRefFolderId As String)
   m_LastRefFolderId = NewLastRefFolderId
End Property

Public Property Get LastRefFolderId() As String
   LastRefFolderId = m_LastRefFolderId
End Property

Public Property Let TopPos(ByVal NewTopPos As Long)
   m_TopPos = NewTopPos
End Property

Public Property Get TopPos() As Long
   TopPos = m_TopPos
End Property

Public Property Let BottomPos(ByVal NewBottomPos As Long)
   m_BottomPos = NewBottomPos
End Property

Public Property Get BottomPos() As Long
   BottomPos = m_BottomPos
End Property

Public Property Let LeftPos(ByVal NewLeftPos As Long)
   m_LeftPos = NewLeftPos
End Property

Public Property Get LeftPos() As Long
   LeftPos = m_LeftPos
End Property

Public Property Let RightPos(ByVal NewRightPos As Long)
   m_RightPos = NewRightPos
End Property

Public Property Get RightPos() As Long
   RightPos = m_RightPos
End Property

Public Property Let MaximizeAble(ByVal NewMaximizeAble As Boolean)
   m_MaximizeAble = NewMaximizeAble
   ViewCaption.MaximizeButton = m_MaximizeAble
End Property

Public Property Get MaximizeAble() As Boolean
   MaximizeAble = m_MaximizeAble
End Property

Private Property Get IsFloating() As Boolean
   
   Dim i As Long
   Dim frm As Variant
   Dim pHwnd As Long
   
   ' Get the window handle of the folders parent
   pHwnd = GetParent(Me.hWnd)
   
   ' Iterate all floating windows and campare window handles
   With m_Windows
      For i = 0 To .Count
         Set frm = .Item(i)
         ' if window handles are equal -> folder is floating
         If frm.hWnd = pHwnd Then
            IsFloating = True
            Exit Property
         End If
      Next i
   End With
      
   IsFloating = False
   
End Property
Private Property Get ShowCaption() As Boolean
   
   ShowCaption = m_Scheme.ViewCaptions
   
   If ShowCaption Then
      ShowCaption = Not IsEditorArea
   End If
   
   If Not ShowCaption Then
      ShowCaption = IsFloating
   End If
   
End Property

' Returns true if this folder is the editor area.
'
' @IsEditorArea True if this folder is the editor area.
Private Function IsEditorArea() As Boolean
   
   Dim l_Perspecive As Perspective
   Set l_Perspecive = New Perspective
   
   IsEditorArea = CBool(StrComp(Me.FolderId, l_Perspecive.ID_EDITOR_AREA, vbBinaryCompare) = 0)

   Set l_Perspecive = Nothing

End Function

' Add a view to the folder.
'
' @View The view to add to this folder.
' @Activate True the activate the view.
Public Sub AddView(ByRef View As View, Optional ByVal Activate As Boolean = True)
   
   Dim l_ShowCaption As Boolean
   
   l_ShowCaption = ShowCaption
   
   m_FolderViews.Add View.ViewId, View
   
   With View.View
      SetParent .hWnd, ViewArea.hWnd
      SetWindowStyle .hWnd, VbNone
   End With
    
   If ViewCaption.Visible Then
      ViewCaption.Caption = View.View.Caption
   End If
   
   With ViewTabs
      If IsEditorArea Then
         .Orientation = VbOrientationTop
      Else
         .Orientation = VbOrientationBottom
      End If
      
      .Add View.ViewId, View.View.Caption, View.View.Icon, View.View.Caption
      .AutoHideButtons = l_ShowCaption
      .CloseButton = l_ShowCaption
      .NextButton = True
      .PrevButton = True
      .Visible = (m_FolderViews.Count > 0 Or Not l_ShowCaption)
   End With
                  
   With View.View
      modSubClass.UnHook .hWnd
      modSubClass.Hook .hWnd
   End With
         
   If Activate Then
      ShowView View.ViewId
   End If
      
End Sub

' Add a editor to the editor area.
'
' @Editor The editor to add to this folder.
' @EditorInput The input to open the editor.
' @Activate True the activate the editor.
Public Sub AddEditor(ByRef Editor As View, Optional ByVal Activate As Boolean = True)
      
'   Dim l_EditorPart As IEditorPart
'
'   Set l_EditorPart = Editor.View
'       l_EditorPart.Init EditorInput
       
   m_FolderViews.Add Editor.ViewId, Editor
   
   With Editor.View
      SetWindowStyle .hWnd, VbNone   ' Set new window style (remove caption & border)
      SetParent .hWnd, ViewArea.hWnd ' Set a new parent for the editor
      
      m_Editors.Add Editor.ViewId, Editor
      
      modSubClass.UnHook .hWnd
      modSubClass.Hook .hWnd
   End With
   
  ' ViewCaption.Visible = True 'False
   
   With ViewTabs
      .Orientation = VbOrientationTop
      
      .Add Editor.ViewId, Editor.View.Caption, Editor.View.Icon, Editor.View.Tag
      
      ' Don't hide the buttons if all tabs are visible
      .AutoHideButtons = False
      .CloseButton = True
      .NextButton = True
      .PrevButton = True
      
      .Visible = True
   End With
         
   ' Activate the editor
   If Activate Then
      ShowView Editor.ViewId
   End If
   
End Sub

' Returns true if this folder contains the view; false otherwise.
'
' @ViewId The id of the view.
'
' @ContainsView True if contains the view; false otherwise.
Public Function ContainsView(ByVal ViewId As String) As Boolean
   ContainsView = ViewTabs.Tabs.Contains(ViewId)
End Function

' Removes a view from this folder and shows the next view.
'
' @ViewId The id of the view that should be removed.
Public Sub RemoveView(ByVal ViewId As String)
   
   m_FolderViews.Remove ViewId
      
   With ViewTabs
      .Remove ViewId
      .Visible = (.Count > 0)
   End With
   
   If ViewTabs.Count > -1 Then
      ShowView ViewTabs.ActiveViewId
   End If
   
   If m_FolderViews.IsEmpty Then
      Parent.EventRaise "RemoveFolder", FolderId, ViewId
   End If
      
   UserControl_Resize
      
End Sub

' Shows / activates a view.
'
' @ViewId The id of the view to show.
Public Sub ShowView(ByVal ViewId As String)

   Dim l_View As View
   Dim l_Tab As ucTab
   Dim i As Long
   
   With m_FolderViews
      For i = 0 To .Count
         
         Set l_View = .Item(i)
         
         If StrComp(l_View.ViewId, ViewId, vbBinaryCompare) = 0 Then
            
            ViewCaption.Caption = l_View.View.Caption
            
            If m_Scheme.ViewCaptionIcons Then
               Set ViewCaption.Icon = l_View.View.Icon
            End If
            
            If ViewTabs.Tabs.Contains(ViewId) Then
               Set l_Tab = ViewTabs.Tabs.Item(ViewId)
                   l_Tab.Caption = l_View.View.Caption
            End If
            
            l_View.View.Visible = True
         Else
            l_View.View.Visible = False
         End If
         
      Next i
   End With
   
   ViewTabs.Show ViewId
   
   ViewArea_Resize
   
End Sub
' Refreshs a view.
'
' @ViewId The id of the view to refresh.
Public Sub RefreshView(ByVal ViewId As String)

   Dim l_View As View
   Dim l_Tab As ucTab
   Dim i As Long
   
   With m_FolderViews
      For i = 0 To .Count
         
         Set l_View = .Item(i)
         
         If StrComp(l_View.ViewId, ViewId, vbBinaryCompare) = 0 Then
            
            ViewCaption.Caption = l_View.View.Caption
            
            If m_Scheme.ViewCaptionIcons Then
               Set ViewCaption.Icon = l_View.View.Icon
            End If
            
            If ViewTabs.Tabs.Contains(ViewId) Then
               Set l_Tab = ViewTabs.Tabs.Item(ViewId)
                   l_Tab.Caption = l_View.View.Caption
            End If
            
            'l_View.View.Visible = True
         'Else
            'l_View.View.Visible = False
            
            Exit For
            
         End If
         
      Next i
   End With
   
   'ViewTabs.Show ViewId
   
   ViewArea_Resize
   
End Sub
Private Sub UserControl_Resize()

   On Error Resume Next
      
   Dim l_ShowCaption As Boolean
   
   l_ShowCaption = ShowCaption
      
   ' Move caption (if visible)
   ViewCaption.Visible = l_ShowCaption
      
   If ViewCaption.Visible Then
      ViewCaption.Move 20, 20, ScaleWidth - 30, ViewCaption.Height
   Else
      ViewTabs.Move 20, 20, ScaleWidth - 30, ViewTabs.Height
   End If
   
   ' Move tabs (if visible)
   ViewTabs.Visible = (IsEditorArea Or ViewTabs.Count > 0 Or Not l_ShowCaption)
   
   If ViewTabs.Visible Then
      If IsEditorArea Then
         If Not m_FolderViews.IsEmpty Then
            ViewTabs.Move 20, 20, ScaleWidth - 30, ViewTabs.Height
            ViewArea.Move 20, ViewTabs.Height + 20, ScaleWidth - 30, ScaleHeight - ViewTabs.Height - 30
            
            If Active Then
               ViewArea.BackColor = vbButtonFace
            Else
               ViewArea.BackColor = vbButtonFace
            End If
         Else
            ViewTabs.Visible = False
            ViewArea.Move 20, 20, ScaleWidth - 30, ScaleHeight - 30
            ViewArea.BackColor = m_Scheme.EditorAreaBackColor
         End If
      Else
         If ViewCaption.Visible Then
            ViewArea.Move 20, ViewCaption.Height, ScaleWidth - 30, ScaleHeight - ViewCaption.Height - ViewTabs.Height - 20
         Else
            ViewArea.Move 20, 20, ScaleWidth - 30, ScaleHeight - ViewTabs.Height - 30
         End If
         ViewTabs.Move 20, ViewArea.Top + ViewArea.ScaleHeight, ScaleWidth - 30, ViewTabs.Height
      End If
   Else
      ViewArea.Move 20, ViewCaption.Height, ScaleWidth - 30, ScaleHeight - ViewCaption.Height - 20
   End If
   
'   UserControl.Cls
   
'   If IsEditorArea Then
'      UserControl.Line (1, 1)-(ScaleWidth - 10, ScaleHeight - 10), vbWhite, B
'      UserControl.Line (1, 1)-(1, ScaleHeight - 10), vbApplicationWorkspace, B
'      UserControl.Line (1, 1)-(ScaleWidth - 10, 1), vbApplicationWorkspace, B
'   Else
      UserControl.BackColor = m_Scheme.FrameColor
'   End If

   
End Sub

Public Sub Refresh()
   
   Dim l_View As View
   
   ViewArea_Resize
   
   UserControl_Resize
   
   ' Refresh the view (form)
   If Len(ActiveViewId) > 0 Then
      Set l_View = m_Views.Item(ActiveViewId)
          l_View.View.Refresh
   End If
       
   ' Refresh caption
   ViewCaption.Refresh
   
   ' Show / hide tab navigation buttons & refresh the tabs.
   ViewTabs.AutoHideButtons = ShowCaption And Not IsFloating
   ViewTabs.CloseButton = Not ShowCaption
   ViewTabs.Refresh

   If Active And Not m_FolderViews.IsEmpty Then
      ViewArea.BackColor = m_Scheme.FocusTabGradient2
   Else
      If IsEditorArea Then
         If m_FolderViews.IsEmpty Then
            ViewArea.BackColor = m_Scheme.EditorAreaBackColor
         Else
            ViewArea.BackColor = vbButtonFace
         End If
      Else
         ViewArea.BackColor = m_Scheme.BackColor
      End If
   End If
   
   Set l_View = Nothing
  
End Sub

Private Sub ViewArea_Resize()

   On Error Resume Next
   
   Dim l_View As View
   Dim l_Margin As Long
   Dim i As Long

   l_Margin = m_Scheme.FrameWidth
   
   If Not m_FolderViews.IsEmpty Then
      For i = 0 To m_FolderViews.Count
      
         Set l_View = m_FolderViews.Item(i)

         With l_View.View
            If .Visible Then
               .Move 0 + l_Margin, 0 + l_Margin, ViewArea.Width - (l_Margin * 2), ViewArea.Height - (l_Margin * 2)
               Exit For
            End If
         End With

      Next i
   End If
   
   Set l_View = Nothing
   
End Sub

Private Sub ViewCaption_Click()
   ViewTabs_Click ViewTabs.ActiveViewId
End Sub

Private Sub ViewCaption_DblClick()
   On Error Resume Next
   If Not IsFloating Then
      Parent.EventRaise "FloatingWindow", FolderId, ActiveViewId
   End If
End Sub

Private Sub ViewCaption_MaximizeView()
   On Error Resume Next
   Parent.EventRaise "ActivateView", FolderId, ActiveViewId
   Parent.EventRaise "MaximizeView", FolderId, ActiveViewId
End Sub

Private Sub ViewCaption_CloseView()
   On Error Resume Next
   Parent.EventRaise "CloseView", FolderId, ActiveViewId
End Sub

Private Sub ViewCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next
   If Button = vbLeftButton Then
      Parent.EventRaise "StartDrag", FolderId, ActiveViewId
   End If
End Sub

Private Sub ViewCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next
   If Button = vbLeftButton Then
      Parent.EventRaise "EndDrag", FolderId, ActiveViewId
   End If
End Sub

Private Sub ViewTabs_Click(ByVal ViewId As String)
   On Error Resume Next
   Parent.EventRaise "ActivateView", FolderId, ViewId
End Sub

Private Sub ViewTabs_DblClick(ByVal ViewId As String)
   On Error Resume Next
   If Not ShowCaption Then
      Parent.EventRaise "MaximizeView", FolderId, ViewId
   End If
End Sub

Private Sub ViewTabs_DragStart(ByVal ViewId As String)
   On Error Resume Next
   If Not ShowCaption And Not IsEditorArea Then
      Parent.EventRaise "StartDrag", FolderId, ActiveViewId
   End If
End Sub

Private Sub ViewTabs_DragEnd(ByVal ViewId As String)
   On Error Resume Next
   If Not ShowCaption And Not IsEditorArea Then
      Parent.EventRaise "EndDrag", FolderId, ActiveViewId
   End If
End Sub

Private Sub ViewTabs_RemoveTab(ByVal ViewId As String)
   Parent.EventRaise "CloseView", FolderId, ActiveViewId
End Sub

Private Sub ViewTabs_NextTab(ByVal ViewId As String)
   Parent.EventRaise "NextView", FolderId, ActiveViewId
End Sub

Private Sub ViewTabs_PrevTab(ByVal ViewId As String)
   Parent.EventRaise "PrevView", FolderId, ActiveViewId
End Sub

