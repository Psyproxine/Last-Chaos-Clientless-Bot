VERSION 5.00
Begin VB.UserControl ucTabStrip 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F5F9FA&
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   ScaleHeight     =   585
   ScaleWidth      =   6750
   ToolboxBitmap   =   "ucTabStrip.ctx":0000
   Begin VB.PictureBox picButtons 
      BackColor       =   &H00F5F9FA&
      BorderStyle     =   0  'Kein
      Height          =   290
      Left            =   5280
      ScaleHeight     =   285
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   120
      Width           =   495
      Begin Aggressive.ucButton btnListViews 
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Picture         =   "ucTabStrip.ctx":0312
         Image           =   "VIEW_LIST"
      End
      Begin Aggressive.ucButton btnCloseEditor 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Picture         =   "ucTabStrip.ctx":062C
         Object.ToolTipText     =   "Close"
         Image           =   "VIEW_CLOSE"
      End
   End
End
Attribute VB_Name = "ucTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum VbOrientation
   VbOrientationTop = 0
   VbOrientationBottom = 1
End Enum

Private m_ActiveViewId As String
Private m_Tabs As List
Private m_Orientation As VbOrientation
Private m_CloseEditor As Boolean
Private m_NextEditor As Boolean
Private m_PrevEditor As Boolean
Private m_Active As Boolean

Public Event Click(ByVal ViewId As String)
Public Event DblClick(ByVal ViewId As String)
Public Event DragStart(ByVal ViewId As String)
Public Event DragEnd(ByVal ViewId As String)
Public Event RemoveTab(ByVal ViewId As String)
Public Event PrevTab(ByVal ViewId As String)
Public Event NextTab(ByVal ViewId As String)
'Standard-Eigenschaftswerte:
Const m_def_AutoHideButtons As Boolean = True
'Eigenschaftsvariablen:
Dim m_AutoHideButtons As Boolean

Private WithEvents m_PopupMenu As PopupMenu
Attribute m_PopupMenu.VB_VarHelpID = -1

Private Sub btnListViews_Click()
   
   Dim l_ucTab As ucTab
   Dim i As Long
   
   Set m_PopupMenu = New PopupMenu
   
   With m_Tabs
      For i = 0 To .Count
         Set l_ucTab = .Item(i)
         m_PopupMenu.AddMenuItem l_ucTab.Caption, l_ucTab.ViewId
      Next i
   End With
   
   m_PopupMenu.PopupMenu btnListViews.hWnd
   
End Sub

Private Sub m_PopupMenu_MenuItemClicked(ByVal Key As String)
   
   EventRaise "Click", Key
   
End Sub

Private Sub UserControl_Initialize()
   Set m_Tabs = New List
End Sub

Private Sub UserControl_Terminate()
   Set m_Tabs = Nothing
End Sub

Private Sub UserControl_Resize()
   On Error Resume Next
   
   Dim l_Color As OLE_COLOR
   Dim l_TabStripWidth As Long
   Dim l_Ctrl As Variant
   
   UserControl.Cls
   
   If Orientation = VbOrientationTop Then
      UserControl.Line (1, ScaleHeight - 10)-(ScaleWidth, ScaleHeight - 10), m_Scheme.FrameColor
   Else
      UserControl.Line (1, 1)-(ScaleWidth, 1), m_Scheme.FrameColor
   End If
   
   For Each l_Ctrl In Controls
      If StrComp(TypeName(l_Ctrl), "ucTab") = 0 Then
         l_TabStripWidth = l_TabStripWidth + l_Ctrl.Width
      End If
   Next
   
   If Not AutoHideButtons Or l_TabStripWidth > ScaleWidth Then
      With picButtons
         .Visible = True
         picButtons.Move ScaleWidth - .ScaleWidth, (ScaleHeight - .Height) * 0.5, .ScaleWidth, .ScaleHeight
      End With
   Else
      picButtons.Visible = False
   End If
      
End Sub

Public Property Get Active() As Boolean
   Active = m_Active
End Property

Public Property Let Active(ByVal NewActive As Boolean)
   
   If m_Active <> NewActive Then
      
      m_Active = NewActive
      Refresh
   End If
   
End Property
Private Sub picButtons_Resize()
   btnCloseEditor.Move btnCloseEditor.Left, (ScaleHeight - btnCloseEditor.Height) * 0.4
   btnListViews.Move btnListViews.Left, (ScaleHeight - btnListViews.Height) * 0.4
  ' btnNextEditor.Move btnNextEditor.left, (ScaleHeight - btnNextEditor.Height) * 0.5
  ' btnPrevEditor.Move btnPrevEditor.left, (ScaleHeight - btnPrevEditor.Height) * 0.5
End Sub

' Raises an event by its name. (ATTENSION: This public method is only for internal use!)
'
' @EventName The name of the event to raise.
' @ViewId The views id of the tab, which is calling this method.
Public Sub EventRaise(ByVal EventName As String, ByVal ViewId As String)
        
   On Error Resume Next
   
   Select Case EventName
      Case "DragStart":     RaiseEvent DragStart(ViewId)
      Case "DragEnd":       RaiseEvent DragEnd(ViewId)
      Case "Click":         RaiseEvent Click(ViewId)
                            Show ViewId
                            Refresh
      Case "DblClick":      RaiseEvent DblClick(ViewId)
   
   End Select
   
   Parent.Active = True
   
End Sub

' Returns the view id of the active view.
'
' @ActiveViewId The view id of the active view.
Public Property Get ActiveViewId() As String
   
   If m_Tabs.Count > -1 Then
      ActiveViewId = m_ActiveViewId
   Else
      ActiveViewId = vbNullString
   End If
   
End Property

' Returns the orientation of the tabs.
' Available orientation constants are VbOrientationTop (0) and VbOrientationBottom (1).
'
' @Orientation The current orientation of the tabs.
Public Property Get Orientation() As VbOrientation
   Orientation = m_Orientation
End Property

' Sets the orientation of the tabs.
' Available orientation constants are VbOrientationTop (0) and VbOrientationBottom (1).
'
' @NewOrientation The new orientation of the tabs.
Public Property Let Orientation(ByVal NewOrientation As VbOrientation)
   m_Orientation = NewOrientation
End Property

Public Property Get Tabs() As List
   Set Tabs = m_Tabs
End Property

' Adds a new tab for a view id.
'
' @ViewId Unique view id.
' @Caption The caption for the new tab.
' @Icon The icon for the new tab.
Public Sub Add(ByVal ViewId As String, ByVal Caption As String, Optional ByVal Icon As Picture, Optional ByVal ToolTipText As String)
   
   Dim l_Tab As Object
   
   If m_Tabs.Contains(ViewId) Then
      Err.Raise 1000, , "Key is not unique!"
   End If
   
   Set l_Tab = Controls.Add("Aggressive.ucTab", "Tab_" & ViewId)
   
   With l_Tab
      .Orientation = Me.Orientation
      .Caption = Caption
      .ToolTip = ToolTipText
      .ViewId = ViewId
      .Visible = True
      .Icon = Icon
   End With
      
   picButtons.ZOrder
   m_Tabs.Add ViewId, l_Tab
      
   m_ActiveViewId = ViewId
   Refresh
   RefreshButtons

End Sub

' Removes the tab for a view id.
'
' @ViewId Unique view id.
Public Sub Remove(ByVal ViewId As String)
   
   Dim l_Tab As Object
   
   If m_Tabs.Contains(ViewId) Then
   
      Controls.Remove "Tab_" & ViewId
   
      m_Tabs.Remove ViewId
      
      If Not m_Tabs.IsEmpty Then
         m_ActiveViewId = m_Tabs.Item(0).ViewId
      Else
         m_ActiveViewId = vbNullString
      End If
      
      Refresh
      RefreshButtons
      
   End If

End Sub

' Returns the count of tabs.
'
' @Count The count of tabs.
Public Function Count() As Long
   Count = m_Tabs.Count
End Function

' Shows (activates) the tab for a view id.
'
' @ViewId The unique view id.
Public Sub Show(ByVal ViewId As String)
   
   
   m_ActiveViewId = ViewId
   
   Refresh
   
   RefreshButtons
   
End Sub

' Refreshs all tabs on the TabStrip.
Public Sub Refresh()
   
   Dim l_ActiveViewFound As Boolean
   Dim l_doRecalcIndent As Boolean
'   Dim l_Pos As Long
   Dim l_Tab As Variant
   
   UserControl.BackColor = m_Scheme.BackColor
   picButtons.BackColor = m_Scheme.BackColor
   btnListViews.BackColor = m_Scheme.BackColor
   btnCloseEditor.BackColor = m_Scheme.BackColor
'   btnNextEditor.BackColor = m_Scheme.BackColor
'   btnPrevEditor.BackColor = m_Scheme.BackColor
   btnListViews.Refresh
       
   If Not m_Tabs.IsEmpty Then

      For Each l_Tab In m_Tabs.Items
         With l_Tab

            If StrComp(.ViewId, ActiveViewId, vbBinaryCompare) = 0 Then
               l_ActiveViewFound = True
               
               If Active Then
                  .State = STATE_FOCUS
               Else
                  .State = STATE_ACTIVE
               End If
               .Refresh
               
               ' If active tag is not visible -> call rearrange tabs
               If l_Tab.Left < 0 Or _
                  l_Tab.Left + l_Tab.Width > ScaleWidth - picButtons.Width Then
                  l_doRecalcIndent = True
               End If
            Else
               .State = STATE_INACTIVE
               .Refresh
            End If
        
         End With
      Next

      Rearrange ' l_doRecalcIndent
      
   End If
   
   UserControl_Resize
   
End Sub

Private Sub Rearrange(Optional ByVal RecalcIndent As Boolean = True)

   Dim l_Indent As Long
   Dim l_ActiveTab As Variant
   Dim l_ActiveTabWidth As Long
   Dim l_Tab As Variant
   
   If Not m_Tabs.IsEmpty Then
   
      ' Calculate indent
      If Len(ActiveViewId) > 0 Then
         
         For Each l_Tab In m_Tabs.Items
            With l_Tab
                                    
               If StrComp(.ViewId, ActiveViewId, vbBinaryCompare) = 0 Then
                  Set l_ActiveTab = l_Tab
                  Exit For
               Else
                  l_Indent = l_Indent + .Width + 20
               End If
              
            End With
         Next
      
      End If
      
      If RecalcIndent Then
      If Not IsEmpty(l_ActiveTab) Then
         l_ActiveTabWidth = l_ActiveTab.Width + 120
      End If
      If picButtons.Visible Then
         l_ActiveTabWidth = l_ActiveTabWidth + picButtons.Width
      End If
      
      If l_Indent > ScaleWidth - l_ActiveTabWidth Then
         l_Indent = (l_Indent - ScaleWidth) + l_ActiveTabWidth
      Else
          l_Indent = 0
      End If
      l_Indent = l_Indent * (-1)
      l_Indent = l_Indent + 40
      Else
      
         Set l_Tab = m_Tabs.Item(0)
         
      l_Indent = l_Tab.Left
      'l_Indent = l_Indent + 60
      
      End If
      
      
      For Each l_Tab In m_Tabs.Items
         With l_Tab
                                    
            Dim h As Long
            
            h = ScaleHeight
            If StrComp(.ViewId, ActiveViewId, vbBinaryCompare) = 0 Then
               h = h + 10
            End If
                                    
            If Orientation = VbOrientationTop Then
               .Move l_Indent, 30, .Width, h - 20
            Else
               .Move l_Indent, 1, .Width, h - 20
            End If
              
            l_Indent = l_Indent + .Width + 20
         End With
      Next

   End If

End Sub

Public Property Let CloseButton(ByVal Visible As Boolean)
   btnCloseEditor.Visible = Visible
   btnCloseEditor.Refresh
   m_CloseEditor = Visible
   ResizeButtonBar
End Property

Public Property Let NextButton(ByVal Visible As Boolean)
'   btnNextEditor.Visible = Visible
'   btnNextEditor.Refresh
   m_NextEditor = Visible
   ResizeButtonBar
End Property
Public Property Let PrevButton(ByVal Visible As Boolean)
'   btnPrevEditor.Visible = Visible
'   btnPrevEditor.Refresh
   m_PrevEditor = Visible
   ResizeButtonBar
End Property
Private Sub ResizeButtonBar()
   
   Dim w As Long
   
'   With btnPrevEditor
'      If m_PrevEditor Then
'         .left = w + 1
'         w = w + .Width
'      End If
'   End With
   
'   With btnNextEditor
'      If m_NextEditor Then
'         .left = w + 1
'         w = w + .Width
'      End If
'   End With
   
   With btnListViews
      'If m_CloseEditor Then
         .Left = w + 1
         w = w + .Width
      'End If
   End With
   
   With btnCloseEditor
      If m_CloseEditor Then
         .Left = w + 1
         w = w + .Width
      End If
   End With

    'w = w + btnListViews.Width
   
   picButtons.Width = w
   
End Sub

Private Sub btnCloseEditor_Click()
   RaiseEvent RemoveTab(m_ActiveViewId)
   RefreshButtons
End Sub
Private Sub btnNextEditor_Click()
   RaiseEvent NextTab(m_ActiveViewId)
   RefreshButtons
End Sub
Private Sub btnPrevEditor_Click()
   RaiseEvent PrevTab(m_ActiveViewId)
   RefreshButtons
End Sub

Public Sub RefreshButtons()
   
   On Error Resume Next
   
   Dim l_Tab As Variant
   
   Set l_Tab = m_Tabs.Item(0)
'   btnPrevEditor.Enabled = Not (StrComp("Tab_" & ActiveViewId, l_Tab.Name) = 0)
   
   Set l_Tab = m_Tabs.Item(m_Tabs.Count)
 '  btnNextEditor.Enabled = Not (StrComp("Tab_" & ActiveViewId, l_Tab.Name) = 0)
   
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MemberInfo=0,0,0,1
Public Property Get AutoHideButtons() As Boolean
    AutoHideButtons = m_AutoHideButtons
End Property

Public Property Let AutoHideButtons(ByVal New_AutoHideButtons As Boolean)
    m_AutoHideButtons = New_AutoHideButtons
    PropertyChanged "AutoHideButtons"
End Property

'Eigenschaften für Benutzersteuerelement initialisieren
Private Sub UserControl_InitProperties()
    m_AutoHideButtons = m_def_AutoHideButtons
End Sub

'Eigenschaftenwerte vom Speicher laden
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_AutoHideButtons = PropBag.ReadProperty("AutoHideButtons", m_def_AutoHideButtons)
End Sub

'Eigenschaftenwerte in den Speicher schreiben
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AutoHideButtons", m_AutoHideButtons, m_def_AutoHideButtons)
End Sub

