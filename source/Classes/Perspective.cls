VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Perspective"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' <B>Description</B><BR />
' <BR />
' A perspective instance describs the presentation of a VbEclipse
' perspective. It contains an editor area and a list of one or more
' folders.<BR />
' <BR />
' Each perspective needs an unique perspective id to show or close
' the perspective. You can add folders by calling the AddFolder() function.<BR /><BR />
' <B>Terms of Use</B><BR />
' <BR />
' You may freely distribute the VbEclipse project and the
' test sample project. You're free to use the VbEclipse control
' in your applications and distribute it with the source code
' of your applications.<BR />
' <BR />
' However, you may not claim that you have written the VbEclipse
' or the Test project.<BR />
' <BR />
' For questions, contact me at vbeclipseproject ab-software.com<BR />
' <BR />
' For more information about the VbEclipse Control, visit the
' following URL: http://www.ab-software.com<BR />
' <BR />
' @See http://www.ab-software.com
'
'
' © Copyright 2004-2006, AB-Software

Option Explicit

' The relation constanst of the vbRelationship enum should be used to
' describe the perspective layout. The relation sets the position of a
' folder view on dependence of another folder (by refId). If you define
' that folder1 should have a vbRelLeft relationship to folder2, then
' folder1 will be displayed on the left side of folder2.
'
Public Enum vbRelationship
   vbRelLeft = 1         ' Describes the relation of a folder view as left of another folder.
   vbRelRight = 2        ' Describes the relation of a folder view as right of another folder.
   vbRelTop = 3          ' Describes the relation of a folder view as top of another folder.
   vbRelBottom = 4       ' Describes the relation of a folder view as bottom of another folder.
   vbRelFloating = 98    ' Describes the relation of a folder view as a floating window.
End Enum

Private m_PerspectiveId As String
Private m_EditorAreaVisible As Boolean  ' <b>True</b> to show the editor area; false otherwise.
Private m_Folders As List               ' This list stores all folder definitions of this perspective.
Private m_ActiveViewId As String        ' Id of the active view

' Initialize the perspective class.
Private Sub Class_Initialize()
   Set m_Folders = New List
End Sub

' Terminate the perspective class.
Private Sub Class_Terminate()
   Set m_Folders = Nothing
End Sub

' The constant ID_EDITOR_AREA returns the id of the edtior area folder. You can make
' relations to that folder if you use the ID_EDITOR_AREA constant as refId.
'
' @return
'        Returns the folder id of the editor area.
Public Function ID_EDITOR_AREA() As String
   ID_EDITOR_AREA = "com_ab_software_ui_editorarea"
End Function

' This property returns the perspective id of this perspective. The perspective id
' is used to show or hide a perspective.
'
' @param PerspectiveId
'        The perspective id.
Public Property Get PerspectiveId() As String
   PerspectiveId = m_PerspectiveId
End Property

' Sets the the perspective id of a new perspective. The perspective id is used to show
' or hide a perspective.
'
' @param NewPerspectiveId
'        The perspective id of the new perspective.
Public Property Let PerspectiveId(ByVal NewPerspectiveId As String)
   m_PerspectiveId = NewPerspectiveId
End Property

' This property returns if the perspective should have an editor area. If there is no
' use to open an editor you can set this property to false.
'
' @return
'         Return if the perspective should have an editor area.
Public Property Get EditorAreaVisible() As Boolean
   EditorAreaVisible = m_EditorAreaVisible
End Property
' This property defines if the perspective should have an editor area. If you don't
' need to open an editor you can set this property to false.
'
' @param EditorAreaVisible
'        Set true to define that the perspective should have an editor area; false otherwise.
Public Property Let EditorAreaVisible(ByVal NewEditorAreaVisible As Boolean)
   m_EditorAreaVisible = NewEditorAreaVisible
End Property

' Returns the id of the perspectives active view.
'
' @return
'        Id of the perspectives active view.
Public Property Get ActiveViewId() As String
   
   Dim i As Integer
   Dim l_Folder As Folder
   
   If Len(m_ActiveViewId) > 0 Then
       
      For i = 0 To Folders.Count
         Set l_Folder = Folders.Item(i)
         
         If l_Folder.Views.Contains(m_ActiveViewId) Then
            ActiveViewId = m_ActiveViewId
            Exit Property
         End If
      Next i
   End If
      
   If Not Folders.IsEmpty Then
      Set l_Folder = Folders.Item(0)
      If Not l_Folder Is Nothing And _
         Not l_Folder.Views.IsEmpty Then
         
         ActiveViewId = l_Folder.Views.Item(0)
      End If
   End If
   
End Property

' Sets the id of the perspectives acitve view.
'
' @param New_ActiveViewId The id of the perspectives active view.
Public Property Let ActiveViewId(ByVal New_ActiveViewId As String)
   m_ActiveViewId = New_ActiveViewId
End Property

' All folders you add with the AddFolder() function are stored in a list. You
' get this list if you use the Folders property.
'
' @return Folders List of all added folders.
Public Property Get Folders() As List
   Set Folders = m_Folders
End Property

' This function creates a new folder. A folder represents a container for a
' view stack. It provides different GUI elements like tabs to enable the user
' to work with stacked view. Each view has to be created on a folder, you can't
' create a view without.
'
' @param FolderId
'        A unique folder id to make other relations to this folder.
' @param Relationship
'        The relationship to another folder (left, right, bottom, top) or floating window.
' @param Ratio
'        The ratio to another folder (a numeric range between 0.05 and 0.95).
' @param RefId
'        The reference id to layout the new folder by a relation to another folder.
'
' @return
'        AddFolder Returns a new folder instance of the added folder.
Public Function AddFolder(ByVal FolderId As String, _
                          ByVal Relationship As vbRelationship, _
                          ByVal Ratio As Double, _
                          ByVal RefId As String) As Folder
                             
   CheckId FolderId
                             
   On Error GoTo ErrorHandle
   
   ' Create a new folder.
   Set AddFolder = New Folder
       
   ' Set folder properties.
   With AddFolder
      .FolderId = FolderId
      .Ratio = Ratio
      .RefId = RefId
      .Relationship = Relationship
   End With
   
   ' Add folder to the folder list (see Folders() property).
   m_Folders.Add FolderId, AddFolder
   
ErrorHandle:

End Function

' Adds a placeholder for a view.
'
' @FolderId
'        A folder id.
' @ViewId
'        The id of the view to set a placeholder.
' @Index
'        An optional position of the view.
Public Sub AddPlaceholder(ByVal ViewId As String, ByVal FolderId As String)
                                                         
   On Error GoTo ErrorHandle
   
   m_Placeholders.Remove PerspectiveId & "." & ViewId
   m_Placeholders.Add PerspectiveId & "." & ViewId, FolderId
   
ErrorHandle:

End Sub
