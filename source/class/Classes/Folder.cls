VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Folder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' <B>Description</B><BR />
' <BR />
' The purpose of this class is to describe a perspective folder.
' Folders are used in perspectives to layout the presentation
' space. A folder represents a container for a stack of one or
' more views. The graphical counterpart to this folders is the
' ucFolder control, which will be initialised by a instance of
' the folder class.<BR />
' <BR />
' The class property <i>Instancing</i> is set to <i>Public Not Createable</i>
' because the endusers should not create Folder instances manually by
' <code>Dim Folder = new Folder</code>. They should use the perspective
' function <i>AddFolder()</i> to create a new folder, which will
' register the folder as a perspective folder.<BR />
' <BR />
' <B>Terms of Use</B><BR />
' <BR />
' You may freely distribute the VbEclipse project and the
' test sample project. You're free to use the VbEclipse control
' in your applications and distribute it with the source code
' of your applications.<BR />
' <BR />
' However, you may not claim that you've written the VbEclipse
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

Private m_FolderId As String        ' The folder id
Private m_RefId As String           ' The reference id
Private m_Ratio As Double           ' The ratio
Private m_Relationship As Long      ' The relationship
Private m_Views As List             ' Views of this folder
Private m_ActiveViewId As String    ' The id the active view

Private m_Position As Rectangle

' Initialize class variables
Private Sub Class_Initialize()
   Set m_Views = New List
   Set m_Position = New Rectangle
End Sub

' Terminate class variables
Private Sub Class_Terminate()
   Set m_Views = Nothing
   Set m_Position = Nothing
End Sub

' Returns the id of this folder.
'
' @FolderId The folder id.
Public Property Get FolderId() As String
   FolderId = m_FolderId
End Property

' Sets the id for this folder.
'
' @New_FolderId The new folder id.
Public Property Let FolderId(ByVal New_FolderId As String)
   m_FolderId = New_FolderId
End Property

' Returns the reference id of this folder.
'
' @RefId The reference id.
Public Property Get RefId() As String
   RefId = m_RefId
End Property

' Sets the reference id for this folder.
'
' @NewRefId The new reference id.
Public Property Let RefId(ByVal New_RefId As String)
   m_RefId = New_RefId
End Property

' Returns the ratio.
'
' @Ratio
Public Property Get Ratio() As Double
   Ratio = m_Ratio
End Property

' Sets the ratio.
'
' @New_Ratio
Public Property Let Ratio(ByVal New_Ratio As Double)
   m_Ratio = New_Ratio
End Property

' Returns the relationship.
'
' @Relationship The relationship.
Public Property Get Relationship() As Long
   Relationship = m_Relationship
End Property

' Sets the relationship.
'
' @New_Relationship
Public Property Let Relationship(ByVal New_Relationship As Long)
   m_Relationship = New_Relationship
End Property

Public Property Get Views() As List
   Set Views = m_Views
End Property

' Returns the id of the folders active view.
'
' @ActiveViewId Id of the folders active view.
Public Property Get ActiveViewId() As String
   
   If Len(m_ActiveViewId) > 0 Then
   
      ActiveViewId = m_ActiveViewId
      
      If Not Views.Contains(m_ActiveViewId) Then
         m_ActiveViewId = vbNullString
      End If
      
   End If
   
   If Len(m_ActiveViewId) = 0 Then
      If Not Views.IsEmpty Then
         m_ActiveViewId = Views.Item(0)
      End If
   End If
   
   ActiveViewId = m_ActiveViewId
   
End Property

' Sets the id of the folders acitve view.
'
' @New_ActiveViewId The id of the folders active view.
Public Property Let ActiveViewId(ByVal New_ActiveViewId As String)
   m_ActiveViewId = New_ActiveViewId
End Property

' Returns the left position of a floating folder.
'
' Position The position of the folder.
Property Get Position() As Rectangle
   Set Position = m_Position
End Property

' Adds a view to this folder.
'
' @ViewId The id of the view which should be added.
' @Index An optional position of the view.
Public Sub AddView(ByVal ViewId As String, Optional ByVal Index As Long = -1)
   
   CheckId ViewId
   
   If Index > 0 Then
      Index = Index - 1
   End If
   
   m_Views.Add ViewId, ViewId, Index
   
End Sub
