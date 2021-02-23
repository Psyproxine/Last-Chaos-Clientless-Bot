VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' <B>Description</B><BR />
' <BR />
' This class represents the model of a view. These views are just Visual Basic
' forms named with a unique key. The properties <i>ViewId</i> and <i>View</i>
' are the only two properies needed to descripe a view.<BR />
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
' @ViewId   The views id.
' @View     The view itself.

Option Explicit

Private m_ViewId As String       ' The view id.
Private m_View As Object         ' The view reference.

' Get the view id. The view id must be a unique key within each perspective.
'
' @ViewId The current view id.
Public Property Get ViewId() As String
   ViewId = m_ViewId
End Property

' Set the view id. The view id must be a unique key within each perspective.
'
' @NewViewId The new view id.
Public Property Let ViewId(ByVal NewViewId As String)
   m_ViewId = NewViewId
End Property

' Returns the view object or nothing if there is no view available.
'
' @View The current view (a Visual Basic form).
Public Property Get View() As Object
   Set View = m_View
End Property

' Set the view object (normally a Visual Basic form) for the unique view id.
' Be sure that each view instance has its own unique view id. If there are
' diffrent view instances with different ids but the same view object, only
' one view will be displayed.
'
' @NewView The view object (normally a Visual Basic form) to set for view id.
Public Property Set View(ByRef NewView As Object)
   Set m_View = NewView
End Property
